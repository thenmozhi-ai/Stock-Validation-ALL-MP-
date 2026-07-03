"""
One reader function per file role. Each returns data in a normalized shape
so validate.py doesn't need to know about file-format quirks.

Stock lookups are returned as {sku: quantity} dicts (duplicates pre-summed).
StockValidation and Product Master are returned as DataFrames.
"""

import io
import zipfile
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd


def clean_sku(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


# ---------------------------------------------------------------------------
# SHOPEE — activePane XML bug patch
# ---------------------------------------------------------------------------

def fix_shopee_xlsx(file_bytes: bytes) -> io.BytesIO:
    """Shopee's Mass Update export has an invalid activePane attribute that
    crashes openpyxl. Patch it in-memory before reading."""
    src = io.BytesIO(file_bytes)
    dst = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") or item.filename.endswith(".rels"):
                text = data.decode("utf-8", errors="ignore")
                for old, new in [
                    ("bottom_left", "bottomLeft"),
                    ("top_left", "topLeft"),
                    ("bottom_right", "bottomRight"),
                    ("top_right", "topRight"),
                ]:
                    text = text.replace(f'activePane="{old}"', f'activePane="{new}"')
                    text = text.replace(f'pane="{old}"', f'pane="{new}"')
                data = text.encode("utf-8")
            zout.writestr(item, data)
    dst.seek(0)
    return dst


def read_shopee_zip_members(zip_bytes: bytes) -> list[io.BytesIO]:
    """Shopee's mass_update_sales_info sometimes arrives as a .zip containing
    one or more .xlsx files. Extract them all as in-memory buffers."""
    members = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for inner_name in zf.namelist():
            if inner_name.lower().endswith(".xlsx"):
                members.append(io.BytesIO(zf.read(inner_name)))
    return members


# ---------------------------------------------------------------------------
# SOH — XML disguised as .xls
# ---------------------------------------------------------------------------

def read_soh(file_bytes: bytes) -> pd.DataFrame:
    """SOHbySKU...xls is actually a SpreadsheetML XML file (NetSuite export).
    SKU sits at column index 6, quantity at column index 14; data rows start
    at row index 7."""
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.fromstring(file_bytes)
    ws = root.find("ss:Worksheet", ns)
    table = ws.find("ss:Table", ns)
    rows = table.findall("ss:Row", ns)

    def row_to_list(row):
        cells = row.findall("ss:Cell", ns)
        vals = []
        for c in cells:
            d = c.find("ss:Data", ns)
            vals.append(d.text if d is not None else None)
        return vals

    records = []
    for row in rows[7:]:
        vals = row_to_list(row)
        if len(vals) > 14 and vals[6]:
            records.append({"SKU": str(vals[6]).strip(), "SOH_Quantity": vals[14]})

    df = pd.DataFrame(records)
    if not df.empty:
        df["SOH_Quantity"] = pd.to_numeric(df["SOH_Quantity"], errors="coerce").fillna(0).astype(int)
        df["SKU"] = clean_sku(df["SKU"])
        df = df.groupby("SKU", as_index=False)["SOH_Quantity"].sum()
    return df


# ---------------------------------------------------------------------------
# ALL REPORT — flexible: could be true warehouse stock, or order-transaction
# export that merely shares the "ALL-DD-Mon-YYYY.csv" filename convention.
# ---------------------------------------------------------------------------

STOCK_LIKE_COLS = {"sellersku", "seller_sku", "sku", "quantity", "stock"}
ORDER_LIKE_COLS = {"order_id", "order_number", "order_item_id", "tracking_number"}


def read_all_report(file_bytes: bytes) -> dict:
    """
    Returns:
    {
        "kind": "stock" | "orders" | "unknown",
        "df": <raw DataFrame>,
        "lookup": {sku: qty} or None,   # only populated when kind == "stock"
        "note": human-readable caveat string,
    }
    """
    df = pd.read_csv(io.BytesIO(file_bytes), low_memory=False)
    cols_lower = {c.lower().strip() for c in df.columns}

    is_order_like = bool(cols_lower & ORDER_LIKE_COLS)
    sku_col = next((c for c in df.columns if c.lower().strip() in ("sellersku", "seller_sku", "sku")), None)
    qty_col = next((c for c in df.columns if c.lower().strip() in ("quantity", "stock")), None)

    if is_order_like:
        return {
            "kind": "orders",
            "df": df,
            "lookup": None,
            "note": (
                "This ALL report looks like order-transaction data (has order_id / tracking "
                "columns), not a warehouse stock snapshot. Skipping it from stock comparisons — "
                "verify this is the right file if you expected a stock report."
            ),
        }

    if sku_col and qty_col:
        clean = df[[sku_col, qty_col]].copy()
        clean.columns = ["SKU", "Quantity"]
        clean["SKU"] = clean_sku(clean["SKU"])
        clean["Quantity"] = pd.to_numeric(clean["Quantity"], errors="coerce").fillna(0).astype(int)
        lookup = clean.groupby("SKU")["Quantity"].sum().to_dict()
        return {
            "kind": "stock",
            "df": df,
            "lookup": lookup,
            "note": f"Read as stock data using columns '{sku_col}' / '{qty_col}'.",
        }

    return {
        "kind": "unknown",
        "df": df,
        "lookup": None,
        "note": "Couldn't identify SKU/quantity columns in this ALL report — skipped.",
    }


# ---------------------------------------------------------------------------
# PRODUCT MASTER — flexible column detection since layouts vary by company.
# ---------------------------------------------------------------------------

def read_product_master(file_bytes: bytes, filename: str) -> dict:
    """
    Best-effort reader: looks for a SKU-like column and, optionally, a
    status/active column. Returns:
    {
        "df": <raw DataFrame>,
        "sku_col": str | None,
        "status_col": str | None,
        "skus": set of cleaned SKUs,
        "status_lookup": {sku: status} or {},
    }
    """
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))

    sku_col = next(
        (c for c in df.columns if "sku" in c.lower() and "parent" not in c.lower()),
        None,
    )
    status_col = next(
        (c for c in df.columns if any(k in c.lower() for k in ("status", "active", "lifecycle"))),
        None,
    )

    skus = set()
    status_lookup = {}
    if sku_col:
        clean = clean_sku(df[sku_col].dropna())
        skus = set(clean)
        if status_col:
            tmp = df[[sku_col, status_col]].dropna(subset=[sku_col])
            tmp[sku_col] = clean_sku(tmp[sku_col])
            status_lookup = tmp.set_index(sku_col)[status_col].astype(str).str.strip().to_dict()

    return {
        "df": df,
        "sku_col": sku_col,
        "status_col": status_col,
        "skus": skus,
        "status_lookup": status_lookup,
    }


# ---------------------------------------------------------------------------
# MARKETPLACE STOCK LOOKUPS
# ---------------------------------------------------------------------------

def build_lazada_lookup(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=0, skiprows=[1, 2, 3])
    expected_cols = [
        "Product ID", "catId", "Product Name", "currencyCode", "sku.skuId", "status",
        "Shop SKU", "SellerSKU", "Quantity", "Price", "SpecialPrice",
        "SpecialPrice Start", "SpecialPrice End", "Variations Combo", "md5key",
    ]
    df.columns = expected_cols[: len(df.columns)]
    df["SellerSKU"] = clean_sku(df["SellerSKU"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("SellerSKU")["Quantity"].sum().to_dict()


def build_shopee_lookup(file_buffers: list) -> dict:
    """file_buffers: list of raw bytes (already extracted from any .zip)."""
    lookup: dict = {}
    for raw in file_buffers:
        try:
            fixed = fix_shopee_xlsx(raw)
            df = pd.read_excel(fixed, header=2, skiprows=[3, 4, 5])
        except Exception:
            continue
        if "SKU" not in df.columns or "Stock" not in df.columns:
            continue
        df = df.dropna(subset=["SKU"])
        df["SKU"] = clean_sku(df["SKU"])
        df = df[df["SKU"] != "nan"]
        if df.empty:
            continue
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)
        for sku, qty in df.groupby("SKU")["Stock"].sum().items():
            lookup[sku] = lookup.get(sku, 0) + qty
    return lookup


def build_tiktok_lookup(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), header=2, skiprows=[3, 4])
    expected_cols = [
        "Product ID", "Category", "Product Name", "SKU ID",
        "Variation Option", "Price", "Quantity", "Seller SKU",
    ]
    df.columns = expected_cols[: len(df.columns)]
    df["Seller SKU"] = clean_sku(df["Seller SKU"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    return df.groupby("Seller SKU")["Quantity"].sum().to_dict()


def build_zalora_lookup(stock_bytes: bytes, status_bytes: bytes | None = None) -> tuple[dict, dict]:
    df = pd.read_excel(io.BytesIO(stock_bytes), header=0)
    df.columns = ["SellerSku", "ShopSku", "Quantity", "Name"][: len(df.columns)]
    df["SellerSku"] = clean_sku(df["SellerSku"])
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    lookup = df.groupby("SellerSku")["Quantity"].sum().to_dict()

    status_lookup = {}
    if status_bytes is not None:
        try:
            ds = pd.read_excel(io.BytesIO(status_bytes), header=0)
            ds.columns = ["SellerSku", "ShopSku", "Name", "Status"][: len(ds.columns)]
            ds["SellerSku"] = clean_sku(ds["SellerSku"])
            status_lookup = ds.set_index("SellerSku")["Status"].astype(str).str.strip().str.lower().to_dict()
        except Exception:
            pass
    return lookup, status_lookup
