"""
Core validation logic — identical remark rules across every marketplace,
plus optional Product Master cross-checks.
"""

import numpy as np
import pandas as pd

from . import config
from .readers import clean_sku


def get_remark(expected, mp_stock, not_found_label: str) -> tuple[str, str]:
    if pd.isna(mp_stock):
        return "Mismatch", not_found_label
    exp, mp = int(expected), int(mp_stock)
    if exp == 0 and mp > 0:
        return "Mismatch", "UPDATE 0"
    if exp > 0 and mp == 0:
        return "Mismatch", "UPDATE STOCK"
    if exp != mp:
        return "Mismatch", "Stock Mismatch"
    return "Match", "TRUE"


def process_marketplace(
    sv_df: pd.DataFrame,
    lookup: dict,
    not_found_label: str,
    status_lookup: dict | None = None,
    master_skus: set | None = None,
) -> pd.DataFrame:
    """
    sv_df: StockValidation dataframe with 'Seller SKU' and 'Expected Stock' columns.
    lookup: {sku: quantity} from the marketplace's own stock export.
    status_lookup: optional {sku: status} (currently used for Zalora).
    master_skus: optional set of SKUs from Product Master, to flag SKUs that
        are live on the marketplace but absent from the master catalogue.
    """
    df = sv_df.copy()
    df["Seller SKU"] = clean_sku(df["Seller SKU"])
    if "Difference" in df.columns:
        df.rename(columns={"Difference": "Original Difference"}, inplace=True)
    df["Expected Stock"] = pd.to_numeric(df["Expected Stock"], errors="coerce").fillna(0).astype(int)

    df["Marketplace Stock"] = df["Seller SKU"].map(lookup)

    statuses, remarks = [], []
    for exp, mp in zip(df["Expected Stock"], df["Marketplace Stock"]):
        status, remark = get_remark(exp, mp, not_found_label)
        statuses.append(status)
        remarks.append(remark)
    df["Status_Result"] = statuses
    df["Remark"] = remarks

    df["Difference"] = df["Marketplace Stock"] - df["Expected Stock"]
    df.loc[df["Marketplace Stock"].isna(), "Difference"] = np.nan

    if status_lookup:
        df["Zalora_Status"] = df["Seller SKU"].map(status_lookup).fillna("—")

    if master_skus is not None:
        df["In Product Master"] = df["Seller SKU"].isin(master_skus)

    return df


def compute_kpis(df: pd.DataFrame, not_found_label: str) -> dict:
    total = len(df)
    matched = int((df["Remark"] == "TRUE").sum())
    stock_mismatch = int((df["Remark"] == "Stock Mismatch").sum())
    update0 = int((df["Remark"] == "UPDATE 0").sum())
    update_stock = int((df["Remark"] == "UPDATE STOCK").sum())
    not_found = int((df["Remark"] == not_found_label).sum())
    total_issues = total - matched
    accuracy = round((matched / total * 100), 2) if total else 0
    result = {
        "Total SKUs": total, "Matched": matched, "Stock Mismatch": stock_mismatch,
        "UPDATE 0": update0, "UPDATE STOCK": update_stock, "NOT FOUND": not_found,
        "Total Issues": total_issues, "Accuracy %": accuracy,
    }
    if "In Product Master" in df.columns:
        result["Not In Product Master"] = int((~df["In Product Master"]).sum())
    return result


def master_coverage_report(master_skus: set, marketplace_lookups: dict) -> pd.DataFrame:
    """
    Given the Product Master SKU set and a dict of {marketplace: {sku: qty}}
    lookups, build a coverage matrix: for every master SKU, which
    marketplaces actually carry it (and at what quantity)?

    marketplace_lookups: {"Lazada": {sku: qty}, "Shopee": {...}, ...}
    """
    rows = []
    for sku in sorted(master_skus):
        row = {"SKU": sku}
        listed_anywhere = False
        for mp_name, lookup in marketplace_lookups.items():
            qty = lookup.get(sku)
            row[f"{mp_name} Stock"] = qty if qty is not None else None
            row[f"Listed on {mp_name}"] = qty is not None
            listed_anywhere = listed_anywhere or (qty is not None)
        row["Listed Anywhere"] = listed_anywhere
        rows.append(row)
    return pd.DataFrame(rows)


def soh_vs_all(soh_df: pd.DataFrame, all_lookup: dict | None) -> pd.DataFrame | None:
    """Optional warehouse cross-check. Only meaningful if the ALL report was
    identified as true stock data (see readers.read_all_report)."""
    if soh_df is None or soh_df.empty or not all_lookup:
        return None

    df = soh_df.copy()
    df["ALL_Quantity"] = df["SKU"].map(all_lookup)

    def remark(row):
        if pd.isna(row["ALL_Quantity"]):
            return "NOT FOUND"
        if int(row["SOH_Quantity"]) != int(row["ALL_Quantity"]):
            return "Mismatch"
        return "TRUE"

    df["Remark"] = df.apply(remark, axis=1)
    return df
