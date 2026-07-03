"""
Orchestrates the full pipeline: classify uploads -> read each source ->
validate -> build workbook. This is the one place that knows about the
end-to-end flow; app.py (Streamlit) and cli.py just call run_pipeline().
"""

from . import config, detect, readers, validate, report


class UploadedBlob:
    """Minimal wrapper so both Streamlit's UploadedFile and plain
    (name, bytes) tuples can be treated uniformly."""

    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


def run_pipeline(files) -> dict:
    """
    files: iterable of file-like objects with `.name` and `.getvalue()`.

    Returns:
    {
        "marketplaces": [(name, df, not_found_label), ...],
        "kpi_results": {name: kpi_dict},
        "soh_vs_all_df": DataFrame | None,
        "master_coverage_df": DataFrame | None,
        "notes": [str, ...],
        "detected": [marketplace_name, ...],
        "unmatched": [filename, ...],
        "missing_pairs": [str, ...],   # e.g. "Lazada StockValidation found but no Stock & Price file"
    }
    """
    classified = detect.classify_uploads(files)
    notes = []
    missing_pairs = []

    marketplace_lookups = {}  # for product master coverage cross-check
    marketplaces = []
    kpi_results = {}
    detected = []

    # ---- Lazada ----
    if "lazada_stockval" in classified:
        if "lazada_stock" in classified:
            lookup = readers.build_lazada_lookup(classified["lazada_stock"].getvalue())
            marketplace_lookups["Lazada"] = lookup
            sv_df = _read_csv(classified["lazada_stockval"])
            out = validate.process_marketplace(sv_df, lookup, config.NOT_FOUND_LABEL["Lazada"])
            marketplaces.append(("Lazada", out, config.NOT_FOUND_LABEL["Lazada"]))
            kpi_results["Lazada"] = validate.compute_kpis(out, config.NOT_FOUND_LABEL["Lazada"])
            detected.append("Lazada")
        else:
            missing_pairs.append("Lazada StockValidation found but no Stock & Price (pricestock) file.")

    # ---- Shopee ----
    if "shopee_stockval" in classified:
        if "shopee_stock" in classified:
            raw_buffers = []
            for f in classified["shopee_stock"]:
                data = f.getvalue()
                if f.name.lower().endswith(".zip"):
                    raw_buffers.extend(b.getvalue() for b in readers.read_shopee_zip_members(data))
                else:
                    raw_buffers.append(data)
            lookup = readers.build_shopee_lookup(raw_buffers)
            marketplace_lookups["Shopee"] = lookup
            sv_df = _read_csv(classified["shopee_stockval"])
            out = validate.process_marketplace(sv_df, lookup, config.NOT_FOUND_LABEL["Shopee"])
            marketplaces.append(("Shopee", out, config.NOT_FOUND_LABEL["Shopee"]))
            kpi_results["Shopee"] = validate.compute_kpis(out, config.NOT_FOUND_LABEL["Shopee"])
            detected.append("Shopee")
        else:
            missing_pairs.append("Shopee StockValidation found but no Mass Update file.")

    # ---- TikTok ----
    if "tiktok_stockval" in classified:
        if "tiktok_stock" in classified:
            lookup = readers.build_tiktok_lookup(classified["tiktok_stock"].getvalue())
            marketplace_lookups["TikTok"] = lookup
            sv_df = _read_csv(classified["tiktok_stockval"])
            out = validate.process_marketplace(sv_df, lookup, config.NOT_FOUND_LABEL["TikTok"])
            marketplaces.append(("TikTok", out, config.NOT_FOUND_LABEL["TikTok"]))
            kpi_results["TikTok"] = validate.compute_kpis(out, config.NOT_FOUND_LABEL["TikTok"])
            detected.append("TikTok")
        else:
            missing_pairs.append("TikTok StockValidation found but no Batch Edit file.")

    # ---- Zalora ----
    if "zalora_stockval" in classified:
        if "zalora_stock" in classified:
            status_bytes = classified.get("zalora_status")
            lookup, status_lookup = readers.build_zalora_lookup(
                classified["zalora_stock"].getvalue(),
                status_bytes.getvalue() if status_bytes else None,
            )
            marketplace_lookups["Zalora"] = lookup
            sv_df = _read_csv(classified["zalora_stockval"])
            out = validate.process_marketplace(sv_df, lookup, config.NOT_FOUND_LABEL["Zalora"], status_lookup)
            marketplaces.append(("Zalora", out, config.NOT_FOUND_LABEL["Zalora"]))
            kpi_results["Zalora"] = validate.compute_kpis(out, config.NOT_FOUND_LABEL["Zalora"])
            detected.append("Zalora")
        else:
            missing_pairs.append("Zalora StockValidation found but no SellerStockTemplate file.")

    # ---- Product Master cross-check (optional, applies to whatever marketplaces were processed) ----
    master_coverage_df = None
    if "product_master" in classified and marketplace_lookups:
        pm = readers.read_product_master(
            classified["product_master"].getvalue(), classified["product_master"].name
        )
        if pm["sku_col"] is None:
            notes.append(
                f"Product Master file '{classified['product_master'].name}' was read, but no SKU-like "
                "column could be identified — coverage check skipped. Rename the SKU column to "
                "something containing 'SKU' or tell me which column to use."
            )
        else:
            master_coverage_df = validate.master_coverage_report(pm["skus"], marketplace_lookups)
            # also flag, per marketplace, SKUs live in marketplace but absent from master
            for name, df, label in marketplaces:
                df["In Product Master"] = df["Seller SKU"].isin(pm["skus"])
            for name in kpi_results:
                mdf = next(d for n, d, l in marketplaces if n == name)
                kpi_results[name]["Not In Product Master"] = int((~mdf["In Product Master"]).sum())
            notes.append(
                f"Product Master matched on column '{pm['sku_col']}' "
                f"({len(pm['skus'])} unique SKUs)."
            )

    # ---- SOH vs ALL (optional warehouse check) ----
    soh_vs_all_df = None
    if "soh_report" in classified and "all_report" in classified:
        soh_df = readers.read_soh(classified["soh_report"].getvalue())
        all_result = readers.read_all_report(classified["all_report"].getvalue())
        notes.append(all_result["note"])
        if all_result["kind"] == "stock":
            soh_vs_all_df = validate.soh_vs_all(soh_df, all_result["lookup"])
        else:
            notes.append("SOH-vs-ALL comparison skipped because the ALL report wasn't recognized as stock data.")
    elif "soh_report" in classified and "all_report" not in classified:
        notes.append("SOH report uploaded but no ALL report found — warehouse check skipped.")
    elif "all_report" in classified and "soh_report" not in classified:
        notes.append("ALL report uploaded but no SOH report found — warehouse check skipped.")

    unmatched = [f.name for f in classified.get("unmatched", [])]

    return {
        "marketplaces": marketplaces,
        "kpi_results": kpi_results,
        "soh_vs_all_df": soh_vs_all_df,
        "master_coverage_df": master_coverage_df,
        "notes": notes,
        "detected": detected,
        "unmatched": unmatched,
        "missing_pairs": missing_pairs,
    }


def _read_csv(file_obj):
    import io
    import pandas as pd
    return pd.read_csv(io.BytesIO(file_obj.getvalue()))


def build_output_workbook(result: dict):
    return report.build_workbook(
        result["marketplaces"],
        result["kpi_results"],
        soh_vs_all_df=result["soh_vs_all_df"],
        master_coverage_df=result["master_coverage_df"],
        notes=result["notes"],
    )
