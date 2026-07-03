"""
Streamlit UI. All the real logic lives in stock_validation/*.py — this file
just wires uploads to the pipeline and renders results.
"""

import streamlit as st

from stock_validation import pipeline

st.set_page_config(page_title="Stock Validation — ALL Sources", page_icon="📦", layout="wide")

st.title("📦 Stock Validation — ALL Sources")
st.caption(
    "Lazada · Shopee · TikTok · Zalora · Warehouse SOH · ALL Report · Product Master — "
    "upload any subset, files are auto-detected by name."
)

with st.expander("ℹ️ What files can I upload?", expanded=False):
    st.markdown("""
| Source | Filename should contain |
|---|---|
| Lazada StockValidation | `stockValidation-lazada` (or just `lazada`), `.csv` |
| Lazada Stock & Price | `pricestock` |
| Shopee StockValidation | `stockValidation-shopee` (or just `shopee`), `.csv` |
| Shopee Mass Update | `mass_update_sales_info` (`.xlsx` or `.zip`) |
| TikTok StockValidation | `stockValidation-tiktok` (or just `tiktok`), `.csv` |
| TikTok Batch Edit | `Tiktoksellercenter_batchedit` or `batchedit` |
| Zalora StockValidation | `stockValidation-zalora` (or just `zalora`), `.csv` |
| Zalora Stock (required) | `SellerStockTemplate` |
| Zalora Status (optional) | `SellerStatusTemplate` |
| Warehouse SOH (optional) | `SOHbySKU` |
| ALL Report (optional) | filename starts with `ALL`, `.csv` — only used if it's genuinely stock data |
| Product Master (optional) | filename contains `ProductMaster`, `product_master`, `master_sku`, or `sku_master` |

Any subset works — upload just what you have.
""")

uploaded_files = st.file_uploader(
    "Upload your files",
    accept_multiple_files=True,
    type=["csv", "xlsx", "xls", "zip"],
)

if not uploaded_files:
    st.info("Upload StockValidation CSVs + matching marketplace stock files above to get started.")
    st.stop()

result = pipeline.run_pipeline(uploaded_files)

if result["unmatched"]:
    st.warning("Couldn't identify these files — check filenames match the patterns above: " + ", ".join(result["unmatched"]))

for note in result["missing_pairs"]:
    st.warning(note)

st.write("**Detected marketplaces:**", ", ".join(result["detected"]) if result["detected"] else "None yet.")

if not result["detected"]:
    st.stop()

if st.button("🚀 Run validation", type="primary"):
    with st.spinner("Building workbook..."):
        workbook_buf = pipeline.build_output_workbook(result)

    st.success("Done!")

    cols = st.columns(len(result["marketplaces"]))
    for i, (name, df, label) in enumerate(result["marketplaces"]):
        with cols[i]:
            k = result["kpi_results"][name]
            st.metric(f"{name} accuracy", f"{k['Accuracy %']}%", f"{k['Total Issues']} issues")

    if result["notes"]:
        with st.expander("📋 Processing notes"):
            for note in result["notes"]:
                st.write("•", note)

    st.download_button(
        "⬇️ Download Stock_Validation_All_Sources.xlsx",
        data=workbook_buf,
        file_name="Stock_Validation_All_Sources.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
