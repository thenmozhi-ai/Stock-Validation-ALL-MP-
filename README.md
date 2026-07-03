# Stock Validation — ALL Sources

A single modular Python project that auto-detects uploaded files by filename and validates
stock across **Lazada, Shopee, TikTok, Zalora, Warehouse SOH, ALL Report, and Product Master**,
producing one colour-coded Excel workbook. Works with any subset of sources — upload just
what you have.

Runs as a Streamlit app (`app.py`) or a CLI (`cli.py`) — both share the same pipeline code,
so there's exactly one place the validation logic lives.

## Project structure

```
stock_validation_project/
├── app.py                     # Streamlit UI (thin — just wiring)
├── cli.py                     # command-line entry point for batch/scheduled runs
├── requirements.txt
├── stock_validation/
│   ├── config.py               # colours, marketplace labels, filename detection patterns
│   ├── detect.py                # classifies each uploaded file by filename
│   ├── readers.py                # one reader per source: parses raw bytes -> DataFrame/lookup
│   ├── validate.py               # remark logic, KPIs, product-master cross-checks
│   ├── report.py                 # builds the styled Excel workbook
│   └── pipeline.py               # orchestrates detect -> read -> validate -> report
└── tests/
    └── test_pipeline.py          # smoke test using the actual sample files
```

Adding a new marketplace means: one filename pattern in `config.py`, one reader function in
`readers.py`, one branch in `pipeline.run_pipeline()`. The remark logic, styling, and workbook
structure are shared automatically.

## What it detects

| Source | Filename should contain |
|---|---|
| Lazada StockValidation | `stockValidation-lazada` (or `lazada`), `.csv` |
| Lazada Stock & Price | `pricestock` |
| Shopee StockValidation | `stockValidation-shopee` (or `shopee`), `.csv` |
| Shopee Mass Update | `mass_update_sales_info` (`.xlsx` or `.zip`) |
| TikTok StockValidation | `stockValidation-tiktok` (or `tiktok`), `.csv` |
| TikTok Batch Edit | `Tiktoksellercenter_batchedit` or `batchedit` |
| Zalora StockValidation | `stockValidation-zalora` (or `zalora`), `.csv` |
| Zalora Stock (required) | `SellerStockTemplate` |
| Zalora Status (optional) | `SellerStatusTemplate` |
| Warehouse SOH (optional) | `SOHbySKU` |
| ALL Report (optional) | starts with `ALL`, `.csv` |
| Product Master (optional) | contains `ProductMaster`, `product_master`, `master_sku`, or `sku_master` |

## Remark logic (same for every marketplace)

1. SKU not found in the marketplace's own stock file → `NOT IN <MARKETPLACE>`
2. Expected stock = 0, marketplace stock > 0 → `UPDATE 0`
3. Expected stock > 0, marketplace stock = 0 → `UPDATE STOCK`
4. Expected ≠ marketplace stock → `Stock Mismatch`
5. Otherwise → `TRUE`

## Product Master cross-check

The Product Master reader is **best-effort**: it looks for any column containing `sku` and,
optionally, `status`/`active`/`lifecycle`. If found, it adds:
- an `In Product Master` boolean column on every marketplace's StockVal sheet
- a `Not In Product Master` KPI on the Summary sheet
- a `Product_Master_Coverage` sheet showing, per master SKU, which marketplaces actually
  carry it

**This part hasn't been tested against a real Product Master export yet** — column-name
guessing is inherently fuzzy. If your file has a non-obvious layout, tell me its actual
columns and I'll hard-code the mapping instead of guessing.

## ALL Report — handled carefully

Some "ALL-DD-Mon-YYYY.csv" exports are genuine warehouse stock snapshots; others are
order-transaction logs that merely share the filename convention (we ran into exactly this
in an earlier session). `readers.read_all_report()` checks for order-like columns
(`order_id`, `tracking_number`, etc.) and skips the SOH-vs-ALL comparison with a note if it
looks transactional, rather than silently comparing the wrong thing.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or via CLI:
```bash
python cli.py path/to/*.csv path/to/*.xlsx --out result.xlsx
```

## Deploy on GitHub + Streamlit Community Cloud

```bash
git init
git add .
git commit -m "Stock validation — all sources"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

Then on [share.streamlit.io](https://share.streamlit.io): **New app** → select the repo/branch →
main file path `app.py` → **Deploy**.

## Tests

```bash
python -m pytest tests/
```

`tests/test_pipeline.py` is a smoke test that runs the pipeline against real sample files and
checks the KPI numbers match known-good values — useful as a regression check whenever the
parsing logic changes.
