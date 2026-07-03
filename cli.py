"""
CLI usage:
    python cli.py path/to/file1.csv path/to/file2.xlsx ... --out result.xlsx

Useful for scheduled/batch runs (cron, CI) without spinning up Streamlit.
"""

import argparse
import os
import sys

from stock_validation import pipeline


def main():
    parser = argparse.ArgumentParser(description="Run stock validation over a folder or list of files.")
    parser.add_argument("paths", nargs="+", help="File paths to process")
    parser.add_argument("--out", default="Stock_Validation_All_Sources.xlsx", help="Output workbook path")
    args = parser.parse_args()

    blobs = []
    for path in args.paths:
        if not os.path.isfile(path):
            print(f"Skipping (not a file): {path}", file=sys.stderr)
            continue
        with open(path, "rb") as fh:
            blobs.append(pipeline.UploadedBlob(os.path.basename(path), fh.read()))

    if not blobs:
        print("No valid files provided.", file=sys.stderr)
        sys.exit(1)

    result = pipeline.run_pipeline(blobs)

    print("Detected marketplaces:", ", ".join(result["detected"]) or "none")
    for note in result["missing_pairs"] + result["notes"]:
        print("NOTE:", note)
    for name, kpis in result["kpi_results"].items():
        print(f"\n{name} KPIs:")
        for k, v in kpis.items():
            print(f"  {k}: {v}")

    if not result["marketplaces"]:
        print("Nothing to report — no complete marketplace pairs found.", file=sys.stderr)
        sys.exit(1)

    workbook_buf = pipeline.build_output_workbook(result)
    with open(args.out, "wb") as fh:
        fh.write(workbook_buf.getvalue())
    print(f"\nSaved: {args.out}")


if __name__ == "__main__":
    main()
