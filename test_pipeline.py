"""
Smoke test: runs the pipeline against the real sample files used in earlier
sessions and checks the KPI numbers match known-good values. Run with:
    python -m pytest tests/
"""

import glob
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from stock_validation import pipeline

UPLOADS_DIR = "/mnt/user-data/uploads"

EXPECTED = {
    "Lazada": {"Total SKUs": 6138, "Matched": 6013, "Accuracy %": 97.96},
    "Shopee": {"Total SKUs": 5830, "Matched": 5489, "Accuracy %": 94.15},
    "TikTok": {"Total SKUs": 8873, "Matched": 7575, "Accuracy %": 85.37},
}


def _load_blobs():
    blobs = []
    for path in glob.glob(os.path.join(UPLOADS_DIR, "*")):
        with open(path, "rb") as fh:
            blobs.append(pipeline.UploadedBlob(os.path.basename(path), fh.read()))
    return blobs


def test_known_good_kpis():
    blobs = _load_blobs()
    if not blobs:
        import pytest
        pytest.skip("Sample files not present in this environment.")

    result = pipeline.run_pipeline(blobs)

    for name, expected in EXPECTED.items():
        assert name in result["kpi_results"], f"{name} not detected"
        actual = result["kpi_results"][name]
        for key, val in expected.items():
            assert actual[key] == val, f"{name} {key}: expected {val}, got {actual[key]}"


if __name__ == "__main__":
    test_known_good_kpis()
    print("All good.")
