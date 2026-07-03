"""
Detects what kind of file each upload is, purely from filename patterns
(with a light content sniff as a fallback for ambiguous CSVs).
"""

from . import config


def classify_file(filename: str) -> str | None:
    """Returns a role string (e.g. 'lazada_stockval', 'pricestock', ...)
    or None if the filename doesn't match anything known."""
    name = filename.strip()

    for pattern, role in config.FILENAME_RULES:
        if pattern.search(name):
            # A file could match both a specific stockval pattern (e.g.
            # 'stockValidation-shopee...csv') AND accidentally contain
            # 'mass_update' etc. Specific rules are listed first in
            # config.FILENAME_RULES so first match wins.
            return role

    # fallback: bare "<marketplace>....csv" StockValidation exports that
    # don't literally say "stockValidation" in the name
    if name.lower().endswith(".csv"):
        for role, pattern in config.GENERIC_MARKETPLACE_KEYWORDS.items():
            if pattern.search(name):
                return role

    return None


def classify_uploads(files) -> dict:
    """
    files: iterable of objects with a `.name` attribute (Streamlit's
    UploadedFile, or anything similar).

    Returns a dict like:
    {
        "lazada_stockval": <file>,
        "lazada_stock": <file>,
        "shopee_stockval": <file>,
        "shopee_stock": [<file>, ...],   # shopee can have multiple xlsx parts
        "product_master": <file>,
        "unmatched": [<file>, ...],
        ...
    }
    """
    result: dict = {"unmatched": []}
    for f in files:
        role = classify_file(f.name)
        if role is None:
            result["unmatched"].append(f)
            continue
        if role == "shopee_stock":
            result.setdefault(role, []).append(f)
        else:
            # last one wins if duplicates uploaded; caller can warn on this
            result[role] = f
    return result
