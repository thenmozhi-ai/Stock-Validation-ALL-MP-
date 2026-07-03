"""
Central configuration: filename detection patterns, colour palette, and
platform-specific labels. Adding a new source or marketplace mostly means
touching this file plus one reader function.
"""

import re

# ---------------------------------------------------------------------------
# COLOURS
# ---------------------------------------------------------------------------
NAVY = "1F3864"
GREEN, GREEN_FONT = "C6EFCE", "375623"
RED, RED_FONT = "FFC7CE", "9C0006"
ORANGE, ORANGE_FONT = "FFEB9C", "7D4800"
GREY, GREY_FONT = "D9D9D9", "595959"
BLUE, BLUE_FONT = "DDEBF7", "1F4E78"
ALT_ROW = "F2F7FB"

# ---------------------------------------------------------------------------
# MARKETPLACES
# ---------------------------------------------------------------------------
MARKETPLACES = ["Lazada", "Shopee", "TikTok", "Zalora"]

NOT_FOUND_LABEL = {
    "Lazada": "NOT IN LAZADA",
    "Shopee": "NOT IN SHOPEE",
    "TikTok": "NOT IN TIKTOK",
    "Zalora": "NOT IN ZALORA",
}

ALL_NOT_FOUND_LABELS = set(NOT_FOUND_LABEL.values()) | {"NOT FOUND", "NOT IN MASTER"}

# ---------------------------------------------------------------------------
# FILENAME DETECTION PATTERNS
# Order matters: more specific patterns should be checked before generic ones.
# Each entry: (regex, file_role)
# file_role is a string consumed by detect.classify_file()
# ---------------------------------------------------------------------------
FILENAME_RULES = [
    (re.compile(r"stockvalidation.*lazada|lazada.*stockvalidation", re.I), "lazada_stockval"),
    (re.compile(r"stockvalidation.*shopee|shopee.*stockvalidation", re.I), "shopee_stockval"),
    (re.compile(r"stockvalidation.*tiktok|tiktok.*stockvalidation", re.I), "tiktok_stockval"),
    (re.compile(r"stockvalidation.*zalora|zalora.*stockvalidation", re.I), "zalora_stockval"),
    (re.compile(r"pricestock", re.I), "lazada_stock"),
    (re.compile(r"mass_update_sales_info", re.I), "shopee_stock"),
    (re.compile(r"tiktoksellercenter.*batchedit|batchedit", re.I), "tiktok_stock"),
    (re.compile(r"sellerstocktemplate", re.I), "zalora_stock"),
    (re.compile(r"sellerstatustemplate", re.I), "zalora_status"),
    (re.compile(r"sohbysku", re.I), "soh_report"),
    (re.compile(r"productmaster|product_master|master_sku|sku_master", re.I), "product_master"),
    (re.compile(r"^all[-_]?.*\.csv$", re.I), "all_report"),
]

# fallback generic stockvalidation match by marketplace keyword, if the
# combined pattern above didn't fire (covers "lazada_something.csv" style names)
GENERIC_MARKETPLACE_KEYWORDS = {
    "lazada_stockval": re.compile(r"lazada", re.I),
    "shopee_stockval": re.compile(r"shopee", re.I),
    "tiktok_stockval": re.compile(r"tiktok", re.I),
    "zalora_stockval": re.compile(r"zalora", re.I),
}
