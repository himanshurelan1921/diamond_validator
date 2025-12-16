import pandas as pd
import requests
import math
import unicodedata
import argparse
from concurrent.futures import ThreadPoolExecutor
import re

# ------------------------------------------------------------
# NORMALIZERS
# ------------------------------------------------------------

def normalize_header_name(value):
    """
    Normalization for HEADER NAMES:
    - lowercase, unicode normalize, remove punctuation, collapse spaces to underscores.
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", "_", s)
    s = s.strip("_")
    return s if s else None


def normalize_value_str(value):
    """
    STRICT normalization for VALUES:
    - lowercase, unicode normalize, strip leading/trailing spaces.
    """
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    return unicodedata.normalize("NFKC", s).lower()


def to_float(x):
    """Helper to safely convert any input to a float for math/range checks."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).strip().replace(",", "")
    # Remove any character that isn't a digit, dot, or minus sign
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except:
        return None


# ------------------------------------------------------------
# CONFIGURATION & RULES
# ------------------------------------------------------------

# [Canonical Name]: (min, max, label)
RANGE_COLS = {
    "carat": (0.05, 20.0, "Carat Weight"),
    "weight": (0.05, 20.0, "Carat Weight"),
    "carat_weight": (0.05, 20.0, "Carat Weight"),
    "table": (40.0, 90.0, "Table %"),
    "depth": (40.0, 90.0, "Depth %"),
}

MANDATORY_COLS = [
    "stock_num", "shape", "color", "clarity", 
    "lab", "image_url_1", "video_url_1", "cert_url_1"
]


def load_header_rules(rules_source):
    df = pd.read_excel(rules_source, sheet_name="Columns")
    header_map = {}
    canonical_set = set()

    for _, row in df.iterrows():
        canon_raw = row.get("Column Name", None)
        canon_norm = normalize_header_name(canon_raw)
        if not canon_norm:
            continue

        canonical_set.add(canon_norm)
        variants_raw = str(row.get("Column Values", ""))
        variants = [v.strip() for v in variants_raw.split(",") if v.strip()]
        variants.append(str(canon_raw))

        for v in variants:
            norm = normalize_header_name(v)
            if norm:
                header_map[norm] = canon_norm
    return header_map, canonical_set


def load_value_rules(rules_source):
    df = pd.read_excel(rules_source, sheet_name="Values")
    rules = {}

    for _, row in df.iterrows():
        vtype_norm = normalize_header_name(row.get("Value Type", None))
        if not vtype_norm:
            continue

        if vtype_norm not in rules:
            rules[vtype_norm] = {"wildcard": False, "allowed": set()}

        base_norm = normalize_value_str(row.get("Base Value", None))
        if base_norm == "any":
            rules[vtype_norm]["wildcard"] = True
            continue

        if base_norm:
            rules[vtype_norm]["allowed"].add(base_norm)

        vars_raw = str(row.get("Value Variations", ""))
        for part in vars_raw.split(","):
            nm = normalize_value_str(part)
            if nm:
                rules[vtype_norm]["allowed"].add(nm)
    return rules


# ------------------------------------------------------------
# VALIDATION FUNCTIONS
# ------------------------------------------------------------

def normalize_headers(df, header_map):
    unknown = []
    new_cols = {}
    for col in df.columns:
        norm = normalize_header_name(col)
        if norm in header_map:
            new_cols[col] = header_map[norm]
        else:
            unknown.append(col)
    return df.rename(columns=new_cols), unknown


def check_values(df, value_rules):
    invalid = []
    for col in df.columns:
        vtype_norm = normalize_header_name(col)
        if vtype_norm not in value_rules:
            continue
        rule = value_rules[vtype_norm]
        if rule["wildcard"]:
            continue

        for idx, val in df[col].items():
            norm_val = normalize_value_str(val)
            if norm_val and norm_val not in rule["allowed"]:
                invalid.append(f"Row {idx + 2}: Invalid '{val}' in column '{col}'")
    return invalid


def check_ranges(df):
    """Checks Carat, Table, and Depth against RANGE_COLS."""
    out_of_range = []
    for col, (min_val, max_val, name) in RANGE_COLS.items():
        if col not in df.columns:
            continue
        for idx, val in df[col].items():
            f_val = to_float(val)
            if f_val is not None and (f_val < min_val or f_val > max_val):
                out_of_range.append(f"Row {idx + 2}: Out of Range '{val}' in column '{col}'")
    return out_of_range


def check_mandatory(df):
    missing = []
    for idx, row in df.iterrows():
        missing_cols = []
        for col in MANDATORY_COLS:
            if col in df.columns:
                if normalize_value_str(row[col]) is None:
                    missing_cols.append(col)
        if missing_cols:
            missing.append(f"Row {idx + 2}: Missing {missing_cols}")
    return missing


def check_all_urls(df):
    """Fast parallel check for image and video URLs (Skips Cert URL status check)."""
    url_cols = [c for c in df.columns if "url" in c.lower() and c != "cert_url_1"]
    bad = []
    tasks = []

    def fast_check(url):
        if not url or str(url).strip() == "": return "NOT PROVIDED"
        try:
            r = requests.head(str(url).strip(), timeout=1)
            return "WORKING" if r.status_code in [200, 301, 302] else f"NOT WORKING ({r.status_code})"
        except: return "NOT WORKING"

    with ThreadPoolExecutor(max_workers=20) as executor:
        for idx, row in df.iterrows():
            for col in url_cols:
                url = row[col]
                tasks.append((idx, col, executor.submit(fast_check, url)))
        
        for idx, col, future in tasks:
            res = future.result()
            if res != "WORKING":
                bad.append(f"Row {idx + 2}: {col} → {res}")
    return bad

def find_non_pdf_cert_urls(df):
    """Validation specifically for Certificate PDF format."""
    issues = []
    count = 0
    if "cert_url_1" in df.columns:
        for idx, row in df.iterrows():
            url = row["cert_url_1"]
            if normalize_value_str(url) and ".pdf" not in str(url).lower():
                count += 1
                issues.append({
                    "Category": "URL Issue", "Stock No.": row.get("stock_num"),
                    "Issue Type": "Cert URL Format", "Column": "cert_url_1",
                    "Value": url, "Details": "Not a direct PDF link", "Row": idx + 2
                })
    return issues, count

def build_price_mismatch_issues(df):
    """Checks if Carat * Price Per Carat matches Total Sales Price."""
    issues = []
    count = 0
    if all(c in df.columns for c in ["carat", "price_per_carat", "total_sales_price"]):
        for idx, row in df.iterrows():
            w = to_float(row["carat"])
            p = to_float(row["price_per_carat"])
            t = to_float(row["total_sales_price"])
            if w and p and t:
                expected = round(w * p, 2)
                if abs(expected - t) > 0.1:
                    count += 1
                    issues.append({
                        "Category": "Price", "Stock No.": row.get("stock_num"),
                        "Issue Type": "Price Mismatch", "Column": "total_sales_price",
                        "Value": t, "Details": f"Expected {expected}", "Row": idx + 2
                    })
    return issues, count
