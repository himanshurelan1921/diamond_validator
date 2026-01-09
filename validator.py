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
    Normalization for HEADER NAMES (column names, Value Type):
    - lowercase
    - unicode normalize
    - remove punctuation
    - collapse spaces to single underscore
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
    STRICT normalization for VALUES.
    """
    if value is None:
        return None

    if isinstance(value, float) and math.isnan(value):
        return None

    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None

    s = unicodedata.normalize("NFKC", s)
    return s.lower()


def is_empty_value(value):
    """
    Returns True if value is None, NaN, empty string, or whitespace.
    """
    if value is None:
        return True
    
    if isinstance(value, float) and math.isnan(value):
        return True
    
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return True
    
    return False


# ------------------------------------------------------------
# LOAD RULES
# ------------------------------------------------------------

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
        variants_raw = row.get("Column Values", "")
        variants = []
        if isinstance(variants_raw, str):
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
        vtype_raw = row.get("Value Type", None)
        vtype_norm = normalize_header_name(vtype_raw)
        if not vtype_norm:
            continue

        base_norm = normalize_value_str(row.get("Base Value", None))
        variations_norm = []
        vars_raw = row.get("Value Variations", None)
        
        if isinstance(vars_raw, str):
            for part in vars_raw.split(","):
                nm = normalize_value_str(part)
                if nm: variations_norm.append(nm)
        else:
            nm = normalize_value_str(vars_raw)
            if nm: variations_norm.append(nm)

        if vtype_norm not in rules:
            rules[vtype_norm] = {"wildcard": False, "allowed": set()}

        if base_norm == "any" or "any" in variations_norm:
            rules[vtype_norm]["wildcard"] = True
            continue

        if base_norm:
            rules[vtype_norm]["allowed"].add(base_norm)

        for nm in variations_norm:
            if nm != "any":
                rules[vtype_norm]["allowed"].add(nm)

    return rules


# ------------------------------------------------------------
# LOAD SUPPLIER FILE
# ------------------------------------------------------------

def load_supplier(path):
    ext = str(path).split(".")[-1].lower()
    if ext == "csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


# ------------------------------------------------------------
# VALIDATION LOGIC
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
    df = df.rename(columns=new_cols)
    return df, unknown


def check_numeric_ranges(df):
    invalid = []
    weight_cols = [c for c in df.columns if c in ["carat", "weight", "carat_weight"]]
    price_cols = [c for c in df.columns if c in ["price_per_carat", "total_sales_price"]]
    
    for col in weight_cols + price_cols:
        if col not in df.columns: continue
        for idx, val in df[col].items():
            if is_empty_value(val): continue
            try:
                num_val = float(str(val).replace(",", ""))
                if num_val <= 0:
                    invalid.append(f"Row {idx + 2}: Invalid value '{val}' in '{col}' (must be > 0)")
            except (ValueError, TypeError):
                pass
    return invalid


def check_values(df, value_rules):
    invalid = []
    for col in df.columns:
        vtype_norm = normalize_header_name(col)
        if vtype_norm not in value_rules: continue
        
        rule = value_rules[vtype_norm]
        if rule["wildcard"]: continue

        for idx, val in df[col].items():
            norm_val = normalize_value_str(val)
            if norm_val is None: continue
            if norm_val not in rule["allowed"]:
                invalid.append(f"Row {idx + 2}: Invalid '{val}' in column '{col}'")
    return invalid


MANDATORY_COLS = ["stock_num", "shape", "color", "clarity", "lab", "image_url_1", "video_url_1", "cert_url_1"]

def check_mandatory(df):
    missing = []
    for idx, row in df.iterrows():
        missing_cols = [col for col in MANDATORY_COLS if col not in df.columns or is_empty_value(row[col])]
        if missing_cols:
            missing.append(f"Row {idx + 2}: Missing {missing_cols}")
    return missing


# ------------------------------------------------------------
# URL CHECKING
# ------------------------------------------------------------

def fast_check_url(url):
    if not url or str(url).strip() == "":
        return "NOT PROVIDED"
    try:
        r = requests.head(str(url).strip(), timeout=2)
        if r.status_code in [200, 301, 302]:
            return "WORKING"
        return f"NOT WORKING ({r.status_code})"
    except:
        return "NOT WORKING"


def check_all_urls(df):
    url_cols = [c for c in df.columns if "url" in c.lower()]
    bad = []
    tasks = []

    with ThreadPoolExecutor(max_workers=20) as executor:
        for idx, row in df.iterrows():
            for col in url_cols:
                url = row[col]
                tasks.append((idx, col, url, executor.submit(fast_check_url, url)))

        for idx, col, url, future in tasks:
            try:
                res = future.result()
                if res != "WORKING":
                    bad.append(f"Row {idx + 2}: {col} -> {res} -> URL: {url}")
            except:
                bad.append(f"Row {idx + 2}: {col} -> ERROR -> URL: {url}")
    return bad


# ------------------------------------------------------------
# MAIN EXECUTION
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Diamond Data Validator")
    parser.add_argument("supplier", nargs='?', help="Path to supplier CSV or XLSX")
    parser.add_argument("--rules", default="headers.xlsx", help="Path to rules file")
    args = parser.parse_args()

    if not args.supplier:
        print("Usage: python app.py <supplier_file_path> [--rules headers.xlsx]")
        return

    print(f"--- Starting Validation ---")
    
    try:
        header_map, canonical = load_header_rules(args.rules)
        value_rules = load_value_rules(args.rules)
        df = load_supplier(args.supplier)
    except Exception as e:
        print(f"Critical Error: {e}")
        return

    df, unknown_headers = normalize_headers(df, header_map)
    
    results = {
        "Unknown Headers": unknown_headers,
        "Missing Mandatory": check_mandatory(df),
        "Invalid Numeric": check_numeric_ranges(df),
        "Invalid Values": check_values(df, value_rules),
        "URL Issues": check_all_urls(df)
    }

    print("\n----- VALIDATION REPORT -----")
    for key, issues in results.items():
        if issues:
            print(f"\n[{key}]: {len(issues)} issues found.")
            for issue in issues[:10]: # Print first 10
                print(f" - {issue}")
            if len(issues) > 10:
                print(f" ... and {len(issues)-10} more.")
        else:
            print(f"\n[{key}]: Passed")

    print("\nDone!")


if __name__ == "__main__":
    main()
