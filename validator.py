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
    if value is None: return None
    s = str(value).strip()
    if not s: return None
    s = unicodedata.normalize("NFKC", s).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_") if s.strip("_") else None

def normalize_value_str(value):
    if value is None: return None
    if isinstance(value, float) and math.isnan(value): return None
    s = str(value).strip()
    if not s or s.lower() == "nan": return None
    return unicodedata.normalize("NFKC", s).lower()

def is_empty_value(value):
    if value is None: return True
    if isinstance(value, float) and math.isnan(value): return True
    s = str(value).strip()
    return not s or s.lower() == "nan"

# ------------------------------------------------------------
# RULES LOADER
# ------------------------------------------------------------

def load_header_rules(rules_source):
    """Loads mapping from synonyms to canonical column names."""
    try:
        df = pd.read_excel(rules_source, sheet_name="Columns")
        header_map = {}
        for _, row in df.iterrows():
            canon_raw = row.get("Column Name")
            canon_norm = normalize_header_name(canon_raw)
            if not canon_norm: continue
            
            variants_raw = row.get("Column Values", "")
            variants = [v.strip() for v in str(variants_raw).split(",")] if pd.notna(variants_raw) else []
            variants.append(str(canon_raw))
            
            for v in variants:
                norm = normalize_header_name(v)
                if norm: header_map[norm] = canon_norm
        return header_map
    except Exception as e:
        print(f"Error loading header rules: {e}")
        return {}

def load_value_rules(rules_source):
    """Loads allowed values for specific columns."""
    try:
        df = pd.read_excel(rules_source, sheet_name="Values")
        rules = {}
        for _, row in df.iterrows():
            vtype_norm = normalize_header_name(row.get("Value Type"))
            if not vtype_norm: continue
            
            base_norm = normalize_value_str(row.get("Base Value"))
            vars_raw = row.get("Value Variations", "")
            vars_norm = [normalize_value_str(v) for v in str(vars_raw).split(",")] if pd.notna(vars_raw) else []
            
            if vtype_norm not in rules: rules[vtype_norm] = {"wildcard": False, "allowed": set()}
            
            if base_norm == "any" or "any" in vars_norm:
                rules[vtype_norm]["wildcard"] = True
                continue
                
            if base_norm: rules[vtype_norm]["allowed"].add(base_norm)
            for v in vars_norm:
                if v and v != "any": rules[vtype_norm]["allowed"].add(v)
        return rules
    except Exception as e:
        print(f"Error loading value rules: {e}")
        return {}

# ------------------------------------------------------------
# VALIDATION ENGINE
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

def check_ranges(df):
    """Matches the call in your error: validator.check_ranges(df)"""
    invalid = []
    # Identify numeric columns to check
    cols_to_check = ["carat", "weight", "carat_weight", "price_per_carat", "total_sales_price"]
    active_cols = [c for c in df.columns if c in cols_to_check]
    
    for col in active_cols:
        for idx, val in df[col].items():
            if is_empty_value(val): continue
            try:
                # Remove commas and convert to float
                num = float(str(val).replace(",", ""))
                if num <= 0:
                    invalid.append(f"Row {idx + 2}: {col} value {val} must be > 0")
            except:
                # If it's not a number, let check_values handle the format error
                pass
    return invalid

def parse_range_issue_strings(issues, df):
    """Matches the call in your error: parse_range_issue_strings(...)"""
    # Simply returns the list of strings for the UI to display
    return issues

def check_values(df, value_rules):
    invalid = []
    for col in df.columns:
        vtype = normalize_header_name(col)
        if vtype not in value_rules: continue
        
        rule = value_rules[vtype]
        if rule["wildcard"]: continue
        
        allowed = rule["allowed"]
        for idx, val in df[col].items():
            norm_val = normalize_value_str(val)
            if norm_val and norm_val not in allowed:
                invalid.append(f"Row {idx + 2}: Invalid '{val}' in '{col}'")
    return invalid

def check_mandatory(df):
    mandatory = ["stock_num", "shape", "color", "clarity", "lab", "image_url_1", "video_url_1", "cert_url_1"]
    missing = []
    for idx, row in df.iterrows():
        errs = [c for c in mandatory if c not in df.columns or is_empty_value(row[c])]
        if errs:
            missing.append(f"Row {idx + 2}: Missing {errs}")
    return missing

def fast_check_url(url):
    if not url or str(url).strip() == "": return "NOT PROVIDED"
    try:
        r = requests.head(str(url).strip(), timeout=1.5, allow_redirects=True)
        return "WORKING" if r.status_code < 400 else f"ERROR ({r.status_code})"
    except:
        return "TIMEOUT/FAILED"

def check_urls(df):
    url_cols = [c for c in df.columns if "url" in c.lower()]
    bad = []
    with ThreadPoolExecutor(max_workers=20) as exe:
        tasks = []
        for idx, row in df.iterrows():
            for col in url_cols:
                tasks.append((idx, col, row[col], exe.submit(fast_check_url, row[col])))
        
        for idx, col, url, fut in tasks:
            res = fut.result()
            if res != "WORKING":
                bad.append(f"Row {idx + 2}: {col} -> {res}")
    return bad

# ------------------------------------------------------------
# MAIN (For Local/CLI Testing)
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("supplier", nargs='?')
    parser.add_argument("--rules", default="headers.xlsx")
    args = parser.parse_args()

    if not args.supplier:
        print("Usage: python app.py <supplier_file>")
        return

    print("Loading rules and data...")
    header_map = load_header_rules(args.rules)
    value_rules = load_value_rules(args.rules)
    
    # Load data
    if args.supplier.endswith('csv'):
        df = pd.read_csv(args.supplier)
    else:
        df = pd.read_excel(args.supplier)
    
    # Process
    df, unknown = normalize_headers(df, header_map)
    
    # This matches the exact line failing in your Streamlit Traceback
    range_issues = parse_range_issue_strings(check_ranges(df), df)
    
    print(f"Found {len(range_issues)} range issues.")
    for issue in range_issues[:10]:
        print(issue)

if __name__ == "__main__":
    main()
