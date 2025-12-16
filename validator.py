import pandas as pd
import requests
import math
import unicodedata
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
    return s.strip("_")

def normalize_value_str(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan": return None
    return unicodedata.normalize("NFKC", s).lower()

def to_float(x):
    if x is None or (isinstance(x, float) and math.isnan(x)): return None
    s = str(x).strip().replace(",", "")
    s = re.sub(r"[^\d.\-]", "", s)
    try: return float(s)
    except: return None

# ------------------------------------------------------------
# CORE LOGIC
# ------------------------------------------------------------

RANGE_COLS = {
    "carat": (0.05, 20.0, "Carat Weight"),
    "weight": (0.05, 20.0, "Carat Weight"),
    "table": (40.0, 90.0, "Table %"),
    "depth": (40.0, 90.0, "Depth %"),
}

MANDATORY_COLS = ["stock_num", "shape", "color", "clarity", "lab", "image_url_1", "video_url_1", "cert_url_1"]

def load_header_rules(rules_source):
    df = pd.read_excel(rules_source, sheet_name="Columns")
    h_map = {}
    for _, row in df.iterrows():
        canon = normalize_header_name(row.get("Column Name"))
        vars = str(row.get("Column Values", "")).split(",") + [str(row.get("Column Name"))]
        for v in vars:
            nm = normalize_header_name(v)
            if nm: h_map[nm] = canon
    return h_map, set(h_map.values())

def load_value_rules(rules_source):
    df = pd.read_excel(rules_source, sheet_name="Values")
    rules = {}
    for _, row in df.iterrows():
        vtype = normalize_header_name(row.get("Value Type"))
        if not vtype: continue
        if vtype not in rules: rules[vtype] = {"wildcard": False, "allowed": set()}
        
        # --- FIX: Support normalized values like 'Y' or 'Uknown' ---
        base_raw = row.get("Base Value")
        base_norm = normalize_value_str(base_raw)
        
        if base_norm == "any":
            rules[vtype]["wildcard"] = True
            continue
        
        if base_norm: rules[vtype]["allowed"].add(base_norm)
        
        vars_raw = str(row.get("Value Variations", ""))
        for part in vars_raw.split(","):
            nm = normalize_value_str(part)
            if nm: rules[vtype]["allowed"].add(nm)
        # -----------------------------------------------------------
    return rules

def normalize_headers(df, header_map):
    unknown, new_cols = [], {}
    for col in df.columns:
        norm = normalize_header_name(col)
        if norm in header_map: new_cols[col] = header_map[norm]
        else: unknown.append(col)
    return df.rename(columns=new_cols), unknown

def check_mandatory(df):
    missing = []
    for idx, row in df.iterrows():
        miss_cols = [c for c in MANDATORY_COLS if c in df.columns and normalize_value_str(row[c]) is None]
        if miss_cols: missing.append(f"Row {idx + 2}: Missing {miss_cols}")
    return missing

def check_values(df, value_rules):
    invalid = []
    for col in df.columns:
        vtype = normalize_header_name(col)
        if vtype not in value_rules or value_rules[vtype]["wildcard"]: continue
        allowed = value_rules[vtype]["allowed"]
        for idx, val in df[col].items():
            nv = normalize_value_str(val)
            if nv and nv not in allowed:
                invalid.append(f"Row {idx + 2}: Invalid '{val}' in column '{col}'")
    return invalid

def check_ranges(df):
    out = []
    for col, (min_v, max_v, name) in RANGE_COLS.items():
        if col in df.columns:
            for idx, val in df[col].items():
                fv = to_float(val)
                if fv is not None and (fv < min_v or fv > max_v):
                    out.append(f"Row {idx + 2}: Out of Range '{val}' in column '{col}'")
    return out

def check_all_urls(df):
    url_cols = [c for c in df.columns if "url" in c.lower() and c != "cert_url_1"]
    bad = []
    def fast_check(url):
        if not url or str(url).strip() == "": return "NOT PROVIDED"
        try:
            r = requests.head(str(url).strip(), timeout=1)
            return "WORKING" if r.status_code in [200, 301, 302] else f"NOT WORKING ({r.status_code})"
        except: return "NOT WORKING"
    with ThreadPoolExecutor(max_workers=20) as exe:
        tasks = [(idx, col, exe.submit(fast_check, row[col])) for idx, row in df.iterrows() for col in url_cols]
        for idx, col, fut in tasks:
            if fut.result() != "WORKING": bad.append(f"Row {idx+2}: {col} → {fut.result()}")
    return bad
