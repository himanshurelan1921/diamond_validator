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
    This is to match synonyms in Columns + Values sheets.
    """
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    s = unicodedata.normalize("NFKC", s).lower()

    # remove all non-word / non-space chars (punctuation)
    s = re.sub(r"[^\w\s]", " ", s)
    # collapse spaces to underscore
    s = re.sub(r"\s+", "_", s)
    s = s.strip("_")

    if not s:
        return None

    return s


def normalize_value_str(value):
    """
    STRICT normalization for VALUES:
    - lowercase
    - unicode normalize
    - strip leading/trailing spaces
    - DO NOT touch internal spaces or punctuation
    - exact text must appear in rules to be accepted
    """
    if value is None:
        return None

    # Pandas NaN
    if isinstance(value, float) and math.isnan(value):
        return None

    s = str(value).strip()
    if not s:
        return None

    if s.lower() == "nan":
        return None

    s = unicodedata.normalize("NFKC", s)

    return s.lower()


def is_empty_value(value):
    """
    Enhanced check for truly empty values.
    Returns True if value is None, NaN, empty string, or whitespace only.
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
# LOAD HEADER RULES (Columns sheet)
# ------------------------------------------------------------

def load_header_rules(rules_source):
    """
    rules_source: path or file-like (BytesIO)
    Uses Columns sheet:
      - Column Name  (canonical)
      - Column Values (comma-separated synonyms)
    Returns:
      header_map: normalized_supplier_header -> canonical_header_name
      canonical_set: set of canonical header names (normalized)
    """
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

        # include the canonical name itself as a valid synonym
        variants.append(str(canon_raw))

        for v in variants:
            norm = normalize_header_name(v)
            if norm:
                header_map[norm] = canon_norm

    return header_map, canonical_set


# ------------------------------------------------------------
# LOAD VALUE RULES (Values sheet)
# ------------------------------------------------------------

def load_value_rules(rules_source):
    """
    Uses Values sheet:
      - Value Type      (which column / field)
      - Base Value
      - Value Variations (comma-separated)
    Returns:
      rules: dict[normalized_value_type] = {
          "wildcard": bool,
          "allowed": set(normalized_value)
      }
    """
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
                if nm:
                    variations_norm.append(nm)
        else:
            nm = normalize_value_str(vars_raw)
            if nm:
                variations_norm.append(nm)

        # ensure entry exists
        if vtype_norm not in rules:
            rules[vtype_norm] = {"wildcard": False, "allowed": set()}

        # wildcard logic
        if base_norm == "any" or "any" in variations_norm:
            rules[vtype_norm]["wildcard"] = True
            # no need to store explicit values if wildcard
            continue

        # add base value
        if base_norm:
            rules[vtype_norm]["allowed"].add(base_norm)

        # add variations
        for nm in variations_norm:
            if nm != "any":
                rules[vtype_norm]["allowed"].add(nm)

    return rules


# ------------------------------------------------------------
# LOAD SUPPLIER FILE
# ------------------------------------------------------------

def load_supplier(path):
    ext = path.split(".")[-1].lower()
    if ext == "csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


# ------------------------------------------------------------
# APPLY HEADER NORMALIZATION
# ------------------------------------------------------------

def normalize_headers(df, header_map):
    """
    Map supplier headers -> canonical headers using header_map.
    Returns:
      new_df, unknown_headers
    """
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


# ------------------------------------------------------------
# NUMERIC RANGE VALIDATION
# ------------------------------------------------------------

def check_numeric_ranges(df):
    """
    Check for invalid numeric values in specific columns:
    - carat/weight: must be > 0
    - price fields: must be > 0
    """
    invalid = []
    
    # Check carat/weight columns
    weight_cols = [c for c in df.columns if c in ["carat", "weight", "carat_weight"]]
    for col in weight_cols:
        for idx, val in df[col].items():
            if is_empty_value(val):
                continue
            
            try:
                num_val = float(str(val).replace(",", ""))
                if num_val <= 0:
                    invalid.append(
                        f"Row {idx + 2}: Invalid carat value '{val}' in column '{col}' (must be > 0)"
                    )
            except (ValueError, TypeError):
                # Will be caught by value validation
                pass
    
    # Check price columns
    price_cols = [c for c in df.columns if c in ["price_per_carat", "total_sales_price"]]
    for col in price_cols:
        for idx, val in df[col].items():
            if is_empty_value(val):
                continue
            
            try:
                num_val = float(str(val).replace(",", ""))
                if num_val <= 0:
                    invalid.append(
                        f"Row {idx + 2}: Invalid price '{val}' in column '{col}' (must be > 0)"
                    )
            except (ValueError, TypeError):
                # Will be caught by value validation
                pass
    
    return invalid


# ------------------------------------------------------------
# VALUE VALIDATION
# ------------------------------------------------------------

def check_values(df, value_rules):
    """
    Strict value validation:
      - for each column, map to vtype via normalize_header_name
      - if vtype in value_rules and not wildcard:
          normalize supplier cell with normalize_value_str
          require EXACT match to an allowed normalized value
    """
    invalid = []

    for col in df.columns:
        vtype_norm = normalize_header_name(col)

        if vtype_norm not in value_rules:
            continue  # no rule for this column

        rule = value_rules[vtype_norm]
        if rule["wildcard"]:
            continue  # any value allowed

        allowed = rule["allowed"]

        for idx, val in df[col].items():
            norm_val = normalize_value_str(val)

            # missing values handled via mandatory check; here we only check non-empty
            if norm_val is None:
                continue

            if norm_val not in allowed:
                invalid.append(
                    f"Row {idx + 2}: Invalid '{val}' in column '{col}'"
                )

    return invalid


# ------------------------------------------------------------
# MANDATORY FIELDS
# ------------------------------------------------------------

# These are CANONICAL column names (after header normalization)
MANDATORY_COLS = [
    "stock_num",
    "shape",
    "color",
    "clarity",
    "lab",
    "image_url_1",
    "video_url_1",
    "cert_url_1",
]

def check_mandatory(df):
    """
    Check that mandatory canonical columns exist and non-empty.
    Uses enhanced is_empty_value() function.
    """
    missing = []
    for idx, row in df.iterrows():
        missing_cols = []
        for col in MANDATORY_COLS:
            if col in df.columns:
                if is_empty_value(row[col]):
                    missing_cols.append(col)
            else:
                # Column doesn't exist in df
                missing_cols.append(col)
        
        if missing_cols:
            missing.append(f"Row {idx + 2}: Missing {missing_cols}")
    return missing


# ------------------------------------------------------------
# FAST URL CHECKING (MULTI-THREADED)
# ------------------------------------------------------------

def fast_check_url(url):
    if url is None or str(url).strip() == "":
        return "NOT PROVIDED"

    try:
        r = requests.head(str(url).strip(), timeout=1)
        if r.status_code in [200, 301, 302]:
            return "WORKING"
        return f"NOT WORKING ({r.status_code})"
    except:
        return "NOT WORKING"


def check_all_urls(df):
    """
    Fast parallel URL checker.
    Any result that's not WORKING is reported.
    """
    url_cols = [c for c in df.columns if "url" in c.lower()]
    bad = []
    tasks = []

    with ThreadPoolExecutor(max_workers=30) as executor:
        for idx, row in df.iterrows():
            for col in url_cols:
                url = row[col] if col in df.columns else None
                future = executor.submit(fast_check_url, url)
                tasks.append((idx, col, url, future))

        for idx, col, url, future in tasks:
            try:
                result = future.result(timeout=2)
            except Exception:
                result = "NOT WORKING (timeout)"

            if result != "WORKING":
                bad.append(f"Row {idx + 2}: {col} → {result} → URL: {url}")

    return bad


# ------------------------------------------------------------
# MAIN (CLI)
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("supplier", help="Supplier CSV/XLSX")
    parser.add_argument("--rules", default="headers.xlsx")
    args = parser.parse_args()

    print("Loading rules...")
    header_map, canonical = load_header_rules(args.rules)
    value_rules = load_value_rules(args.rules)

    print("Loading supplier file...")
    df = load_supplier(args.supplier)

    print("Normalizing headers...")
    df, unknown_headers = normalize_headers(df, header_map)

    print("Checking mandatory fields...")
    missing = check_mandatory(df)

    print("Checking numeric ranges...")
    numeric_invalid = check_numeric_ranges(df)

    print("Checking values...")
    invalid = check_values(df, value_rules)

    print("Checking URLs (fast mode)...")
    url_bad = check_all_urls(df)

    print("\n----- VALIDATION REPORT -----\n")

    if unknown_headers:
        print("Unknown Headers:", unknown_headers, "\n")

    if missing:
        print("Missing Mandatory:", missing, "\n")

    if numeric_invalid:
        print("Invalid Numeric Values:", numeric_invalid, "\n")

    if invalid:
        print("Invalid Values:", invalid, "\n")

    if url_bad:
        print("URL Issues:", url_bad, "\n")

    print("Done!")


if __name__ == "__main__":
    main()
