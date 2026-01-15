import pandas as pd
import requests
import math
import unicodedata
import argparse
import io
from concurrent.futures import ThreadPoolExecutor
import re
from urllib.parse import urlparse

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

    if not s:
        return None

    return s


def normalize_value_str(value):
    """
    STRICT normalization for VALUES:
    - lowercase
    - unicode normalize
    - strip leading/trailing spaces
    """
    if value is None:
        return None

    if isinstance(value, float) and math.isnan(value):
        return None

    s = str(value).strip()
    if not s:
        return None

    if s.lower() == "nan":
        return None

    s = unicodedata.normalize("NFKC", s)
    s = s.lower()
    # Collapse internal whitespace for robustness
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_value_key(value):
    """
    Normalization for matching VALUES robustly:
    - normalize_value_str
    - then drop all non-alphanumeric characters (spaces, hyphens, punctuation)
    This allows matching variants like "VVS 1" == "VVS1", "IF-VVS1" == "IF VVS1", etc.
    """
    s = normalize_value_str(value)
    if s is None:
        return None
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s or None


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
# SUPPLIER FILE LOADING (CSV/XLSX)
# ------------------------------------------------------------

def _score_columns_against_header_map(columns, header_map):
    """
    Score how likely a header row is by counting matches against known header variants.
    """
    if not header_map:
        return 0

    score = 0
    for col in list(columns):
        norm = normalize_header_name(col)
        if norm and norm in header_map:
            score += 1
    return score


def _unnamed_column_count(columns):
    count = 0
    for c in list(columns):
        s = str(c) if c is not None else ""
        if s.startswith("Unnamed:"):
            count += 1
    return count


def _looks_like_xlsx(file_bytes):
    """
    XLSX files are ZIP containers; they typically start with 'PK'.
    This quickly detects common user mistakes like renaming a CSV to .xlsx.
    """
    if not file_bytes or len(file_bytes) < 2:
        return False
    return file_bytes[:2] == b"PK"


def _parse_number(value):
    """
    Parse a numeric-ish field robustly (handles commas, currency symbols, text).
    Returns float or None.
    """
    if is_empty_value(value):
        return None
    s = str(value).strip()
    s = s.replace(",", "")
    # keep digits, dot, and minus
    s = re.sub(r"[^\d.\-]", "", s)
    if not s or s in {"-", ".", "-."}:
        return None
    try:
        return float(s)
    except Exception:
        return None


def _color_is_satisfied(row):
    """
    If 'color' is missing, treat it as satisfied when Fancy Color fields are present.
    """
    color_val = row.get("color", None)
    if not is_empty_value(color_val):
        return True

    fancy_color = row.get("fancy_color_dominant_color", None)
    fancy_intensity = row.get("fancy_color_intensity", None)
    return (not is_empty_value(fancy_color)) or (not is_empty_value(fancy_intensity))


def _get_url_extension(url):
    """
    Return lowercase extension from URL path, without the dot. None if absent.
    """
    try:
        p = urlparse(str(url).strip())
        path = p.path or ""
        m = re.search(r"\.([A-Za-z0-9]+)$", path)
        if not m:
            return None
        return m.group(1).lower()
    except Exception:
        return None


def detect_best_excel_layout(file_bytes, header_map=None, sheet_name=None, max_header_row=6):
    """
    Detect best (sheet_name, header_row) for an uploaded XLSX by:
    - trying multiple sheets (unless sheet_name provided)
    - trying header rows 0..max_header_row-1
    - picking the combo with the most header_map matches, then fewest Unnamed columns.
    """
    if not file_bytes:
        raise ValueError("The uploaded `.xlsx` file is empty.")

    if not _looks_like_xlsx(file_bytes):
        raise ValueError(
            "This file does not look like a real `.xlsx` workbook (it is not a ZIP-based Excel file). "
            "If you started from a CSV, either upload the `.csv` directly or re-save/export as an **Excel Workbook (.xlsx)** "
            "from Excel/Sheets (not by renaming the file extension)."
        )

    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        raise ValueError(
            "Unable to read the uploaded Excel file. Please ensure it is a valid, non-password-protected `.xlsx` "
            f"(not `.xls` or a renamed non-Excel file). Root cause: {type(e).__name__}: {e}"
        ) from e

    if not getattr(xls, "sheet_names", None):
        raise ValueError(
            "This Excel workbook contains 0 worksheets. Please re-export or resave it as a standard `.xlsx` file."
        )
    sheets = [sheet_name] if sheet_name else list(xls.sheet_names)
    header_rows = range(0, max_header_row)

    best = None  # (score, -unnamed, -num_cols, sheet, header_row)
    for sh in sheets:
        for hr in header_rows:
            try:
                probe = pd.read_excel(xls, sheet_name=sh, header=hr, nrows=0)
            except Exception:
                continue
            cols = list(probe.columns)
            score = _score_columns_against_header_map(cols, header_map)
            unnamed = _unnamed_column_count(cols)
            num_cols = len(cols)
            candidate = (score, -unnamed, -num_cols, sh, hr)
            if best is None or candidate > best:
                best = candidate

    if best is None:
        # Fallback: first sheet, first row
        return (xls.sheet_names[0], 0)

    return (best[3], best[4])


def load_supplier_bytes(file_bytes, filename, header_map=None, sheet_name=None):
    """
    Load supplier inventory file content (bytes) into a DataFrame.
    Returns: (df, meta)
    """
    ext = filename.rsplit(".", 1)[-1].lower() if filename and "." in filename else ""
    meta = {"ext": ext, "sheet_name": None, "header_row": None}

    if ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes))
        return df, meta

    if ext == "xlsx":
        try:
            best_sheet, best_header = detect_best_excel_layout(
                file_bytes=file_bytes,
                header_map=header_map,
                sheet_name=sheet_name,
            )
            meta["sheet_name"] = best_sheet
            meta["header_row"] = best_header
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=best_sheet,
                header=best_header,
                engine="openpyxl",
            )
            return df, meta
        except Exception as e:
            raise ValueError(
                "Failed to load the uploaded `.xlsx`. "
                "Common causes: corrupted file, password protection, or the file is not a real `.xlsx` workbook. "
                f"Root cause: {type(e).__name__}: {e}"
            ) from e

    # Fallback: let pandas attempt.
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        return df, meta
    except Exception as e:
        raise ValueError("Failed to load the uploaded file.") from e


# ------------------------------------------------------------
# LOAD HEADER RULES (Columns sheet)
# ------------------------------------------------------------

def load_header_rules(rules_source):
    """
    Uses Columns sheet: Column Name, Column Values (comma-separated synonyms)
    Returns: header_map, canonical_set
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
    Uses Values sheet: Value Type, Base Value, Value Variations
    Returns: rules dict with wildcard and allowed values
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

        if vtype_norm not in rules:
            rules[vtype_norm] = {"wildcard": False, "allowed": set(), "allowed_keys": set()}

        if base_norm == "any" or "any" in variations_norm:
            rules[vtype_norm]["wildcard"] = True
            continue

        if base_norm:
            rules[vtype_norm]["allowed"].add(base_norm)
            base_key = normalize_value_key(base_norm)
            if base_key:
                rules[vtype_norm]["allowed_keys"].add(base_key)

        for nm in variations_norm:
            if nm != "any":
                rules[vtype_norm]["allowed"].add(nm)
                nm_key = normalize_value_key(nm)
                if nm_key:
                    rules[vtype_norm]["allowed_keys"].add(nm_key)

    return rules


# ------------------------------------------------------------
# LOAD SUPPLIER FILE
# ------------------------------------------------------------

def load_supplier(path, header_map=None, sheet_name=None):
    ext = path.split(".")[-1].lower()
    if ext == "csv":
        return pd.read_csv(path)
    if ext == "xlsx" and header_map:
        with open(path, "rb") as f:
            b = f.read()
        df, _meta = load_supplier_bytes(b, filename=path, header_map=header_map, sheet_name=sheet_name)
        return df
    if ext == "xlsx":
        return pd.read_excel(path, engine="openpyxl", sheet_name=sheet_name if sheet_name else 0)
    return pd.read_excel(path, sheet_name=sheet_name if sheet_name else 0)


# ------------------------------------------------------------
# APPLY HEADER NORMALIZATION
# ------------------------------------------------------------

def normalize_headers(df, header_map):
    """
    Map supplier headers -> canonical headers using header_map.
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
    Check for invalid numeric values:
    - carat/weight: must be > 0
    - price fields: must be > 0
    """
    invalid = []
    
    weight_cols = [
        c for c in df.columns
        if normalize_header_name(c) in {"carat", "weight", "carat_weight", "size"}
    ]
    for col in weight_cols:
        for idx, val in df[col].items():
            if is_empty_value(val):
                continue
            
            num_val = _parse_number(val)
            if num_val is None:
                continue
            if num_val <= 0:
                invalid.append(
                    f"Row {idx + 2}: Invalid carat value '{val}' in column '{col}' (must be > 0)"
                )
    
    price_cols = [
        c for c in df.columns
        if normalize_header_name(c) in {"price_per_carat", "total_sales_price"}
    ]
    for col in price_cols:
        for idx, val in df[col].items():
            if is_empty_value(val):
                continue
            
            num_val = _parse_number(val)
            if num_val is None:
                continue
            if num_val <= 0:
                invalid.append(
                    f"Row {idx + 2}: Invalid price '{val}' in column '{col}' (must be > 0)"
                )
    
    return invalid


# ------------------------------------------------------------
# VALUE VALIDATION
# ------------------------------------------------------------

def check_values(df, value_rules):
    """
    Strict value validation against allowed values list
    """
    invalid = []

    for col in df.columns:
        vtype_norm = normalize_header_name(col)

        if vtype_norm not in value_rules:
            continue

        rule = value_rules[vtype_norm]
        if rule["wildcard"]:
            continue

        allowed = rule["allowed"]
        allowed_keys = rule.get("allowed_keys") or set()

        for idx, val in df[col].items():
            norm_val = normalize_value_str(val)
            if norm_val is None:
                continue

            # Robust matching: compare both raw normalized and key-normalized forms.
            if norm_val in allowed:
                continue

            key = normalize_value_key(norm_val)
            if key and key in allowed_keys:
                continue

            invalid.append(f"Row {idx + 2}: Invalid '{val}' in column '{col}'")

    return invalid


# ------------------------------------------------------------
# MANDATORY FIELDS
# ------------------------------------------------------------

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
    Check mandatory fields are non-empty
    """
    missing = []
    for idx, row in df.iterrows():
        missing_cols = []
        for col in MANDATORY_COLS:
            if col in df.columns:
                if col == "color":
                    if not _color_is_satisfied(row):
                        missing_cols.append(col)
                else:
                    if is_empty_value(row[col]):
                        missing_cols.append(col)
            else:
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
    Fast parallel URL checker
    """
    url_cols = [c for c in df.columns if "url" in c.lower()]
    bad = []
    tasks = []

    with ThreadPoolExecutor(max_workers=30) as executor:
        for idx, row in df.iterrows():
            for col in url_cols:
                url = row[col] if col in df.columns else None
                # Certificate URLs: accept only PDF/JPG/JPEG
                if col.lower() == "cert_url_1" and not is_empty_value(url):
                    ext = _get_url_extension(url)
                    if ext not in {"pdf", "jpg", "jpeg"}:
                        bad.append(
                            f"Row {idx + 2}: {col} → UNACCEPTABLE FORMAT (expected .pdf/.jpg) → URL: {url}"
                        )
                        continue

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
    parser.add_argument("--sheet", default=None, help="Excel sheet name (XLSX only)")
    args = parser.parse_args()

    print("Loading rules...")
    header_map, canonical = load_header_rules(args.rules)
    value_rules = load_value_rules(args.rules)

    print("Loading supplier file...")
    df = load_supplier(args.supplier, header_map=header_map, sheet_name=args.sheet)

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
