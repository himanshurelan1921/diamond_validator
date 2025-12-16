import streamlit as st
import pandas as pd
import io
import re
import math
from collections import Counter
import validator  # our backend module
import os

# ------------------------------------------------------------
# STREAMLIT SETUP
# ------------------------------------------------------------

st.set_page_config(page_title="Diamond Inventory Validator", layout="wide")
st.title("Diamond & Lab-Grown Inventory Validator")
st.markdown("""
Upload your **supplier inventory file** (.csv, .xlsx, or .json) to validate against the internal rule set.
""")

# ------------------------------------------------------------
# SMALL HELPERS
# ------------------------------------------------------------

def find_canonical_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def sanitize_sheet_name(name):
    """Removes characters invalid for Excel sheet names: \ / ? * [ ] :"""
    name = re.sub(r'[\\/?*\[\]:]', '', name)
    return name[:31]

def build_mandatory_issues(df):
    issues = []
    missing_by_col = Counter()
    mandatory_cols = getattr(validator, "MANDATORY_COLS", [])
    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        for col in mandatory_cols:
            if col in df.columns:
                v_norm = validator.normalize_value_str(row[col])
                if v_norm is None:
                    issues.append({
                        "Category": "Missing Mandatory",
                        "Stock No.": stock,
                        "Issue Type": "Missing Value",
                        "Column": col,
                        "Value": row[col],
                        "Details": "Missing mandatory field",
                        "Row": idx + 2,
                    })
                    missing_by_col[col] += 1
    return issues, missing_by_col

def parse_invalid_value_strings(invalid_list, df):
    issues = []
    invalid_shape_values = set()
    pattern = re.compile(r"Row (\d+): Invalid '(.*)' in column '([^']+)'")
    for msg in invalid_list:
        m = pattern.match(msg)
        if not m: continue
        row_num, value, column = int(m.group(1)), m.group(2), m.group(3)
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        stock = df.iloc[data_idx].get("stock_num", None)
        issues.append({
            "Category": "Invalid Value",
            "Stock No.": stock,
            "Issue Type": "Invalid Value",
            "Column": column,
            "Value": value,
            "Details": "Value not in accepted list",
            "Row": row_num,
        })
        if column == "shape": invalid_shape_values.add(value)
    return issues, sorted(invalid_shape_values)

def parse_range_issue_strings(range_list, df):
    issues = []
    pattern = re.compile(r"Row (\d+): Out of Range '(.*)' in column '([^']+)'")
    for msg in range_list:
        m = pattern.match(msg)
        if not m: continue
        row_num, value, column = int(m.group(1)), m.group(2), m.group(3)
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        stock = df.iloc[data_idx].get("stock_num", None)
        range_details = validator.RANGE_COLS.get(column.lower(), (0, 100, "Value"))
        min_val, max_val, name = range_details
        issues.append({
            "Category": "Range Issue",
            "Stock No.": stock,
            "Issue Type": "Value Out of Range",
            "Column": column,
            "Value": value,
            "Details": f"{name} must be between {min_val} and {max_val}",
            "Row": row_num,
        })
    return issues

def parse_url_issue_strings(url_list, df):
    issues = []
    pattern = re.compile(r"Row (\d+): ([^ ]+) → (.+)")
    counts = Counter()
    for msg in url_list:
        m = pattern.match(msg)
        if not m: continue
        row_num, col, status = int(m.group(1)), m.group(2), m.group(3)
        if col == "cert_url_1" and "NOT PROVIDED" not in status: continue
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        stock = df.iloc[data_idx].get("stock_num", None)
        issues.append({
            "Category": "URL Issue", "Stock No.": stock, "Issue Type": "URL Error",
            "Column": col, "Value": df.iloc[data_idx].get(col, None), "Details": status, "Row": row_num,
        })
        if "NOT PROVIDED" in status: counts[f"missing_{col.split('_')[0]}"] += 1
        else: counts[f"bad_{col.split('_')[0]}"] += 1
    return issues, counts

def build_excel_report(structured_issues):
    section_map = {
        "shape": "1. Shape", "weight": "2. Weight", "carat": "2. Weight", "carat_weight": "2. Weight",
        "color": "3. Color", "clarity": "4. Clarity", "image_url_1": "5. Image URL",
        "video_url_1": "6. Video URL", "cert_url_1": "7. Certificate URL",
        "price_per_carat": "8. Price", "total_sales_price": "8. Price"
    }
    other_sheet, range_sheet = "9. Other Issues", "10. Range Issues"
    issues_by_sheet = {s: [] for s in list(section_map.values()) + [other_sheet, range_sheet]}
    
    for issue in structured_issues:
        if issue["Category"] == "Range Issue": sheet_name = range_sheet
        elif issue["Issue Type"] == "Price Mismatch": sheet_name = "8. Price"
        elif issue["Issue Type"] == "Cert URL Format": sheet_name = "7. Certificate URL"
        else: sheet_name = section_map.get(issue["Column"].lower(), other_sheet)
        issues_by_sheet[sheet_name].append(issue)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name in sorted(issues_by_sheet.keys()):
            if issues_by_sheet[sheet_name]:
                safe_name = sanitize_sheet_name(sheet_name)
                df_report = pd.DataFrame(issues_by_sheet[sheet_name]).drop(columns=["Category"], errors='ignore')
                df_report.to_excel(writer, sheet_name=safe_name, index=False)
    buffer.seek(0)
    return buffer

def build_email_body(supplier_name, inv_shapes, miss_col, inv_col, url_counts, cut_miss, pdf_miss, price_miss):
    lines = [f"Hi {supplier_name},", "", "Hope you're doing well.", "", 
             "During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:", ""]
    
    sections = [
        (inv_shapes or miss_col.get("shape"), "1. Shape", f"- Shape is missing or invalid."),
        (miss_col.get("carat") or inv_col.get("carat"), "2. Weight", f"- Weight issues identified."),
        (miss_col.get("color") or inv_col.get("color"), "3. Color", f"- Color issues identified."),
        (miss_col.get("clarity") or inv_col.get("clarity"), "4. Clarity", f"- Clarity issues identified."),
        (url_counts['missing_image'] or url_counts['bad_image'], "5. Image URLs", f"- Image URL issues identified."),
        (pdf_miss or miss_col.get("cert_url_1"), "7. Certificate URLs", f"- Certificate must be direct PDF links."),
        (price_miss, "8. Price", f"- Pricing mismatches found (Carat x Price per Carat).")
    ]
    for condition, title, text in sections:
        if condition: lines.extend([title, text, ""])

    lines.extend(["A spreadsheet outlining the above items has been attached. We appreciate your corrections.", "", "Best Regards,", "Himanshu", "VDB Marketplace Support"])
    return "\n".join(lines)

# ------------------------------------------------------------
# MAIN FLOW
# ------------------------------------------------------------

supplier_file = st.file_uploader("Upload Supplier Inventory (.csv, .xlsx, or .json)", type=["csv", "xlsx", "json"])
supplier_name = st.text_input("Supplier Name", value="Supplier")

if st.button("Run Validation") and supplier_file:
    rules_path = "headers.xlsx"
    header_map, _ = validator.load_header_rules(rules_path)
    value_rules = validator.load_value_rules(rules_path)
    
    ext = supplier_file.name.split(".")[-1].lower()
    if ext == "csv": df = pd.read_csv(supplier_file)
    elif ext == "json": df = pd.read_json(supplier_file)
    else: df = pd.read_excel(supplier_file)

    df, unknown = validator.normalize_headers(df, header_map)
    miss_str = validator.check_mandatory(df)
    mand_issues, miss_col = build_mandatory_issues(df)
    inv_str = validator.check_values(df, value_rules)
    inv_issues, inv_shapes = parse_invalid_value_strings(inv_str, df)
    range_str = validator.check_ranges(df)
    range_issues = parse_range_issue_strings(range_str, df)
    url_str = validator.check_all_urls(df)
    url_issues, url_counts = parse_url_issue_strings(url_str, df)
    
    # Special checks
    cert_issues, pdf_count = validator.find_non_pdf_cert_urls(df)
    price_issues, price_count = validator.build_price_mismatch_issues(df)

    structured_issues = mand_issues + inv_issues + range_issues + url_issues + cert_issues + price_issues
    st.download_button("📥 Download Report", build_excel_report(structured_issues), "validation_report.xlsx")
    st.text_area("Email Summary", build_email_body(supplier_name, inv_shapes, miss_col, Counter(), url_counts, 0, pdf_count, price_count), height=300)
