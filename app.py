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
Upload your **supplier inventory file** (.csv or .xlsx) to validate against the internal rule set.
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
        row_num = int(m.group(1))
        value = m.group(2)
        column = m.group(3)
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        row = df.iloc[data_idx]
        stock = row.get("stock_num", None)
        issues.append({
            "Category": "Invalid Value",
            "Stock No.": stock,
            "Issue Type": "Invalid Value",
            "Column": column,
            "Value": value,
            "Details": "Value not in accepted list",
            "Row": row_num,
        })
        if column == "shape":
            invalid_shape_values.add(value)
    return issues, sorted(invalid_shape_values)

def parse_range_issue_strings(range_list, df):
    issues = []
    pattern = re.compile(r"Row (\d+): Out of Range '(.*)' in column '([^']+)'")
    for msg in range_list:
        m = pattern.match(msg)
        if not m: continue
        row_num = int(m.group(1))
        value = m.group(2)
        column = m.group(3)
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        row = df.iloc[data_idx]
        stock = row.get("stock_num", None)
        range_details = validator.RANGE_COLS.get(column.lower(), (0, 100, "Value"))
        min_val, max_val, name = range_details
        issues.append({
            "Category": "Range Issue", "Stock No.": stock, "Issue Type": "Value Out of Range",
            "Column": column, "Value": value, "Details": f"{name} must be {min_val}-{max_val}", "Row": row_num,
        })
    return issues

def parse_url_issue_strings(url_list, df):
    issues = []
    pattern = re.compile(r"Row (\d+): ([^ ]+) → (.+)")
    missing_image, missing_video, bad_video, bad_image, bad_cert = 0, 0, 0, 0, 0
    for msg in url_list:
        m = pattern.match(msg)
        if not m: continue
        row_num, col, status = int(m.group(1)), m.group(2), m.group(3)
        if col == "cert_url_1" and "NOT PROVIDED" not in status: continue
        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df): continue
        row = df.iloc[data_idx]
        stock, url_value = row.get("stock_num", None), row.get(col, None)
        issue_type = "Missing URL" if "NOT PROVIDED" in status else "URL Error"
        issues.append({
            "Category": "URL Issue", "Stock No.": stock, "Issue Type": issue_type,
            "Column": col, "Value": url_value, "Details": status, "Row": row_num,
        })
        if col == "image_url_1":
            if "NOT PROVIDED" in status: missing_image += 1
            else: bad_image += 1
        elif col == "video_url_1":
            if "NOT PROVIDED" in status: missing_video += 1
            else: bad_video += 1
        elif col == "cert_url_1":
            if "NOT PROVIDED" in status: bad_cert += 1 
    return issues, {"missing_image": missing_image, "missing_video": missing_video, "bad_video": bad_video, "bad_image": bad_image, "bad_cert": bad_cert}

def build_excel_report(structured_issues):
    section_map = {
        "shape": "1. Shape", "weight": "2. Weight", "carat": "2. Weight", "carat_weight": "2. Weight",
        "color": "3. Color", "clarity": "4. Clarity", "image_url_1": "5. Image URL",
        "video_url_1": "6. Video URL", "cert_url_1": "7. Certificate URL",
        "price_per_carat": "8. Price", "total_sales_price": "8. Price",
    }
    other_sheet, range_sheet = "9. Other Issues / Cut Grade", "10. Range Issues"
    issues_by_sheet = {s: [] for s in list(section_map.values()) + [other_sheet, range_sheet]}
    for issue in structured_issues:
        col = issue.get("Column", "").lower()
        if issue["Category"] == "Range Issue": sheet_name = range_sheet
        elif issue["Issue Type"] == "Price Mismatch": sheet_name = "8. Price"
        elif issue["Issue Type"] == "Cert URL Format": sheet_name = "7. Certificate URL"
        else: sheet_name = section_map.get(col, other_sheet)
        issues_by_sheet[sheet_name].append(issue)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name in sorted(issues_by_sheet.keys()):
            if issues_by_sheet[sheet_name]:
                df_report = pd.DataFrame(issues_by_sheet[sheet_name]).drop(columns=["Category"], errors='ignore')
                df_report.to_excel(writer, sheet_name=sanitize_sheet_name(sheet_name), index=False)
    buffer.seek(0)
    return buffer

def build_email_body(supplier_name, invalid_shape_values, missing_by_col, url_counts, pdf_miss, price_miss):
    lines = [f"Hi {supplier_name},", "", "Hope you're doing well.", "", "During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:", ""]
    lines.append("--------------------------------------------------")
    
    if invalid_shape_values or missing_by_col.get("shape"):
        lines.append("SHAPE ISSUES:")
        if missing_by_col.get("shape"): lines.append(f" - Shape is missing for {missing_by_col['shape']} item(s).")
        if invalid_shape_values:
            lines.append(" - Invalid shape values found:")
            for sh in invalid_shape_values: lines.append(f"    • {sh}")
        lines.append("")

    if missing_by_col.get("color"):
        lines.append("COLOR ISSUES:")
        lines.append(f" - Color is missing for {missing_by_col['color']} item(s).")
        lines.append("")

    img_miss = missing_by_col.get("image_url_1", 0) + url_counts.get("missing_image", 0)
    if img_miss or url_counts.get("bad_image"):
        lines.append("IMAGE URL ISSUES:")
        if img_miss: lines.append(f" - Image URLs are missing for {img_miss} item(s).")
        if url_counts.get("bad_image"): lines.append(f" - {url_counts['bad_image']} image URL(s) are not working.")
        lines.append("")

    vid_miss = missing_by_col.get("video_url_1", 0) + url_counts.get("missing_video", 0)
    if vid_miss or url_counts.get("bad_video"):
        lines.append("VIDEO URL ISSUES:")
        if vid_miss: lines.append(f" - Video URLs are missing for {vid_miss} item(s).")
        if url_counts.get("bad_video"): lines.append(f" - {url_counts['bad_video']} video URL(s) are not working.")
        lines.append("")

    cert_miss = missing_by_col.get("cert_url_1", 0) + url_counts.get("bad_cert", 0)
    if cert_miss or pdf_miss:
        lines.append("CERTIFICATE URL ISSUES:")
        if cert_miss: lines.append(f" - Certificate URLs are missing for {cert_miss} item(s).")
        if pdf_miss: lines.append(f" - {pdf_miss} cert URL(s) are not direct PDF links.")
        lines.append("")

    lines.append("--------------------------------------------------")
    lines.append("")
    lines.append("A spreadsheet outlining the above items has been attached. We would appreciate it if you could make the necessary corrections at your earliest convenience.")
    lines.append("")
    lines.append("Best Regards,")
    lines.append("VDB Marketplace Support Team")
    return "\n".join(lines)

# ------------------------------------------------------------
# MAIN UI
# ------------------------------------------------------------

supplier_file = st.file_uploader("Upload Supplier Inventory (.csv or .xlsx)", type=["csv", "xlsx"])
supplier_name = st.text_input("Supplier Name", value="Supplier")

# Initialize Session State
if 'v_complete' not in st.session_state:
    st.session_state.v_complete = False
    st.session_state.s_issues = None
    st.session_state.e_body = ""

if st.button("Run Validation"):
    if not supplier_file:
        st.error("⚠ Please upload the Supplier Inventory file.")
        st.stop()
    
    rules_path = "headers.xlsx"
    header_map, _ = validator.load_header_rules(rules_path)
    value_rules = validator.load_value_rules(rules_path)
    
    ext = supplier_file.name.split(".")[-1].lower()
    if ext == "csv": df = pd.read_csv(supplier_file)
    else: df = pd.read_excel(supplier_file)

    df, _ = validator.normalize_headers(df, header_map)
    mand_issues, miss_col = build_mandatory_issues(df)
    inv_issues, inv_shapes = parse_invalid_value_strings(validator.check_values(df, value_rules), df)
    range_issues = parse_range_issue_strings(validator.check_ranges(df), df)
    url_issues, url_counts = parse_url_issue_strings(validator.check_all_urls(df), df)
    
    price_issues, price_miss = validator.build_price_mismatch_issues(df)
    cert_pdf_issues, pdf_miss = validator.find_non_pdf_cert_urls(df)

    st.session_state.s_issues = mand_issues + inv_issues + range_issues + url_issues + price_issues + cert_pdf_issues
    st.session_state.e_body = build_email_body(supplier_name, inv_shapes, miss_col, url_counts, pdf_miss, price_miss)
    st.session_state.v_complete = True
    st.success("Validation completed!")

# Persistent display section
if st.session_state.v_complete:
    st.download_button(
        label="📥 Download Report",
        data=build_excel_report(st.session_state.s_issues),
        file_name="validation_report.xlsx"
    )
    st.text_area("Email Summary", value=st.session_state.e_body, height=400)
