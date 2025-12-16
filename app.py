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
    """
    Return the first column name from `candidates` that exists in df.columns.
    """
    for c in candidates:
        if c in df.columns:
            return c
    return None

def sanitize_sheet_name(name):
    """
    Removes characters invalid for Excel sheet names.
    Invalid characters: \ / ? * [ ] :
    Max length: 31 characters.
    """
    name = re.sub(r'[\\/?*\[\]:]', '', name)
    return name[:31]

def build_mandatory_issues(df):
    """
    Return structured list of missing mandatory field issues AND a count of missing columns.
    """
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
    """
    Parse strings like:
      "Row 9: Invalid 'FNT BL' in column 'fluor_intensity'"
    into structured records.
    Also collect invalid SHAPE values for email.
    """
    issues = []
    invalid_shape_values = set()

    pattern = re.compile(r"Row (\d+): Invalid '(.*)' in column '([^']+)'")

    for msg in invalid_list:
        m = pattern.match(msg)
        if not m:
            continue
        row_num = int(m.group(1))
        value = m.group(2)
        column = m.group(3)

        data_idx = row_num - 2  # header row + 1-based row index
        if data_idx < 0 or data_idx >= len(df):
            continue
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
    """
    Parse strings like:
      "Row 9: Out of Range '95.0' in column 'table'"
    into structured records.
    """
    issues = []
    pattern = re.compile(r"Row (\d+): Out of Range '(.*)' in column '([^']+)'")

    for msg in range_list:
        m = pattern.match(msg)
        if not m:
            continue
        row_num = int(m.group(1))
        value = m.group(2)
        column = m.group(3)

        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df):
            continue
        row = df.iloc[data_idx]
        stock = row.get("stock_num", None)
        
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
    """
    Parse strings like:
      "Row 2: image_url_1 → NOT WORKING"
      "Row 3: cert_url_1 → NOT PROVIDED"
    into structured records and counts.
    """
    issues = []
    pattern = re.compile(r"Row (\d+): ([^ ]+) → (.+)")

    missing_image = 0
    missing_video = 0
    bad_video = 0
    bad_image = 0
    bad_cert = 0

    for msg in url_list:
        m = pattern.match(msg)
        if not m:
            continue
        row_num = int(m.group(1))
        col = m.group(2)
        status = m.group(3)

        if col == "cert_url_1" and "NOT PROVIDED" not in status:
            continue

        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df):
            continue
        row = df.iloc[data_idx]
        stock = row.get("stock_num", None)
        url_value = row.get(col, None)

        if "NOT PROVIDED" in status:
            issue_type = "Missing URL"
        else:
            issue_type = "URL Error"

        issues.append({
            "Category": "URL Issue",
            "Stock No.": stock,
            "Issue Type": issue_type,
            "Column": col,
            "Value": url_value,
            "Details": status,
            "Row": row_num,
        })

        if col == "image_url_1":
            if "NOT PROVIDED" in status:
                missing_image += 1
            else:
                bad_image += 1
        elif col == "video_url_1":
            if "NOT PROVIDED" in status:
                missing_video += 1
            else:
                bad_video += 1
        elif col == "cert_url_1":
            if "NOT PROVIDED" in status:
                bad_cert += 1 
            
    counts = {
        "missing_image": missing_image,
        "missing_video": missing_video,
        "bad_video": bad_video,
        "bad_image": bad_image,
        "bad_cert": bad_cert,
    }
    return issues, counts


def find_missing_cut_grade(df):
    issues = []
    count = 0
    cut_cols = [c for c in ["cut_grade", "cut"] if c in df.columns]
    if not cut_cols:
        return issues, 0

    col = cut_cols[0]
    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        v_norm = validator.normalize_value_str(row[col])
        if v_norm is None:
            count += 1
            issues.append({
                "Category": "Missing Value",
                "Stock No.": stock,
                "Issue Type": "Missing Value",
                "Column": col,
                "Value": row[col],
                "Details": "Missing cut grade",
                "Row": idx + 2,
            })
    return issues, count


def find_non_pdf_cert_urls(df):
    issues = []
    count = 0
    if "cert_url_1" not in df.columns:
        return issues, 0
    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        url = row.get("cert_url_1", None)
        if validator.normalize_value_str(url) is None:
            continue
        if ".pdf" not in str(url).lower():
            count += 1
            issues.append({
                "Category": "URL Issue",
                "Stock No.": stock,
                "Issue Type": "Cert URL Format",
                "Column": "cert_url_1",
                "Value": url,
                "Details": "Certificate URL is not a direct PDF link",
                "Row": idx + 2,
            })
    return issues, count


def build_price_mismatch_issues(df):
    issues = []
    count = 0
    weight_col = find_canonical_col(df, ["carat", "weight", "carat_weight"])
    ppc_col = find_canonical_col(df, ["price_per_carat"])
    tsp_col = find_canonical_col(df, ["total_sales_price"])
    if not (weight_col and ppc_col and tsp_col):
        return issues, 0

    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        w = validator.to_float(row.get(weight_col, None))
        ppc = validator.to_float(row.get(ppc_col, None))
        tsp = validator.to_float(row.get(tsp_col, None))
        if w is None or ppc is None or tsp is None:
            continue
        expected = round(w * ppc, 2)
        if abs(expected - tsp) > 0.01:
            count += 1
            issues.append({
                "Category": "Price Issue",
                "Stock No.": stock,
                "Issue Type": "Price Mismatch",
                "Column": tsp_col,
                "Value": row.get(tsp_col, None),
                "Details": f"Expected {expected} = {w} * {ppc}, got {tsp}",
                "Row": idx + 2,
            })
    return issues, count


def build_excel_report(structured_issues):
    section_map = {
        "shape": "1. Shape",
        "weight": "2. Weight",
        "carat": "2. Weight",
        "carat_weight": "2. Weight",
        "color": "3. Color",
        "clarity": "4. Clarity",
        "image_url_1": "5. Image URL",
        "video_url_1": "6. Video URL",
        "cert_url_1": "7. Certificate URL",
        "price_per_carat": "8. Price",
        "total_sales_price": "8. Price",
    }
    other_sheet = "9. Other Issues / Cut Grade"
    range_sheet = "10. Range Issues"
    
    issues_by_sheet = {s: [] for s in list(section_map.values()) + [other_sheet, range_sheet]}
        
    for issue in structured_issues:
        column = issue.get("Column", "").lower()
        if issue["Category"] == "Range Issue":
            sheet_name = range_sheet
        elif issue["Issue Type"] == "Price Mismatch":
             sheet_name = "8. Price"
        elif issue["Issue Type"] == "Cert URL Format":
             sheet_name = "7. Certificate URL"
        else:
            sheet_name = section_map.get(column, other_sheet)
        issues_by_sheet[sheet_name].append(issue)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name in sorted(issues_by_sheet.keys()):
            if issues_by_sheet[sheet_name]:
                safe_sheet_name = sanitize_sheet_name(sheet_name)
                df_report = pd.DataFrame(issues_by_sheet[sheet_name]).drop(columns=["Category"], errors='ignore')
                df_report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    buffer.seek(0)
    return buffer


def build_email_body(supplier_name, invalid_shape_values, missing_by_col, invalid_by_col, url_counts, cut_missing_count, non_pdf_cert_count, price_mismatch_count):
    lines = [f"Hi {supplier_name},", "", "Hope you're doing well.", "", "During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:", ""]

    if invalid_shape_values or missing_by_col.get("shape") or invalid_by_col.get("shape"):
        lines.append("1. Shape")
        if missing_by_col.get("shape"):
            lines.append(f"- Shape is missing for {missing_by_col['shape']} item(s).")
        if invalid_shape_values:
            lines.append("- We found invalid shape values that do not match VDB’s standardised shape list, for example:")
            for sh in invalid_shape_values:
                lines.append(f"  • {sh}")
        lines.append("")

    weight_col = None
    for cand in ["carat", "weight", "carat_weight"]:
        if cand in missing_by_col or cand in invalid_by_col:
            weight_col = cand
            break
    if weight_col:
        lines.append("2. Weight")
        if missing_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) is missing for {missing_by_col[weight_col]} item(s).")
        if invalid_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) has invalid values for {invalid_by_col[weight_col]} item(s).")
        lines.append("")

    if missing_by_col.get("color") or invalid_by_col.get("color"):
        lines.append("3. Color")
        if missing_by_col.get("color"): lines.append(f"- Color is missing for {missing_by_col['color']} item(s).")
        if invalid_by_col.get("color"): lines.append(f"- Color has invalid values for {invalid_by_col['color']} item(s).")
        lines.append("")

    if missing_by_col.get("clarity") or invalid_by_col.get("clarity"):
        lines.append("4. Clarity")
        if missing_by_col.get("clarity"): lines.append(f"- Clarity is missing for {missing_by_col['clarity']} item(s).")
        if invalid_by_col.get("clarity"): lines.append(f"- Clarity has invalid values for {invalid_by_col['clarity']} item(s).")
        lines.append("")

    if missing_by_col.get("image_url_1") or url_counts.get("missing_image") or url_counts.get("bad_image"):
        lines.append("5. Image URLs")
        total_miss = missing_by_col.get("image_url_1", 0) + url_counts.get("missing_image", 0)
        if total_miss: lines.append(f"- Image URLs are missing for {total_miss} item(s).")
        if url_counts.get("bad_image"): lines.append(f"- {url_counts['bad_image']} image URL(s) are not working.")
        lines.append("")

    if missing_by_col.get("video_url_1") or url_counts.get("missing_video") or url_counts.get("bad_video"):
        lines.append("6. Video URLs")
        total_miss = missing_by_col.get("video_url_1", 0) + url_counts.get("missing_video", 0)
        if total_miss: lines.append(f"- Video URLs are missing for {total_miss} item(s).")
        if url_counts.get("bad_video"): lines.append(f"- {url_counts['bad_video']} video URL(s) are not working.")
        lines.append("")

    if missing_by_col.get("cert_url_1") or url_counts.get("bad_cert") or non_pdf_cert_count:
        lines.append("7. Certificate URLs")
        total_miss = missing_by_col.get("cert_url_1", 0) + url_counts.get("bad_cert", 0)
        if total_miss: lines.append(f"- Certificate URLs are missing for {total_miss} item(s).")
        if non_pdf_cert_count: lines.append(f"- {non_pdf_cert_count} certificate URL(s) do not appear to be direct PDF links.")
        lines.append("")

    if price_mismatch_count or missing_by_col.get("price_per_carat"):
        lines.append("8. Price")
        if price_mismatch_count: lines.append(f"- For {price_mismatch_count} item(s), Price does not match.")
        lines.append("")

    lines.append("A spreadsheet outlining the above items has been attached. Best Regards, Himanshu, VDB Marketplace Support")
    return "\n".join(lines)


# ------------------------------------------------------------
# MAIN UI
# ------------------------------------------------------------

supplier_file = st.file_uploader("Upload Supplier Inventory (.csv, .xlsx, or .json)", type=["csv", "xlsx", "json"])
supplier_name = st.text_input("Supplier Name (for email)", value="Supplier")

if st.button("Run Validation"):
    if not supplier_file:
        st.error("⚠ Please upload the Supplier Inventory file.")
        st.stop()
    rules_path = "headers.xlsx"
    st.info("📘 Loading rules…")
    header_map, _ = validator.load_header_rules(rules_path)
    value_rules = validator.load_value_rules(rules_path)
    
    st.info("📄 Loading supplier inventory…")
    ext = supplier_file.name.split(".")[-1].lower()
    if ext == "csv":
        df = pd.read_csv(supplier_file)
    elif ext == "json":
        df = pd.read_json(supplier_file)
    else:
        df = pd.read_excel(supplier_file)
    st.success(f"Supplier file loaded: **{len(df)} rows**")

    # STEP 1
    df, unknown_headers = validator.normalize_headers(df, header_map)
    # STEP 2
    mandatory_issues, missing_by_col = build_mandatory_issues(df)
    # STEP 3
    invalid_issues, invalid_shape_values = parse_invalid_value_strings(validator.check_values(df, value_rules), df)
    # STEP 4
    range_issues = parse_range_issue_strings(validator.check_ranges(df), df)
    # STEP 5
    url_issues_struct, url_counts = parse_url_issue_strings(validator.check_all_urls(df), df)
    # STEP 6
    cut_issues, cut_missing_count = find_missing_cut_grade(df)
    cert_format_issues, non_pdf_cert_count = find_non_pdf_cert_urls(df) 
    price_issues, price_mismatch_count = build_price_mismatch_issues(df)

    st.success("Validation completed!")

    structured_issues = mandatory_issues + invalid_issues + range_issues + url_issues_struct + cut_issues + cert_format_issues + price_issues
    st.download_button("📥 Download Report", build_excel_report(structured_issues), "validation_report.xlsx")
    
    email_body = build_email_body(supplier_name, invalid_shape_values, missing_by_col, Counter(), url_counts, cut_missing_count, non_pdf_cert_count, price_mismatch_count)
    st.text_area("Email Summary", value=email_body, height=400)
