import streamlit as st
import pandas as pd
import io
import re
import math
from collections import Counter
import validator
import os

# ------------------------------------------------------------
# STREAMLIT SETUP
# ------------------------------------------------------------

st.set_page_config(page_title="Diamond Inventory Validator", layout="wide")
st.title("Diamond & Lab-Grown Inventory Validator")
st.markdown("""
Upload your **supplier inventory file** to validate against the internal rule set.
""")

# ------------------------------------------------------------
# HELPERS
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
                if validator.is_empty_value(row[col]):
                    issues.append({
                        "Category": "Missing Mandatory",
                        "Stock No.": stock if not validator.is_empty_value(stock) else f"Row {idx + 2}",
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
    invalid_color_values = set()
    invalid_by_col = Counter()

    pattern = re.compile(r"Row (\d+): Invalid '(.*)' in column '([^']+)'")

    for msg in invalid_list:
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
        if validator.is_empty_value(stock):
            stock = f"Row {row_num}"

        issues.append({
            "Category": "Invalid Value",
            "Stock No.": stock,
            "Issue Type": "Invalid Value",
            "Column": column,
            "Value": value,
            "Details": "Value not in accepted list",
            "Row": row_num,
        })
        
        invalid_by_col[column] += 1

        if column == "shape":
            invalid_shape_values.add(value)
        elif column == "color":
            invalid_color_values.add(value)

    return issues, sorted(invalid_shape_values), sorted(invalid_color_values), invalid_by_col

def parse_numeric_invalid_strings(numeric_list, df):
    issues = []
    invalid_by_col = Counter()

    pattern = re.compile(r"Row (\d+): Invalid (?:carat value|price) '(.*)' in column '([^']+)'")

    for msg in numeric_list:
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
        if validator.is_empty_value(stock):
            stock = f"Row {row_num}"

        if "carat" in column or "weight" in column:
            detail = "Carat/Weight must be greater than 0"
            category = "Invalid Value"
        else:
            detail = "Price must be greater than 0"
            category = "Price Issue"

        issues.append({
            "Category": category,
            "Stock No.": stock,
            "Issue Type": "Invalid Numeric Value",
            "Column": column,
            "Value": value,
            "Details": detail,
            "Row": row_num,
        })
        
        invalid_by_col[column] += 1

    return issues, invalid_by_col

def parse_url_issue_strings(url_list, df):
    issues = []
    pattern = re.compile(r"Row (\d+): ([^ ]+) → (.+?) → URL: (.+)")

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
        url_value = m.group(4)

        data_idx = row_num - 2
        if data_idx < 0 or data_idx >= len(df):
            continue
        row = df.iloc[data_idx]
        stock = row.get("stock_num", None)
        if validator.is_empty_value(stock):
            stock = f"Row {row_num}"

        if "NOT PROVIDED" in status:
            issue_type = "Missing URL"
        else:
            issue_type = "URL Error"

        issues.append({
            "Category": "URL Issue",
            "Stock No.": stock,
            "Issue Type": issue_type,
            "Column": col,
            "URL": url_value,
            "Status": status,
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
        if validator.is_empty_value(stock):
            stock = f"Row {idx + 2}"
        if validator.is_empty_value(row[col]):
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

def build_price_mismatch_issues(df):
    issues = []
    count = 0

    weight_col = find_canonical_col(df, ["carat", "weight", "carat_weight"])
    ppc_col = find_canonical_col(df, ["price_per_carat"])
    tsp_col = find_canonical_col(df, ["total_sales_price"])

    if not (weight_col and ppc_col and tsp_col):
        return issues, 0

    def to_float(x):
        if validator.is_empty_value(x):
            return None
        s = str(x).strip()
        s = s.replace(",", "")
        s = re.sub(r"[^\d.\-]", "", s)
        try:
            return float(s)
        except Exception:
            return None

    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        if validator.is_empty_value(stock):
            stock = f"Row {idx + 2}"
        w = to_float(row.get(weight_col, None))
        ppc = to_float(row.get(ppc_col, None))
        tsp = to_float(row.get(tsp_col, None))

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
    # Start numbering from 1 (removed "0. Stock Number")
    section_map = {
        "stock_num": "1. Stock Number",
        "shape": "2. Shape",
        "weight": "3. Weight",
        "carat": "3. Weight",
        "carat_weight": "3. Weight",
        "color": "4. Color",
        "clarity": "5. Clarity",
        "image_url_1": "6. Image URL",
        "video_url_1": "7. Video URL",
        "cert_url_1": "8. Certificate URL",
        "price_per_carat": "9. Price",
        "total_sales_price": "9. Price",
    }
    other_sheet = "10. Other Issues / Cut Grade"
    
    issues_by_sheet = {}
    
    all_sheet_names = set(section_map.values())
    all_sheet_names.add(other_sheet)
    for name in all_sheet_names:
        issues_by_sheet[name] = []
        
    for issue in structured_issues:
        column = issue.get("Column", "").lower()
        sheet_name = section_map.get(column, other_sheet)
        
        if issue["Issue Type"] == "Price Mismatch":
             sheet_name = "9. Price"
        elif issue["Column"] in ["cut", "cut_grade"]:
            sheet_name = other_sheet

        issues_by_sheet[sheet_name].append(issue)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name in sorted(issues_by_sheet.keys()):
            safe_sheet_name = sanitize_sheet_name(sheet_name)
            
            if issues_by_sheet[sheet_name]:
                df_report = pd.DataFrame(issues_by_sheet[sheet_name]).drop(columns=["Category"], errors='ignore')
                df_report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            
        if not structured_issues:
            df_empty = pd.DataFrame(columns=["Stock No.", "Issue Type", "Column", "Value", "Details", "Row"])
            df_empty.to_excel(writer, sheet_name="No Issues Found", index=False)
                
    buffer.seek(0)
    return buffer

def build_email_body(
    supplier_name,
    invalid_shape_values,
    invalid_color_values,
    missing_by_col,
    invalid_by_col,
    url_counts,
    cut_missing_count,
    price_mismatch_count,
    missing_stock_count,
):
    lines = []
    lines.append(f"Hi {supplier_name},")
    lines.append("")
    lines.append("Hope you're doing well.")
    lines.append("")
    lines.append("During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:")
    lines.append("")

    # Start numbering from 1
    section_num = 1

    if missing_stock_count > 0:
        lines.append(f"{section_num}. Stock Number")
        lines.append(f"- Stock number is missing for {missing_stock_count} item(s).")
        lines.append("")
        section_num += 1

    if invalid_shape_values or missing_by_col.get("shape") or invalid_by_col.get("shape"):
        lines.append(f"{section_num}. Shape")
        if missing_by_col.get("shape"):
            lines.append(f"- Shape is missing for {missing_by_col['shape']} item(s).")
        if invalid_shape_values:
            lines.append("- We found invalid shape values that do not match VDB's standardised shape list, for example:")
            for sh in invalid_shape_values:
                lines.append(f"  • {sh}")
        lines.append("")
        section_num += 1

    weight_col = None
    for cand in ["carat", "weight", "carat_weight"]:
        if cand in missing_by_col or cand in invalid_by_col:
            weight_col = cand
            break

    if weight_col and (missing_by_col.get(weight_col) or invalid_by_col.get(weight_col)):
        lines.append(f"{section_num}. Weight")
        if missing_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) is missing for {missing_by_col[weight_col]} item(s).")
        if invalid_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) has invalid values (zero, negative, or not in accepted format) for {invalid_by_col[weight_col]} item(s).")
        lines.append("")
        section_num += 1

    if missing_by_col.get("color") or invalid_by_col.get("color") or invalid_color_values:
        lines.append(f"{section_num}. Color")
        if missing_by_col.get("color"):
            lines.append(f"- Color is missing for {missing_by_col['color']} item(s).")
        if invalid_color_values:
            lines.append("- We found invalid color values that do not match VDB's standardised color list, for example:")
            for clr in invalid_color_values:
                lines.append(f"  • {clr}")
        lines.append("")
        section_num += 1

    if missing_by_col.get("clarity") or invalid_by_col.get("clarity"):
        lines.append(f"{section_num}. Clarity")
        if missing_by_col.get("clarity"):
            lines.append(f"- Clarity is missing for {missing_by_col['clarity']} item(s).")
        if invalid_by_col.get("clarity"):
            lines.append(f"- Clarity has invalid values for {invalid_by_col['clarity']} item(s).")
        lines.append("")
        section_num += 1

    missing_image = missing_by_col.get("image_url_1", 0) + url_counts.get("missing_image", 0)
    if missing_image or url_counts.get("bad_image", 0):
        lines.append(f"{section_num}. Image URLs")
        if missing_image:
            lines.append(f"- Image URLs are missing for {missing_image} item(s).")
        if url_counts.get("bad_image", 0):
            lines.append(f"- {url_counts['bad_image']} image URL(s) are not working (HTTP errors).")
        lines.append("")
        section_num += 1

    missing_video = missing_by_col.get("video_url_1", 0) + url_counts.get("missing_video", 0)
    if missing_video or url_counts.get("bad_video", 0):
        lines.append(f"{section_num}. Video URLs")
        if missing_video:
            lines.append(f"- Video URLs are missing for {missing_video} item(s).")
        if url_counts.get("bad_video", 0):
            lines.append(f"- {url_counts['bad_video']} video URL(s) are not working (HTTP errors).")
        lines.append("")
        section_num += 1

    cert_issue_present = (
        missing_by_col.get("cert_url_1", 0)
        or url_counts.get("bad_cert", 0)
    )
    if cert_issue_present:
        lines.append(f"{section_num}. Certificate URLs")
        if missing_by_col.get("cert_url_1", 0) + url_counts.get("bad_cert", 0) > 0:
             lines.append(f"- Certificate URLs are missing for {missing_by_col['cert_url_1'] + url_counts['bad_cert']} item(s).")
        lines.append("")
        section_num += 1

    price_issue_present = (
        missing_by_col.get("price_per_carat")
        or missing_by_col.get("total_sales_price")
        or invalid_by_col.get("price_per_carat")
        or invalid_by_col.get("total_sales_price")
        or price_mismatch_count
    )

    if price_issue_present:
        lines.append(f"{section_num}. Price")
        if missing_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat is missing for {missing_by_col['price_per_carat']} item(s).")
        if missing_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price is missing for {missing_by_col['total_sales_price']} item(s).")
        if invalid_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat has invalid values (zero, negative, or not in accepted format) for {invalid_by_col['price_per_carat']} item(s).")
        if invalid_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price has invalid values (zero, negative, or not in accepted format) for {invalid_by_col['total_sales_price']} item(s).")
        if price_mismatch_count:
            lines.append(f"- For {price_mismatch_count} item(s), Total Sales Price does not match (Carat x Price Per Carat).")
        lines.append("")
        section_num += 1

    if cut_missing_count:
        lines.append(f"{section_num}. Other Issues (Cut Grade)")
        lines.append(f"- Cut grade information is missing for {cut_missing_count} item(s).")
        lines.append("")

    lines.append("A spreadsheet outlining the above items has been attached for your reference. We would appreciate it if you could make the necessary corrections at your earliest convenience.")
    lines.append("")
    lines.append("If you have any questions or need further clarification, feel free to reach out. We'll be happy to assist.")
    lines.append("")
    lines.append("Best Regards,")
    lines.append("VDB Marketplace Support Team")

    return "\n".join(lines)

# ------------------------------------------------------------
# INITIALIZE SESSION STATE
# ------------------------------------------------------------

if 'validation_complete' not in st.session_state:
    st.session_state.validation_complete = False
if 'validation_results' not in st.session_state:
    st.session_state.validation_results = None

# ------------------------------------------------------------
# FILE UPLOAD UI
# ------------------------------------------------------------

col_upload, col_reset = st.columns([4, 1])

with col_upload:
    supplier_file = st.file_uploader("Upload Supplier Inventory (.csv or .xlsx)", type=["csv", "xlsx"])
    supplier_name = st.text_input("Supplier Name (for email)", value="Supplier")

with col_reset:
    st.write("")
    st.write("")
    if st.button("🔄 Reset", help="Clear all results and start fresh"):
        st.session_state.validation_complete = False
        st.session_state.validation_results = None
        st.session_state.last_file_name = None
        st.rerun()

# Reset validation when new file is uploaded
if supplier_file and st.session_state.get('last_file_name') != supplier_file.name:
    st.session_state.validation_complete = False
    st.session_state.validation_results = None
    st.session_state.last_file_name = supplier_file.name

start_btn = st.button("Run Validation")

# ------------------------------------------------------------
# MAIN FLOW
# ------------------------------------------------------------

if start_btn and supplier_file:

    if not supplier_file:
        st.error("⚠ Please upload the Supplier Inventory file.")
        st.stop()

    rules_path = "headers.xlsx"
    if not os.path.exists(rules_path):
        st.error(f"Configuration error: The rules file ({rules_path}) was not found.")
        st.stop()
        
    st.info("📘 Loading rules…")
    try:
        header_map, canonical_set = validator.load_header_rules(rules_path)
        value_rules = validator.load_value_rules(rules_path)
        st.success("Rules loaded successfully.")
    except Exception as e:
        st.error(f"Failed to load rules. Error: {e}")
        st.stop()
    
    st.info("📄 Loading supplier inventory…")
    supplier_bytes = supplier_file.read()
    ext = supplier_file.name.split(".")[-1].lower()

    if ext == "csv":
        df = pd.read_csv(io.BytesIO(supplier_bytes))
    else:
        df = pd.read_excel(io.BytesIO(supplier_bytes))

    st.success(f"Supplier file loaded: **{len(df)} rows**")

    progress = st.progress(0)
    status = st.empty()

    status.text("Normalizing headers…")
    df, unknown_headers = validator.normalize_headers(df, header_map)
    progress.progress(12)

    status.text("Checking mandatory fields…")
    missing_strings = validator.check_mandatory(df)
    mandatory_issues, missing_by_col = build_mandatory_issues(df)
    missing_stock_count = sum(1 for issue in mandatory_issues if issue["Column"] == "stock_num")
    progress.progress(25)

    status.text("Checking numeric ranges…")
    numeric_invalid_strings = validator.check_numeric_ranges(df)
    numeric_invalid_issues, numeric_invalid_by_col = parse_numeric_invalid_strings(numeric_invalid_strings, df)
    progress.progress(40)

    status.text("Validating values…")
    invalid_strings = validator.check_values(df, value_rules)
    invalid_issues, invalid_shape_values, invalid_color_values, invalid_by_col = parse_invalid_value_strings(invalid_strings, df)
    progress.progress(60)

    for col, count in numeric_invalid_by_col.items():
        invalid_by_col[col] += count

    status.text("Checking URLs… (fast mode)")
    url_strings = validator.check_all_urls(df)
    url_issues_struct, url_counts = parse_url_issue_strings(url_strings, df)
    progress.progress(75)

    status.text("Checking cut grade and price consistency…")
    cut_issues, cut_missing_count = find_missing_cut_grade(df)
    price_issues, price_mismatch_count = build_price_mismatch_issues(df)
    progress.progress(100)

    # --------------------------------------------------------
    # BUILD FINAL RESULTS
    # --------------------------------------------------------
    structured_issues = []
    structured_issues.extend(mandatory_issues)
    structured_issues.extend(numeric_invalid_issues)
    structured_issues.extend(invalid_issues)
    structured_issues.extend(url_issues_struct)
    structured_issues.extend(cut_issues)
    structured_issues.extend(price_issues)

    excel_buffer = build_excel_report(structured_issues)

    email_body = build_email_body(
        supplier_name=supplier_name,
        invalid_shape_values=invalid_shape_values,
        invalid_color_values=invalid_color_values,
        missing_by_col=missing_by_col,
        invalid_by_col=invalid_by_col,
        url_counts=url_counts,
        cut_missing_count=cut_missing_count,
        price_mismatch_count=price_mismatch_count,
        missing_stock_count=missing_stock_count,
    )

    st.success("✅ Validation completed!")

    # Store results in session state
    st.session_state.validation_complete = True
    st.session_state.validation_results = {
        'df': df,
        'unknown_headers': unknown_headers,
        'missing_strings': missing_strings,
        'numeric_invalid_strings': numeric_invalid_strings,
        'invalid_strings': invalid_strings,
        'url_strings': url_strings,
        'structured_issues': structured_issues,
        'email_body': email_body,
        'invalid_shape_values': invalid_shape_values,
        'invalid_color_values': invalid_color_values,
        'missing_by_col': missing_by_col,
        'invalid_by_col': invalid_by_col,
        'url_counts': url_counts,
        'cut_missing_count': cut_missing_count,
        'price_mismatch_count': price_mismatch_count,
        'missing_stock_count': missing_stock_count,
        'excel_buffer': excel_buffer,
    }

# ------------------------------------------------------------
# DISPLAY RESULTS (from session state if available)
# ------------------------------------------------------------

if st.session_state.validation_complete and st.session_state.validation_results:
    
    results = st.session_state.validation_results
    
    # --------------------------------------------------------
    # SHOW RAW RESULTS
    # --------------------------------------------------------
    st.subheader("📌 Raw Validation Output")

    if results['unknown_headers']:
        st.warning("⚠ Unknown Headers Found")
        st.write(results['unknown_headers'])

    if results['missing_strings']:
        st.error("❌ Missing Mandatory Fields")
        st.write(results['missing_strings'])

    if results['numeric_invalid_strings']:
        st.error("❌ Invalid Numeric Values (Zero/Negative)")
        st.write(results['numeric_invalid_strings'])

    if results['invalid_strings']:
        st.error("❌ Invalid Values Found")
        st.write(results['invalid_strings'])

    filtered_url_strings = [s for s in results['url_strings'] if 'cert_url_1' not in s or 'NOT PROVIDED' in s]
    
    if filtered_url_strings:
        st.error("❌ URL Issues (Image/Video/Cert)")
        st.write(filtered_url_strings)
    else:
        st.success("✅ All URLs are working or missing.")

    # --------------------------------------------------------
    # EXCEL REPORT DOWNLOAD
    # --------------------------------------------------------
    st.subheader("📊 Download Detailed Spreadsheet")
    st.download_button(
        label="📥 Download validation_report.xlsx",
        data=results['excel_buffer'],
        file_name="validation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel"
    )

    # --------------------------------------------------------
    # EMAIL SUMMARY WITH COPY
    # --------------------------------------------------------
    st.subheader("📧 Email Summary")

    st.text_area("Email to Supplier", value=results['email_body'], height=400, key="email_text")
    
    # Copy button that shows copyable text
    if st.button("📋 Copy Email to Clipboard", use_container_width=False, help="Click to display copyable email"):
        st.info("👇 Select all the text below and copy it (Ctrl+A, Ctrl+C or Cmd+A, Cmd+C)")
        st.code(results['email_body'], language=None)
