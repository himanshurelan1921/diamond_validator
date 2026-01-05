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
Upload your **supplier inventory file** to validate against the internal rule set.
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


def parse_url_issue_strings(url_list, df):
    """
    Parse strings like:
      "Row 2: image_url_1 → NOT WORKING"
      "Row 3: cert_url_1 → NOT PROVIDED"
    into structured records and counts.
    NOTE: URL checks for cert_url_1 are ignored in validator.py now.
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

        # Skip cert_url_1 entries from the URL checker string list,
        # as we rely on the PDF format check now.
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

        # counts for email
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
            # Only count missing certs here; bad certs (non-PDF) counted in Step 5
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
    """
    Look for missing cut grade values if such a column exists.
    We'll look for 'cut' or 'cut_grade' canonical names.
    """
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
    """
    Detect certificate URLs that are not direct PDF links.
    """
    issues = []
    count = 0

    if "cert_url_1" not in df.columns:
        return issues, 0

    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        url = row.get("cert_url_1", None)
        v_norm = validator.normalize_value_str(url)
        if v_norm is None:
            continue

        url_lower = str(url).lower()
        if ".pdf" not in url_lower:
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
    """
    Check if (carat * price_per_carat) == total_sales_price
    using canonical columns if present.
    Only returns mismatches.
    """
    issues = []
    count = 0

    weight_col = find_canonical_col(df, ["carat", "weight", "carat_weight"])
    ppc_col = find_canonical_col(df, ["price_per_carat"])
    tsp_col = find_canonical_col(df, ["total_sales_price"])

    if not (weight_col and ppc_col and tsp_col):
        return issues, 0

    def to_float(x):
        if x is None:
            return None
        if isinstance(x, float) and math.isnan(x):
            return None
        s = str(x).strip()
        if not s:
            return None
        s = s.replace(",", "")
        s = re.sub(r"[^\d.\-]", "", s)
        try:
            return float(s)
        except Exception:
            return None

    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        w = to_float(row.get(weight_col, None))
        ppc = to_float(row.get(ppc_col, None))
        tsp = to_float(row.get(tsp_col, None))

        if w is None or ppc is None or tsp is None:
            continue

        expected = round(w * ppc, 2)
        # strict check; small tolerance for float rounding
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
    """
    Build an Excel file in memory, writing issues to separate sheets
    based on the affected canonical column/category.
    """
    # Define sections and their sheets based on canonical column or issue type
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
    # Catch-all sheet for other issues (e.g., Missing Stock_num, Unknown Header, general errors)
    other_sheet = "9. Other Issues / Cut Grade"
    
    # Organize issues by sheet name
    issues_by_sheet = {}
    
    # Initialize all possible sheets to ensure they are sorted correctly, even if empty
    all_sheet_names = set(section_map.values())
    all_sheet_names.add(other_sheet)
    for name in all_sheet_names:
        issues_by_sheet[name] = []
        
    for issue in structured_issues:
        column = issue.get("Column", "").lower()
        sheet_name = section_map.get(column, other_sheet)
        
        # Special handling for Price Mismatch and Cert URL Format
        if issue["Issue Type"] == "Price Mismatch":
             sheet_name = "8. Price"
        elif issue["Issue Type"] == "Cert URL Format":
             sheet_name = "7. Certificate URL"
        elif issue["Column"] in ["cut", "cut_grade"]:
            sheet_name = other_sheet

        issues_by_sheet[sheet_name].append(issue)

    # Write to Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Sort sheets by name (e.g., 1., 2., 3.)
        for sheet_name in sorted(issues_by_sheet.keys()):
            
            # --- APPLY SANITIZATION HERE ---
            safe_sheet_name = sanitize_sheet_name(sheet_name)
            # --- END SANITIZATION ---
            
            # Only write sheets that have data
            if issues_by_sheet[sheet_name]:
                # Remove the temporary 'Category' column before writing
                df_report = pd.DataFrame(issues_by_sheet[sheet_name]).drop(columns=["Category"], errors='ignore')
                df_report.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            
        # If no issues, create a single "No Issues" sheet
        if not structured_issues:
            df_empty = pd.DataFrame(columns=["Stock No.", "Issue Type", "Column", "Value", "Details", "Row"])
            df_empty.to_excel(writer, sheet_name="No Issues Found", index=False)
                
    
    buffer.seek(0)
    return buffer


def build_email_body(
    supplier_name,
    invalid_shape_values,
    missing_by_col,
    invalid_by_col,
    url_counts,
    cut_missing_count,
    non_pdf_cert_count,
    price_mismatch_count,
):
    """
    Build the final email text based on actual issues.
    (2nd Change: No emojis/icons, No horizontal rule)
    (3rd Change: Removed redundant 'refer to sheet' text)
    """
    lines = []

    lines.append(f"Hi {supplier_name},")
    lines.append("")
    lines.append("Hope you're doing well.")
    lines.append("")
    lines.append("During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:")
    lines.append("")


    # --- Shape (1. Shape) ---
    if invalid_shape_values or missing_by_col.get("shape") or invalid_by_col.get("shape"):
        lines.append("1. Shape")
        if missing_by_col.get("shape"):
            lines.append(f"- Shape is missing for {missing_by_col['shape']} item(s).")
        if invalid_shape_values:
            lines.append("- We found invalid shape values that do not match VDB’s standardised shape list, for example:")
            for sh in invalid_shape_values:
                lines.append(f"  • {sh}")
        lines.append("")


    # --- Weight (2. Weight) ---
    # Determine weight column label for email (if any)
    weight_col = None
    for cand in ["carat", "weight", "carat_weight"]:
        if cand in missing_by_col or cand in invalid_by_col:
            weight_col = cand
            break

    if weight_col and (missing_by_col.get(weight_col) or invalid_by_col.get(weight_col)):
        lines.append("2. Weight")
        if missing_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) is missing for {missing_by_col[weight_col]} item(s).")
        if invalid_by_col.get(weight_col):
            lines.append(f"- Weight ({weight_col}) has invalid values for {invalid_by_col[weight_col]} item(s).")
        lines.append("")


    # --- Color (3. Color) ---
    if missing_by_col.get("color") or invalid_by_col.get("color"):
        lines.append("3. Color")
        if missing_by_col.get("color"):
            lines.append(f"- Color is missing for {missing_by_col['color']} item(s).")
        if invalid_by_col.get("color"):
            lines.append(f"- Color has invalid values for {invalid_by_col['color']} item(s).")
        lines.append("")


    # --- Clarity (4. Clarity) ---
    if missing_by_col.get("clarity") or invalid_by_col.get("clarity"):
        lines.append("4. Clarity")
        if missing_by_col.get("clarity"):
            lines.append(f"- Clarity is missing for {missing_by_col['clarity']} item(s).")
        if invalid_by_col.get("clarity"):
            lines.append(f"- Clarity has invalid values for {invalid_by_col['clarity']} item(s).")
        lines.append("")


    # --- Image (5. Image URL) ---
    missing_image = missing_by_col.get("image_url_1", 0) + url_counts.get("missing_image", 0)
    if missing_image or url_counts.get("bad_image", 0):
        lines.append("5. Image URLs")
        if missing_image:
            lines.append(f"- Image URLs are missing for {missing_image} item(s).")
        if url_counts.get("bad_image", 0):
            lines.append(f"- {url_counts['bad_image']} image URL(s) are not working (HTTP errors).")
        lines.append("")

    # --- Video (6. Video URL) ---
    missing_video = missing_by_col.get("video_url_1", 0) + url_counts.get("missing_video", 0)
    if missing_video or url_counts.get("bad_video", 0):
        lines.append("6. Video URLs")
        if missing_video:
            lines.append(f"- Video URLs are missing for {missing_video} item(s).")
        if url_counts.get("bad_video", 0):
            lines.append(f"- {url_counts['bad_video']} video URL(s) are not working (HTTP errors).")
        lines.append("")


    # --- Certificate (7. Certificate URL) ---
    cert_issue_present = (
        missing_by_col.get("cert_url_1", 0)
        or non_pdf_cert_count
        or url_counts.get("bad_cert", 0)
    )
    if cert_issue_present:
        lines.append("7. Certificate URLs")
        if missing_by_col.get("cert_url_1", 0) + url_counts.get("bad_cert", 0) > 0:
             lines.append(f"- Certificate URLs are missing for {missing_by_col['cert_url_1'] + url_counts['bad_cert']} item(s).")
        if non_pdf_cert_count:
            lines.append(f"- {non_pdf_cert_count} certificate URL(s) do not appear to be direct PDF links. Please use direct PDF links (e.g., `https://.../LG700531186.pdf`).")
        lines.append("")

    # --- Price (8. Price) ---
    price_issue_present = (
        missing_by_col.get("price_per_carat")
        or missing_by_col.get("total_sales_price")
        or invalid_by_col.get("price_per_carat")
        or invalid_by_col.get("total_sales_price")
        or price_mismatch_count
    )

    if price_issue_present:
        lines.append("8. Price")
        if missing_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat is missing for {missing_by_col['price_per_carat']} item(s).")
        if missing_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price is missing for {missing_by_col['total_sales_price']} item(s).")
        if invalid_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat has invalid values for {invalid_by_col['price_per_carat']} item(s).")
        if invalid_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price has invalid values for {invalid_by_col['total_sales_price']} item(s).")
        if price_mismatch_count:
            lines.append(f"- For {price_mismatch_count} item(s), Total Sales Price does not match (Carat x Price Per Carat).")
        lines.append("")


    # --- Cut Grade (9. Other Issues) ---
    if cut_missing_count:
        lines.append("9. Other Issues (Cut Grade)")
        lines.append(f"- Cut grade information is missing for {cut_missing_count} item(s).")
        lines.append("")

    # Closing
    lines.append("A spreadsheet outlining the above items has been attached for your reference. We would appreciate it if you could make the necessary corrections at your earliest convenience.")
    lines.append("")
    lines.append("If you have any questions or need further clarification, feel free to reach out. We’ll be happy to assist.")
    lines.append("")
    lines.append("Best Regards,")
    lines.append("Himanshu")
    lines.append("VDB Marketplace Support")

    return "\n".join(lines)


# ------------------------------------------------------------
# FILE UPLOAD UI
# ------------------------------------------------------------

# (4th Change: Removed rules_file upload)
supplier_file = st.file_uploader("Upload Supplier Inventory (.csv or .xlsx)", type=["csv", "xlsx"])
supplier_name = st.text_input("Supplier Name (for email)", value="Supplier")

start_btn = st.button("Run Validation")


# ------------------------------------------------------------
# MAIN FLOW
# ------------------------------------------------------------

if start_btn:

    if not supplier_file:
        st.error("⚠ Please upload the Supplier Inventory file.")
        st.stop()

    # (4th Change: Loading rules locally)
    rules_path = "headers.xlsx"
    if not os.path.exists(rules_path):
        st.error(f"Configuration error: The rules file ({rules_path}) was not found in the application directory.")
        st.stop()
        
    # Load rules
    st.info("📘 Loading rules…")
    try:
        header_map, canonical_set = validator.load_header_rules(rules_path)
        value_rules = validator.load_value_rules(rules_path)
        st.success("Rules loaded successfully.")
    except Exception as e:
        st.error(f"Failed to load rules from {rules_path}. Ensure it is a valid Excel file with 'Columns' and 'Values' sheets. Error: {e}")
        st.stop()
    
    # Load supplier file
    st.info("📄 Loading supplier inventory…")
    # ... (rest of file loading remains the same)
    supplier_bytes = supplier_file.read()
    ext = supplier_file.name.split(".")[-1].lower()

    if ext == "csv":
        df = pd.read_csv(io.BytesIO(supplier_bytes))
    else:
        df = pd.read_excel(io.BytesIO(supplier_bytes))

    st.success(f"Supplier file loaded: **{len(df)} rows**")

    progress = st.progress(0)
    status = st.empty()

    # STEP 1 — Normalize headers
    status.text("Normalizing headers…")
    df, unknown_headers = validator.normalize_headers(df, header_map)
    progress.progress(15)

    # STEP 2 — Mandatory fields
    status.text("Checking mandatory fields…")
    missing_strings = validator.check_mandatory(df) # Raw string list for CLI output
    mandatory_issues, missing_by_col = build_mandatory_issues(df) # Structured issues & counts
    progress.progress(35)

    # STEP 3 — Value checks
    status.text("Validating values…")
    invalid_strings = validator.check_values(df, value_rules)
    invalid_issues, invalid_shape_values = parse_invalid_value_strings(invalid_strings, df)
    progress.progress(60)

    # STEP 4 — URL checks
    status.text("Checking URLs… (fast mode)")
    url_strings = validator.check_all_urls(df)
    url_issues_struct, url_counts = parse_url_issue_strings(url_strings, df)
    progress.progress(80)

    # STEP 5 — Special: cut grade, cert format, price mismatches
    status.text("Checking cut grade, certificate URL format and price consistency…")
    cut_issues, cut_missing_count = find_missing_cut_grade(df)
    # The URL status check for certs is implicitly handled by the change in validator.py
    cert_format_issues, non_pdf_cert_count = find_non_pdf_cert_urls(df) 
    price_issues, price_mismatch_count = build_price_mismatch_issues(df)
    progress.progress(100)

    st.success("Validation completed!")

    # --------------------------------------------------------
    # SHOW RAW RESULTS (like CLI)
    # --------------------------------------------------------
    st.subheader("📌 Raw Validation Output")

    if unknown_headers:
        st.warning("⚠ Unknown Headers Found")
        st.write(unknown_headers)

    if missing_strings:
        st.error("❌ Missing Mandatory Fields")
        st.write(missing_strings)

    if invalid_strings:
        st.error("❌ Invalid Values Found")
        st.write(invalid_strings)

    # Filter out cert_url_1 URL errors if they are not just 'NOT PROVIDED'
    filtered_url_strings = [s for s in url_strings if 'cert_url_1' not in s or 'NOT PROVIDED' in s]
    
    if filtered_url_strings:
        st.error("❌ URL Issues (Image/Video)")
        st.write(filtered_url_strings)
    else:
        st.success("All Image and Video URLs are working or missing.")

    # --------------------------------------------------------
    # STRUCTURED ISSUES & EXCEL REPORT
    # --------------------------------------------------------
    structured_issues = []
    structured_issues.extend(mandatory_issues)
    structured_issues.extend(invalid_issues)
    structured_issues.extend(url_issues_struct)
    structured_issues.extend(cut_issues)
    structured_issues.extend(cert_format_issues)
    structured_issues.extend(price_issues)

    excel_buffer = build_excel_report(structured_issues)

    st.subheader("📊 Download Detailed Spreadsheet")
    st.download_button(
        label="📥 Download validation_report.xlsx",
        data=excel_buffer,
        file_name="validation_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # --------------------------------------------------------
    # EMAIL SUMMARY
    # --------------------------------------------------------
    st.subheader("📧 Email Summary (copy & paste)")
    
    invalid_by_col = Counter()
    for i in invalid_issues:
        invalid_by_col[i["Column"]] += 1

    email_body = build_email_body(
        supplier_name=supplier_name,
        invalid_shape_values=invalid_shape_values,
        missing_by_col=missing_by_col,
        invalid_by_col=invalid_by_col,
        url_counts=url_counts,
        cut_missing_count=cut_missing_count,
        non_pdf_cert_count=non_pdf_cert_count,
        price_mismatch_count=price_mismatch_count,
    )

    st.text_area("Email to Supplier", value=email_body, height=400)
