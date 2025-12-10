import streamlit as st
import pandas as pd
import io
import re
import math
from collections import Counter
import validator  # our backend module


# ------------------------------------------------------------
# STREAMLIT SETUP
# ------------------------------------------------------------

st.set_page_config(page_title="Diamond Inventory Validator", layout="wide")
st.title("💎 Diamond & Lab-Grown Inventory Validator")
st.markdown("""
Upload your **headers.xlsx** (rules file)  
and your **supplier inventory file** to validate and generate a client-ready report.
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


def build_mandatory_issues(df):
    """Return structured list of missing mandatory field issues."""
    issues = []
    mandatory_cols = getattr(validator, "MANDATORY_COLS", [])
    for idx, row in df.iterrows():
        stock = row.get("stock_num", None)
        for col in mandatory_cols:
            if col in df.columns:
                v_norm = validator.normalize_value_str(row[col])
                if v_norm is None:
                    issues.append({
                        "Stock No.": stock,
                        "Issue Type": "Missing Value",
                        "Column": col,
                        "Value": row[col],
                        "Details": "Missing mandatory field",
                        "Row": idx + 2,
                    })
    return issues


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
            if "NOT WORKING" in status:
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
    Build an Excel file in memory from structured issues.
    """
    if not structured_issues:
        df_report = pd.DataFrame(columns=["Stock No.", "Issue Type", "Column", "Value", "Details", "Row"])
    else:
        df_report = pd.DataFrame(structured_issues)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_report.to_excel(writer, sheet_name="Validation Report", index=False)
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
    Only mention sections that have problems.
    Focus on:
      Shape, Weight, Color, Clarity, Price, Image, Video, Certificate.
    """
    lines = []

    lines.append(f"Hi {supplier_name},")
    lines.append("")
    lines.append("Hope you're doing well.")
    lines.append("")
    lines.append("During a routine validation of your inventory on the VDB Marketplace, we identified a few issues that need your attention. Please find the details below:")
    lines.append("")

    # ----- Shape -----
    if invalid_shape_values or missing_by_col.get("shape"):
        lines.append("Shape")
        if missing_by_col.get("shape"):
            lines.append(f"- Shape is missing for {missing_by_col['shape']} item(s).")
        if invalid_shape_values:
            lines.append("- We found shape values that do not match VDB’s standardised shape list, for example:")
            for sh in invalid_shape_values:
                lines.append(f"  • {sh}")
        lines.append("Please refer to the attached sheet for affected stock numbers.")
        lines.append("")

    # Determine weight column label for email (if any)
    weight_col = None
    for cand in ["carat", "weight", "carat_weight"]:
        if cand in missing_by_col or cand in invalid_by_col:
            weight_col = cand
            break

    # ----- Weight -----
    if weight_col and (missing_by_col.get(weight_col) or invalid_by_col.get(weight_col)):
        lines.append("Weight")
        if missing_by_col.get(weight_col):
            lines.append(f"- Weight is missing for {missing_by_col[weight_col]} item(s).")
        if invalid_by_col.get(weight_col):
            lines.append(f"- Weight has invalid values for {invalid_by_col[weight_col]} item(s).")
        lines.append("Affected stock numbers are listed in the attached sheet.")
        lines.append("")

    # ----- Color -----
    if missing_by_col.get("color") or invalid_by_col.get("color"):
        lines.append("Color")
        if missing_by_col.get("color"):
            lines.append(f"- Color is missing for {missing_by_col['color']} item(s).")
        if invalid_by_col.get("color"):
            lines.append(f"- Color has invalid values for {invalid_by_col['color']} item(s).")
        lines.append("")

    # ----- Clarity -----
    if missing_by_col.get("clarity") or invalid_by_col.get("clarity"):
        lines.append("Clarity")
        if missing_by_col.get("clarity"):
            lines.append(f"- Clarity is missing for {missing_by_col['clarity']} item(s).")
        if invalid_by_col.get("clarity"):
            lines.append(f"- Clarity has invalid values for {invalid_by_col['clarity']} item(s).")
        lines.append("")

    # ----- Price (PPC / TSP) -----
    price_issue_present = (
        missing_by_col.get("price_per_carat")
        or missing_by_col.get("total_sales_price")
        or invalid_by_col.get("price_per_carat")
        or invalid_by_col.get("total_sales_price")
        or price_mismatch_count
    )

    if price_issue_present:
        lines.append("Price")
        if missing_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat is missing for {missing_by_col['price_per_carat']} item(s).")
        if missing_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price is missing for {missing_by_col['total_sales_price']} item(s).")
        if invalid_by_col.get("price_per_carat"):
            lines.append(f"- Price per carat has invalid values for {invalid_by_col['price_per_carat']} item(s).")
        if invalid_by_col.get("total_sales_price"):
            lines.append(f"- Total sales price has invalid values for {invalid_by_col['total_sales_price']} item(s).")
        if price_mismatch_count:
            lines.append(f"- For {price_mismatch_count} item(s), Total Sales Price does not match (Carat × Price Per Carat).")
        lines.append("Please review and correct the pricing for the affected stock numbers listed in the attached sheet.")
        lines.append("")

    # ----- Image / Video / Certificate URLs -----
    missing_image = missing_by_col.get("image_url_1", 0) + url_counts.get("missing_image", 0)
    missing_video = missing_by_col.get("video_url_1", 0) + url_counts.get("missing_video", 0)

    # Image
    if missing_image or url_counts.get("bad_image", 0):
        lines.append("Image URLs")
        if missing_image:
            lines.append(f"- Image URLs are missing for {missing_image} item(s).")
        if url_counts.get("bad_image", 0):
            lines.append(f"- {url_counts['bad_image']} image URL(s) are not working.")
        lines.append("")

    # Video
    if missing_video or url_counts.get("bad_video", 0):
        lines.append("Video URLs")
        if missing_video:
            lines.append(f"- Video URLs are missing for {missing_video} item(s).")
        if url_counts.get("bad_video", 0):
            lines.append(f"- {url_counts['bad_video']} video URL(s) are not working.")
        lines.append("")

    # Certificate
    cert_issue_present = (
        missing_by_col.get("cert_url_1", 0)
        or url_counts.get("bad_cert", 0)
        or non_pdf_cert_count
    )
    if cert_issue_present:
        lines.append("Certificate URLs")
        if missing_by_col.get("cert_url_1", 0):
            lines.append(f"- Certificate URLs are missing for {missing_by_col['cert_url_1']} item(s).")
        if url_counts.get("bad_cert", 0):
            lines.append(f"- {url_counts['bad_cert']} certificate URL(s) are not currently working (HTTP errors).")
        if non_pdf_cert_count:
            lines.append(f"- {non_pdf_cert_count} certificate URL(s) do not appear to be direct PDF links.")
        lines.append("Kindly update them with secure, direct PDF links where possible. For example: https://api.igi.org/viewpdf.php?r=LG700531186")
        lines.append("")

    # Cut grade (optional, not in main 8 but still useful)
    if cut_missing_count:
        lines.append("Cut Grade")
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

rules_file = st.file_uploader("Upload headers.xlsx (rules file)", type=["xlsx"])
supplier_file = st.file_uploader("Upload Supplier Inventory (.csv or .xlsx)", type=["csv", "xlsx"])
supplier_name = st.text_input("Supplier Name (for email)", value="Supplier")

start_btn = st.button("Run Validation")


# ------------------------------------------------------------
# MAIN FLOW
# ------------------------------------------------------------

if start_btn:

    if not rules_file or not supplier_file:
        st.error("⚠ Please upload BOTH files.")
        st.stop()

    # Load rules
    st.info("📘 Loading rules…")
    rules_bytes = rules_file.read()
    header_map, canonical_set = validator.load_header_rules(io.BytesIO(rules_bytes))
    value_rules = validator.load_value_rules(io.BytesIO(rules_bytes))
    st.success("Rules loaded successfully.")

    # Load supplier file
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

    # STEP 1 — Normalize headers
    status.text("Normalizing headers…")
    df, unknown_headers = validator.normalize_headers(df, header_map)
    progress.progress(15)

    # STEP 2 — Mandatory fields (string list for display)
    status.text("Checking mandatory fields…")
    missing_strings = validator.check_mandatory(df)
    # structured mandatory issues for report & email
    mandatory_issues = build_mandatory_issues(df)
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

    if url_strings:
        st.error("❌ URL Issues")
        st.write(url_strings)
    else:
        st.success("All URLs are working!")

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

    # aggregate counts per column
    missing_by_col = Counter()
    for i in mandatory_issues:
        missing_by_col[i["Column"]] += 1

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
