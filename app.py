"""
Invoice Data Extractor - Professional Business UI
Brillspark v3.0
"""

import streamlit as st
import pandas as pd
from PIL import Image
import pytesseract
import re
from datetime import datetime
import io

# ---------------- PAGE CONFIG ---------------- #
st.set_page_config(
    page_title="Brillspark Invoice Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------- CUSTOM CSS ---------------- #
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

* {
    font-family: 'Inter', sans-serif;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display: none;}

:root {
    --primary-blue: #1E4FA3;
    --dark-blue: #163D7A;
    --accent-green: #5FAF2D;
    --soft-bg: #F8FAFC;
    --border-color: #E5E7EB;
    --text-dark: #111827;
    --text-light: #6B7280;
}

.stApp {
    background: var(--soft-bg);
}

/* HERO SECTION */
.hero-card {
    background: white;
    border-radius: 24px;
    padding: 3.5rem 2rem;
    text-align: center;
    margin-bottom: 1.5rem;
    box-shadow: 0 10px 30px rgba(0,0,0,0.06);
}

.hero-title {
    font-size: 2.8rem;
    font-weight: 700;
    color: var(--text-dark);
    margin-bottom: 0.75rem;
}

.hero-subtitle {
    font-size: 1.1rem;
    color: var(--primary-blue);
    font-weight: 600;
    margin-bottom: 0.75rem;
}

.hero-description {
    font-size: 1rem;
    color: var(--text-light);
    max-width: 700px;
    margin: auto;
    line-height: 1.6;
}

/* FEATURE TAGS */
.feature-tags {
    display:flex;
    gap:14px;
    justify-content:center;
    margin-top:20px;
    flex-wrap:wrap;
}

.feature-pill {
    background:white;
    padding:10px 18px;
    border-radius:999px;
    border:1px solid var(--border-color);
    font-size:14px;
    color:#374151;
    font-weight:500;
}

/* CARDS */
.clean-card {
    background: white;
    border-radius: 18px;
    padding: 1.6rem;
    box-shadow: 0 4px 14px rgba(0,0,0,0.05);
    margin-bottom: 1.5rem;
}

.card-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: var(--text-dark);
    margin-bottom: 1rem;
}

/* BUTTONS */
.stButton > button {
    background: white;
    color: var(--primary-blue);
    border: 2px solid var(--primary-blue);
    border-radius: 10px;
    padding: 0.7rem 1.2rem;
    font-weight: 600;
    font-size: 0.95rem;
    transition: 0.2s ease;
}

.stButton > button:hover {
    background: var(--primary-blue);
    color: white;
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(30,79,163,0.25);
}

.stButton > button[kind="primary"] {
    background: var(--primary-blue);
    color: white;
    border: none;
}

.stButton > button[kind="primary"]:hover {
    background: var(--dark-blue);
}

/* FILE UPLOADER */
.stFileUploader {
    background: white;
    border: 2px dashed #CBD5E1;
    border-radius: 16px;
    padding: 2.5rem;
    transition: all 0.2s ease;
}

.stFileUploader:hover {
    border-color: var(--primary-blue);
    background: #F8FBFF;
}

/* CHECKBOXES */
.stCheckbox > label {
    font-size: 0.92rem;
    font-weight: 500;
}

/* TABS */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background: #F3F4F6;
    padding: 0.5rem;
    border-radius: 10px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 0.55rem 1rem;
    font-weight: 500;
}

.stTabs [aria-selected="true"] {
    background: white;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}

/* TEXT INPUT */
.stTextInput > div > div > input {
    border-radius: 10px;
    border: 2px solid #E5E7EB;
    padding: 0.75rem;
}

/* ACTIVITY LOG */
.activity-log {
    background: #F9FAFB;
    border-radius: 10px;
    padding: 1rem;
    font-family: monospace;
    font-size: 0.85rem;
    line-height: 1.6;
    color: #4B5563;
    height: 420px;
    overflow-y: auto;
}

/* METRICS */
[data-testid="metric-container"] {
    background: white;
    border: 1px solid #E5E7EB;
    padding: 1rem;
    border-radius: 14px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}

/* DATAFRAME */
.stDataFrame {
    border-radius: 12px;
    overflow: hidden;
}

h1, h2, h3 {
    color: var(--text-dark);
}
</style>
""", unsafe_allow_html=True)

# ---------------- SESSION STATE ---------------- #
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = []

if 'activity_log' not in st.session_state:
    st.session_state.activity_log = []

if 'excel_path' not in st.session_state:
    st.session_state.excel_path = "extracted_invoices.xlsx"

# ---------------- FUNCTIONS ---------------- #
def log_activity(message):
    timestamp = datetime.now().strftime('[%H:%M:%S]')
    st.session_state.activity_log.append(f"{timestamp} {message}")

def extract_field(text, field_name):
    text_lower = text.lower()
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    if field_name == 'Invoice Number':
        patterns = [
            r'Invoice\s+No\.?\s*:?\s*([A-Z\d#\-/]+)',
            r'Quotation\s+No\.?\s*:?\s*([A-Z\d#\-/]+)',
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)

    elif field_name == 'Date':
        patterns = [
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
            r'(\d{1,2}\s+[A-Za-z]+,?\s+\d{4})'
        ]

        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return match.group(1)

    elif field_name == 'Company Name':
        if lines:
            return lines[0]

    elif field_name == 'Customer Name':
        for i, line in enumerate(lines):
            if 'bill to' in line.lower() or 'invoice to' in line.lower():
                if i + 1 < len(lines):
                    return lines[i + 1]

    elif field_name == 'Total Amount':
        amounts = re.findall(r'[\d,]+\.?\d*', text)

        cleaned = []

        for amt in amounts:
            try:
                value = float(amt.replace(',', ''))
                if value > 500:
                    cleaned.append(value)
            except:
                continue

        if cleaned:
            return f"₹ {max(cleaned):,.0f}"

    return "N/A"

def process_invoice(image_file, fields_to_extract):
    try:
        image = Image.open(image_file)
        image = image.convert('L')

        log_activity(f"Processing: {image_file.name}")

        text = pytesseract.image_to_string(image, lang='eng')

        extracted = {
            'Filename': image_file.name,
            'Processed Time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        for field in fields_to_extract:
            extracted[field] = extract_field(text, field)

        return extracted

    except Exception as e:
        log_activity(f"Error: {str(e)}")
        return None

def convert_df_to_excel_bytes(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoice Data')

    output.seek(0)
    return output

# ---------------- HERO SECTION ---------------- #
st.markdown("""
<div class="hero-card">

    <h1 class="hero-title">
        Invoice Data Extractor
    </h1>

    <p class="hero-subtitle">
        Fast. Accurate. Professional.
    </p>

    <p class="hero-description">
        Extract structured invoice data quickly and export it
        in Excel or CSV format.
    </p>

    <div class="feature-tags">

        <div class="feature-pill">
            📄 Invoice Processing
        </div>

        <div class="feature-pill">
            📊 Excel Export
        </div>

        <div class="feature-pill">
            ⚡ Fast Extraction
        </div>

        <div class="feature-pill">
            🔒 Secure Processing
        </div>

    </div>

</div>
""", unsafe_allow_html=True)

# ---------------- MAIN LAYOUT ---------------- #
left_col, right_col = st.columns([2, 1], gap="large")

with left_col:

    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📤 Upload & Process</p>', unsafe_allow_html=True)

    st.markdown("### Select Fields")

    c1, c2, c3 = st.columns(3)

    with c1:
        inv_num = st.checkbox('📋 Invoice Number', value=True)
        date = st.checkbox('📅 Date', value=True)

    with c2:
        customer = st.checkbox('👤 Customer Name', value=True)
        total = st.checkbox('💰 Total Amount', value=True)

    with c3:
        company = st.checkbox('🏢 Company Name', value=True)

    field_options = {
        'Invoice Number': inv_num,
        'Date': date,
        'Company Name': company,
        'Customer Name': customer,
        'Total Amount': total
    }

    selected_fields = [
        field for field, selected in field_options.items() if selected
    ]

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["📁 Multiple Invoices", "📄 Single Invoice"])

    # -------- MULTIPLE FILES -------- #
    with tab1:

        folder_files = st.file_uploader(
            "Upload Invoice Files",
            type=['jpg', 'jpeg', 'png'],
            accept_multiple_files=True
        )

        if folder_files:

            st.success(f"{len(folder_files)} files ready for processing")

            if st.button(
                "🚀 Process Invoices",
                type="primary",
                use_container_width=True
            ):

                results = []

                progress = st.progress(0)

                for i, file in enumerate(folder_files):

                    progress.progress((i + 1) / len(folder_files))

                    data = process_invoice(file, selected_fields)

                    if data:
                        results.append(data)

                progress.empty()

                if results:
                    st.session_state.extracted_data = results
                    st.toast("Invoices processed successfully")
                    st.success(f"{len(results)} invoices processed successfully")

    # -------- SINGLE FILE -------- #
    with tab2:

        single_file = st.file_uploader(
            "Upload Invoice",
            type=['jpg', 'jpeg', 'png']
        )

        if single_file:

            image = Image.open(single_file)

            st.image(
                image,
                caption=single_file.name,
                use_container_width=True
            )

            if st.button(
                "🚀 Extract Data",
                type="primary",
                use_container_width=True
            ):

                data = process_invoice(single_file, selected_fields)

                if data:
                    st.session_state.extracted_data = [data]
                    st.toast("Data extracted successfully")
                    st.success("Invoice processed successfully")

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- RIGHT COLUMN ---------------- #
with right_col:

    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📋 Activity Log</p>', unsafe_allow_html=True)

    if st.session_state.activity_log:

        log_text = "\n".join(
            st.session_state.activity_log[-15:]
        )

        st.markdown(
            f'<div class="activity-log">{log_text}</div>',
            unsafe_allow_html=True
        )

        if st.button(
            "🗑️ Clear Log",
            use_container_width=True
        ):
            st.session_state.activity_log = []
            st.rerun()

    else:
        st.info("Activity updates will appear here.")

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- RESULTS ---------------- #
if st.session_state.extracted_data:

    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📊 Extracted Data</p>', unsafe_allow_html=True)

    df = pd.DataFrame(st.session_state.extracted_data)

    m1, m2, m3 = st.columns(3)

    with m1:
        st.metric("Invoices Processed", len(df))

    with m2:
        st.metric("Fields Extracted", len(df.columns))

    with m3:
        st.metric("Export Status", "Ready")

    st.markdown("<br>", unsafe_allow_html=True)

    st.dataframe(
        df,
        use_container_width=True,
        height=420,
        hide_index=True
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        excel_data = convert_df_to_excel_bytes(df)

        st.download_button(
            "📥 Download Excel",
            data=excel_data,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c2:

        csv = df.to_csv(index=False)

        st.download_button(
            "📥 Download CSV",
            data=csv,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    with c3:

        if st.button(
            "🗑️ Clear Data",
            use_container_width=True
        ):
            st.session_state.extracted_data = []
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- FOOTER ---------------- #
st.markdown("""
<div style="
    text-align:center;
    padding:2rem 0;
    color:#6B7280;
    font-size:0.9rem;
    margin-top:2rem;
">

    <p style="margin:0;font-weight:600;">
        Brillspark Invoice Extractor
    </p>

    <p style="margin-top:6px;">
        Secure invoice processing • Optimized for business documents
    </p>

</div>
""", unsafe_allow_html=True)
