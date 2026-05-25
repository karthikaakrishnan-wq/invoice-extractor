"""
Invoice Data Extractor - Clean Professional Design
Brillspark - v2.0
"""

import streamlit as st
import pandas as pd
from PIL import Image
import pytesseract
import re
from datetime import datetime
import io

# Page configuration
st.set_page_config(
    page_title="Invoice Extractor - Brillspark",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Clean Modern CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .stApp {
        background: #f8f9fa;
    }
    
    /* Hero Section - Clean White Card */
    .hero-card {
        background: white;
        border-radius: 16px;
        padding: 3rem 2rem;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    
    .hero-title {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1a1a1a;
        margin-bottom: 0.5rem;
    }
    
    .hero-subtitle {
        font-size: 1.1rem;
        color: #667eea;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .hero-description {
        font-size: 0.95rem;
        color: #6b7280;
        margin-bottom: 1.5rem;
    }
    
    /* Clean Card Style */
    .clean-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
    }
    
    .card-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1a1a1a;
        margin-bottom: 1rem;
    }
    
    /* Buttons - Clean Style */
    .stButton > button {
        background: white;
        color: #667eea;
        border: 2px solid #667eea;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.2s;
    }
    
    .stButton > button:hover {
        background: #667eea;
        color: white;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
    }
    
    .stButton > button[kind="primary"] {
        background: #667eea;
        color: white;
        border: none;
    }
    
    .stButton > button[kind="primary"]:hover {
        background: #5568d3;
        box-shadow: 0 4px 12px rgba(102, 126, 234, 0.4);
    }
    
    /* Field Selection - Modern Pills */
    .stCheckbox > label {
        background: #f3f4f6;
        border: 2px solid transparent;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        cursor: pointer;
        transition: all 0.2s;
        font-size: 0.9rem;
        font-weight: 500;
    }
    
    .stCheckbox > label:hover {
        background: #e5e7eb;
        border-color: #d1d5db;
    }
    
    .stCheckbox input[type="checkbox"]:checked + div {
        background: #ede9fe;
        border-color: #667eea;
        color: #667eea;
    }
    
    /* Text Input */
    .stTextInput > div > div > input {
        border-radius: 8px;
        border: 2px solid #e5e7eb;
        padding: 0.75rem;
        font-size: 0.9rem;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    /* File Uploader */
    .stFileUploader {
        background: #f9fafb;
        border: 2px dashed #d1d5db;
        border-radius: 8px;
        padding: 2rem;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #f3f4f6;
        padding: 0.5rem;
        border-radius: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background: white;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    /* Success Message */
    .stSuccess {
        background: #d1fae5;
        border: none;
        border-radius: 8px;
        color: #065f46;
        font-weight: 500;
    }
    
    /* Activity Log */
    .activity-log {
        background: #f9fafb;
        border-radius: 8px;
        padding: 1rem;
        font-family: 'Monaco', monospace;
        font-size: 0.85rem;
        line-height: 1.6;
        color: #4b5563;
        height: 400px;
        overflow-y: auto;
    }
    
    /* Data Table */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Hide Streamlit elements */
    .stDeployButton {display: none;}
    
    h1, h2, h3 {
        color: #1a1a1a;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = []
if 'activity_log' not in st.session_state:
    st.session_state.activity_log = []
if 'excel_path' not in st.session_state:
    st.session_state.excel_path = "extracted_invoices.xlsx"

def log_activity(message):
    """Add message to activity log"""
    timestamp = datetime.now().strftime('[%H:%M:%S]')
    st.session_state.activity_log.append(f"{timestamp} {message}")

def extract_field(text, field_name):
    """Extract specific field from text"""
    text_lower = text.lower()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    if field_name == 'Invoice Number':
        patterns = [
            r'quotation\s+no\.?\s*:?\s*([a-z0-9#\-/]+)',
            r'invoice\s+no\.?\s*:?\s*([a-z0-9#\-/]+)',
            r'bill\s+no\.?\s*:?\s*([a-z0-9#\-/]+)',
            r'inv\s*#?\s*:?\s*([a-z0-9#\-/]+)',
            r'#\s*([a-z]{2,}\d+)',
            r'([a-z]{3,}\s*#\s*[a-z]?\d+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                value = match.group(1)
                if value and len(value) > 2:
                    return value.upper()
    
    elif field_name == 'Date':
        patterns = [
            r'quotation\s+date\s*:?\s*(\d{1,2}\s+[a-z]+,?\s+\d{4})',
            r'invoice\s+date\s*:?\s*(\d{1,2}\s+[a-z]+,?\s+\d{4})',
            r'(\d{1,2}\s+[a-z]+,?\s+\d{4})',
            r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                return match.group(1)
    
    elif field_name == 'Company Name':
        for i, line in enumerate(lines[:10]):
            line_clean = line.strip()
            if len(line_clean) > 3 and line_clean.lower() not in ['quote', 'invoice', 'bill', 'quotation']:
                if ':' not in line_clean and 'no' not in line_clean.lower()[:5]:
                    return line_clean
    
    elif field_name == 'Customer Name':
        for i, line in enumerate(lines):
            if 'invoice to' in line.lower() or (line.lower().startswith('to') and ':' in line):
                if i + 1 < len(lines):
                    customer = lines[i + 1]
                    if customer and len(customer) > 2:
                        return customer
    
    elif field_name == 'Total Amount':
        patterns = [
            r'total\s*(?:amount)?\s*:?\s*(?:rs\.?|₹|inr)?\s*(\d[\d,]+\.?\d*)',
            r'grand\s+total\s*:?\s*(?:rs\.?|₹|inr)?\s*(\d[\d,]+\.?\d*)',
            r'net\s+total\s*:?\s*(?:rs\.?|₹|inr)?\s*(\d[\d,]+\.?\d*)',
            r'amount\s+payable\s*:?\s*(?:rs\.?|₹|inr)?\s*(\d[\d,]+\.?\d*)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                amount = match.group(1).replace(',', '')
                if float(amount) > 100:  # Filter out small numbers
                    return f"₹ {amount}"
    
    elif field_name == 'Tax Amount':
        patterns = [
            r'gst\s*\(\d+%\)\s*(?:₹|rs\.?|2)?\s*(\d+[,\d]+)',
            r'tax\s*:?\s*(?:₹|rs\.?|2)?\s*(\d+[,\d]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                amount = match.group(1)
                amount = re.sub(r'^[₹2]\s*', '', amount)
                if len(amount) >= 2 and amount[0].isdigit():
                    return '₹ ' + amount
    
    elif field_name == 'Subtotal':
        patterns = [
            r'sub\s*total\s*:?\s*(?:₹|rs\.?|2)?\s*(\d+[,\d]+)',
            r'subtotal\s*:?\s*(?:₹|rs\.?|2)?\s*(\d+[,\d]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                amount = match.group(1)
                amount = re.sub(r'^[₹2]\s*', '', amount)
                if len(amount) >= 2 and amount[0].isdigit():
                    return '₹ ' + amount
    
    return 'N/A'

def process_invoice(image_file, fields_to_extract):
    """Process invoice image and extract data"""
    try:
        image = Image.open(image_file)
        image = image.convert('L')
        
        log_activity(f"Processing: {image_file.name}")
        
        text = pytesseract.image_to_string(image, lang='eng')
        
        if not text.strip():
            log_activity(f"⚠️ No text found in {image_file.name}")
            return None
        
        log_activity(f"OCR extracted {len(text)} characters")
        
        extracted = {
            'Filename': image_file.name,
            'Extraction Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        for field in fields_to_extract:
            value = extract_field(text, field)
            extracted[field] = value
            if value != 'N/A':
                log_activity(f"  {field}: {value}")
        
        return extracted
    
    except Exception as e:
        log_activity(f"❌ Error: {str(e)}")
        return None

def convert_df_to_excel_bytes(df):
    """Convert dataframe to Excel bytes"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoice Data')
    output.seek(0)
    return output

# Hero Section
st.markdown("""
    <div class="hero-card">
        <h1 class="hero-title">✨ Invoice Extractor AI</h1>
        <p class="hero-subtitle">Smart. Fast. Accurate.</p>
        <p class="hero-description">Extract data from invoices instantly with AI-powered OCR technology</p>
    </div>
""", unsafe_allow_html=True)

# Main Layout
col1, col2 = st.columns([2, 1], gap="large")

with col1:
    # Upload & Process Section
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📤 Upload & Process</p>', unsafe_allow_html=True)
    
    # Field Selection
    st.markdown("**Select Fields to Extract:**")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        inv_num = st.checkbox('📋 Invoice Number', value=True, key='f1')
        date = st.checkbox('📅 Date', value=True, key='f2')
    with col_b:
        customer = st.checkbox('👤 Customer Name', value=True, key='f3')
        total = st.checkbox('💰 Total Amount', value=True, key='f4')
    with col_c:
        company = st.checkbox('🏢 Company Name', value=True, key='f5')
        tax = st.checkbox('📊 Tax Amount', value=False, key='f6')
    
    subtotal = st.checkbox('💵 Subtotal', value=False, key='f7')
    
    field_options = {
        'Invoice Number': inv_num,
        'Date': date,
        'Company Name': company,
        'Customer Name': customer,
        'Total Amount': total,
        'Tax Amount': tax,
        'Subtotal': subtotal
    }
    
    selected_fields = [field for field, selected in field_options.items() if selected]
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Tabs
    tab1, tab2 = st.tabs(["📁 Process Folder", "📄 Single Invoice"])
    
    with tab1:
        folder_files = st.file_uploader(
            "Upload multiple invoices",
            type=['jpg', 'jpeg', 'png', 'tiff', 'bmp'],
            accept_multiple_files=True,
            key="folder"
        )
        
        if folder_files:
            st.success(f"✅ {len(folder_files)} files ready")
            
            if st.button("🚀 Process All", type="primary", use_container_width=True):
                if not selected_fields:
                    st.error("Please select fields")
                else:
                    log_activity(f"📁 Processing {len(folder_files)} files")
                    
                    progress = st.progress(0)
                    results = []
                    
                    for i, file in enumerate(folder_files):
                        progress.progress((i + 1) / len(folder_files))
                        data = process_invoice(file, selected_fields)
                        if data:
                            results.append(data)
                    
                    progress.empty()
                    
                    if results:
                        st.session_state.extracted_data = results
                        st.balloons()
                        log_activity(f"✅ Processed {len(results)} invoices")
                        st.success(f"✅ Extracted {len(results)} invoices!")
    
    with tab2:
        single_file = st.file_uploader(
            "Upload one invoice",
            type=['jpg', 'jpeg', 'png', 'tiff', 'bmp'],
            key="single"
        )
        
        if single_file:
            image = Image.open(single_file)
            st.image(image, caption=single_file.name, use_column_width=True)
            
            if st.button("🚀 Extract Data", type="primary", use_container_width=True):
                if not selected_fields:
                    st.error("Please select fields")
                else:
                    log_activity(f"📄 Processing: {single_file.name}")
                    data = process_invoice(single_file, selected_fields)
                    
                    if data:
                        st.session_state.extracted_data = [data]
                        st.balloons()
                        st.success("✅ Data extracted!")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Excel Output
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">💾 Excel Output</p>', unsafe_allow_html=True)
    
    excel_path = st.text_input("File path:", value=st.session_state.excel_path)
    st.session_state.excel_path = excel_path
    
    col_x, col_y = st.columns(2)
    with col_x:
        if st.button("📂 Browse", use_container_width=True):
            st.info("💡 Web version: Files download to your Downloads folder automatically. Edit the filename above if needed.")
    with col_y:
        if st.button("📊 Open Excel", use_container_width=True):
            st.info("💡 Web version: Click 'Download Excel' in the results section below to get your file!")
    
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    # Activity Log
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📋 Activity Log</p>', unsafe_allow_html=True)
    
    if st.session_state.activity_log:
        log_text = "\n".join(st.session_state.activity_log[-15:])
        st.markdown(f'<div class="activity-log">{log_text}</div>', unsafe_allow_html=True)
        
        if st.button("🗑️ Clear Log", use_container_width=True):
            st.session_state.activity_log = []
            st.rerun()
    else:
        st.info("Activity will appear here...")
    
    st.markdown('</div>', unsafe_allow_html=True)

# Results Section
if st.session_state.extracted_data:
    st.markdown('<div class="clean-card">', unsafe_allow_html=True)
    st.markdown('<p class="card-title">📊 Extracted Data</p>', unsafe_allow_html=True)
    
    df = pd.DataFrame(st.session_state.extracted_data)
    st.dataframe(df, use_container_width=True, height=350)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        excel_data = convert_df_to_excel_bytes(df)
        st.download_button(
            "📥 Download Excel",
            data=excel_data,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        csv = df.to_csv(index=False)
        st.download_button(
            "📥 Download CSV",
            data=csv,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col3:
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.extracted_data = []
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("""
    <div style="text-align: center; padding: 2rem 0; color: #9ca3af; font-size: 0.85rem;">
        <p style="margin: 0;">✨ Invoice Extractor AI v2.0 | Brillspark</p>
        <p style="margin: 0;">Powered by AI • Optimized for Indian Invoices</p>
    </div>
""", unsafe_allow_html=True)
