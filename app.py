"""
Invoice Data Extractor - Complete Version
Modern UI + All Desktop Features
No Settings Panel, All Working Features
"""

import streamlit as st
import pandas as pd
from PIL import Image
import pytesseract
import re
from datetime import datetime
import io
import os

# Page configuration
st.set_page_config(
    page_title="Invoice Extractor - Brillspark",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Modern CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        background-attachment: fixed;
    }
    
    .hero-section {
        background: #5a67d8;
        backdrop-filter: blur(10px);
        border-radius: 30px;
        padding: 3rem 2rem;
        text-align: center;
        margin-bottom: 2rem;
        border: 1px solid rgba(255, 255, 255, 0.3);
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
    }
    
    .hero-title {
        font-size: 3.5rem;
        font-weight: 800;
        color: white;
        margin-bottom: 1rem;
        letter-spacing: -2px;
    }
    
    .glass-card {
        background: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 2rem;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
        margin-bottom: 2rem;
    }
    
    .section-title {
        color: white;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .stButton > button {
        background: white !important;
        color: #667eea !important;
        border: 2px solid #667eea !important;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background: #667eea !important;
        color: white !important;
        transform: translateY(-2px);
    }
    
    /* Primary action buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
    }
    
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 25px rgba(102, 126, 234, 0.6);
    }
    
    .success-banner {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
        font-weight: 600;
        box-shadow: 0 4px 20px rgba(56, 239, 125, 0.3);
        margin: 1rem 0;
    }
    
    h1, h2, h3, h4, h5, h6, p, label, .stMarkdown {
        color: white !important;
    }
    
    .stCheckbox {
        color: white !important;
    }
    
    /* Modern checkbox styling */
    .stCheckbox > label {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 10px;
        padding: 10px 14px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .stCheckbox > label:hover {
        background: rgba(255, 255, 255, 0.1);
        border-color: rgba(255, 255, 255, 0.3);
        transform: translateY(-1px);
    }
    
    .stCheckbox input[type="checkbox"]:checked + label {
        background: rgba(102, 126, 234, 0.15);
        border-color: rgba(102, 126, 234, 0.5);
    }
    
    .section-divider {
        height: 2px;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
        margin: 2rem 0;
    }
    
    .stTextInput > div > div > input {
        background: rgba(255, 255, 255, 0.1);
        color: white;
        border: 1px solid rgba(255, 255, 255, 0.3);
        border-radius: 10px;
    }
    
    .badge {
        display: inline-block;
        padding: 0.5rem 1rem;
        background: rgba(255, 255, 255, 0.25);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        color: white;
        border: 1px solid rgba(255, 255, 255, 0.4);
        margin: 0.25rem;
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
    """Extract specific field from text - Optimized for Indian invoices"""
    text_lower = text.lower()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    if field_name == 'Invoice Number':
        patterns = [
            r'quotation\s+no\.?\s*:?\s*(no\.\s*)?([a-z0-9#\-/]+)',
            r'invoice\s+no\.?\s*:?\s*(no\.\s*)?([a-z0-9#\-/]+)',
            r'bill\s+no\.?\s*:?\s*(no\.\s*)?([a-z0-9#\-/]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                value = match.group(2) if match.lastindex >= 2 else match.group(1)
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
            r'total\s*(?:\n|:)?\s*(?:₹|rs\.?|inr|2)?\s*(\d+[,\d]+)',
            r'(?:₹|rs\.?|2)\s*(\d+[,\d]+)',
            r'grand\s+total\s*:?\s*(?:₹|rs\.?|inr|2)?\s*(\d+[,\d]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                amount = match.group(1)
                amount = re.sub(r'^[₹2]\s*', '', amount)
                if len(amount) >= 3 and amount[0].isdigit():
                    return '₹ ' + amount
    
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
        log_activity(f"❌ Error processing {image_file.name}: {str(e)}")
        return None

def convert_df_to_excel_bytes(df):
    """Convert dataframe to Excel bytes for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoice Data')
    output.seek(0)
    return output

# Hero Section
st.markdown("""
    <div class="hero-section">
        <h1 class="hero-title">✨ Invoice Extractor AI</h1>
        <p style="color: white; font-size: 1.3rem; font-weight: 400; margin-bottom: 0.5rem;">Smart. Fast. Accurate.</p>
        <p style="color: rgba(255, 255, 255, 0.9); font-size: 1rem;">
            Extract data from invoices instantly with AI-powered OCR technology
        </p>
        <br>
        <span class="badge">🚀 Powered by AI</span>
        <span class="badge">🇮🇳 Indian Format Support</span>
        <span class="badge">⚡ Lightning Fast</span>
    </div>
""", unsafe_allow_html=True)

st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)

# Main Layout
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>📤 Upload & Process</h3>", unsafe_allow_html=True)
    
    # Field selection - Modern Pill Style
    st.markdown("**Select Fields to Extract:**")
    st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)
    
    # Row 1
    col1, col2, col3 = st.columns(3)
    with col1:
        inv_num = st.checkbox('📋 Invoice Number', value=True, key='inv_num')
    with col2:
        customer = st.checkbox('👤 Customer Name', value=True, key='customer')
    with col3:
        tax = st.checkbox('📊 Tax Amount', value=False, key='tax')
    
    # Row 2
    col4, col5, col6 = st.columns(3)
    with col4:
        date = st.checkbox('📅 Date', value=True, key='date')
    with col5:
        total = st.checkbox('💰 Total Amount', value=True, key='total')
    with col6:
        subtotal = st.checkbox('💵 Subtotal', value=False, key='subtotal')
    
    # Row 3
    col7, col8, col9 = st.columns(3)
    with col7:
        company = st.checkbox('🏢 Company Name', value=True, key='company')
    
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
    
    # Upload options
    tab1, tab2 = st.tabs(["📁 Process Folder", "📄 Process Single Invoice"])
    
    with tab1:
        st.markdown("**Upload multiple invoice images:**")
        folder_files = st.file_uploader(
            "Choose invoice images",
            type=['jpg', 'jpeg', 'png', 'tiff', 'bmp'],
            accept_multiple_files=True,
            key="folder_upload",
            label_visibility="collapsed"
        )
        
        if folder_files:
            st.success(f"✅ {len(folder_files)} file(s) ready to process")
            
            if st.button("🚀 Process All Invoices", type="primary", use_container_width=True):
                if not selected_fields:
                    st.error("⚠️ Please select at least one field to extract")
                else:
                    log_activity(f"📁 Processing folder with {len(folder_files)} files")
                    
                    with st.spinner("🔮 Processing invoices..."):
                        progress_bar = st.progress(0)
                        results = []
                        
                        for i, file in enumerate(folder_files):
                            progress = (i + 1) / len(folder_files)
                            progress_bar.progress(progress)
                            
                            data = process_invoice(file, selected_fields)
                            if data:
                                results.append(data)
                        
                        progress_bar.empty()
                        
                        if results:
                            st.session_state.extracted_data = results
                            st.balloons()
                            log_activity(f"✅ Successfully processed {len(results)} invoices!")
                            st.markdown(f"""
                                <div class="success-banner">
                                    🎉 Success! Extracted data from {len(results)} invoices
                                </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.error("❌ Could not extract data from any images")
    
    with tab2:
        st.markdown("**Upload a single invoice image:**")
        single_file = st.file_uploader(
            "Choose one invoice",
            type=['jpg', 'jpeg', 'png', 'tiff', 'bmp'],
            key="single_upload",
            label_visibility="collapsed"
        )
        
        if single_file:
            col_img, col_btn = st.columns([1, 1])
            
            with col_img:
                image = Image.open(single_file)
                st.image(image, caption=single_file.name, use_column_width=True)
            
            with col_btn:
                if st.button("🚀 Extract Data", type="primary", use_container_width=True):
                    if not selected_fields:
                        st.error("⚠️ Please select at least one field")
                    else:
                        log_activity(f"📄 Processing single invoice: {single_file.name}")
                        
                        with st.spinner("🔮 Extracting data..."):
                            data = process_invoice(single_file, selected_fields)
                            
                            if data:
                                st.session_state.extracted_data = [data]
                                st.balloons()
                                log_activity("✅ Data extracted successfully!")
                                st.success("✅ Data extracted and saved!")
                            else:
                                st.error("❌ Could not extract data")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Excel Output Section
    st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>💾 Excel Output File</h3>", unsafe_allow_html=True)
    
    excel_path = st.text_input(
        "Excel file path:",
        value=st.session_state.excel_path,
        help="Path where Excel file will be saved"
    )
    st.session_state.excel_path = excel_path
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("📂 Browse", use_container_width=True):
            st.info("💡 In web version, files download to your Downloads folder. Enter desired filename above.")
    
    with col_btn2:
        if st.button("📊 Open Excel", use_container_width=True):
            st.info("💡 Use the Download buttons in the results section below to get your Excel file!")
    
    st.markdown("</div>", unsafe_allow_html=True)

with col2:
    # Activity Log
    st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>📋 Activity Log</h3>", unsafe_allow_html=True)
    
    if st.session_state.activity_log:
        log_text = "\n".join(st.session_state.activity_log[-20:])  # Last 20 entries
        st.text_area(
            "Activity",
            value=log_text,
            height=400,
            label_visibility="collapsed"
        )
        
        if st.button("🗑️ Clear Log", use_container_width=True):
            st.session_state.activity_log = []
            st.rerun()
    else:
        st.info("📝 Activity will appear here...")
    
    st.markdown("</div>", unsafe_allow_html=True)

# Results Section
if st.session_state.extracted_data:
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
    st.markdown("<h3 class='section-title'>📊 Extracted Data</h3>", unsafe_allow_html=True)
    
    df = pd.DataFrame(st.session_state.extracted_data)
    st.dataframe(df, use_container_width=True, height=400)
    
    # Download buttons
    col1, col2, col3 = st.columns(3)
    
    with col1:
        excel_data = convert_df_to_excel_bytes(df)
        st.download_button(
            label="📥 Download Excel",
            data=excel_data,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"brillspark_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col3:
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.extracted_data = []
            log_activity("🗑️ Cleared extracted data")
            st.rerun()
    
    st.markdown("</div>", unsafe_allow_html=True)

# Footer
st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
st.markdown("""
    <div style='text-align: center; padding: 2rem; color: rgba(255,255,255,0.7);'>
        <p style='font-size: 0.9rem;'>✨ Invoice Extractor AI v2.0 | Brillspark</p>
        <p style='font-size: 0.8rem;'>Powered by AI • Optimized for Indian Invoices • Built with ❤️</p>
    </div>
""", unsafe_allow_html=True)
