"""
Invoice Data Extractor - Web Application
Deploy to Streamlit Cloud, Heroku, or any web hosting
No installation needed for users!
"""

import streamlit as st
import pandas as pd
from PIL import Image
import pytesseract
import re
from datetime import datetime
import io
import base64

# Page configuration
st.set_page_config(
    page_title="Invoice Data Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #2C3E50;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border-left: 5px solid #17a2b8;
        border-radius: 5px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = []

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
    
    return 'N/A'

def process_invoice(image_file, fields_to_extract):
    """Process invoice image and extract data"""
    try:
        # Open image
        image = Image.open(image_file)
        
        # Convert to grayscale for better OCR
        image = image.convert('L')
        
        # Extract text using OCR
        text = pytesseract.image_to_string(image, lang='eng')
        
        if not text.strip():
            return None
        
        # Extract data
        extracted = {
            'Filename': image_file.name,
            'Extraction Date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        for field in fields_to_extract:
            value = extract_field(text, field)
            extracted[field] = value
        
        return extracted
    
    except Exception as e:
        st.error(f"Error processing {image_file.name}: {str(e)}")
        return None

def convert_df_to_excel(df):
    """Convert dataframe to Excel file for download"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Invoice Data')
    output.seek(0)
    return output

# Main UI
st.markdown('<h1 class="main-header">📄 Invoice Data Extractor</h1>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    
    st.subheader("Fields to Extract")
    field_options = {
        'Invoice Number': st.checkbox('Invoice Number', value=True),
        'Date': st.checkbox('Date', value=True),
        'Company Name': st.checkbox('Company Name', value=True),
        'Customer Name': st.checkbox('Customer Name', value=True),
        'Total Amount': st.checkbox('Total Amount', value=True),
        'Tax Amount': st.checkbox('Tax Amount', value=False),
    }
    
    selected_fields = [field for field, selected in field_options.items() if selected]
    
    st.markdown("---")
    st.subheader("ℹ️ About")
    st.info("""
    **Invoice Data Extractor v1.0**
    
    Automatically extracts data from invoice images.
    
    Supports:
    - Indian invoice formats
    - ₹ (Rupee) symbol
    - Multiple image formats
    - Batch processing
    """)

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📁 Upload Invoices")
    
    uploaded_files = st.file_uploader(
        "Choose invoice images",
        type=['jpg', 'jpeg', 'png', 'tiff', 'bmp'],
        accept_multiple_files=True,
        help="Upload one or more invoice images (JPG, PNG, TIFF, BMP)"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        
        # Process button
        if st.button("🚀 Extract Data", type="primary", use_container_width=True):
            if not selected_fields:
                st.error("⚠️ Please select at least one field to extract")
            else:
                with st.spinner("Processing invoices..."):
                    progress_bar = st.progress(0)
                    results = []
                    
                    for i, uploaded_file in enumerate(uploaded_files):
                        # Update progress
                        progress = (i + 1) / len(uploaded_files)
                        progress_bar.progress(progress)
                        
                        # Process invoice
                        data = process_invoice(uploaded_file, selected_fields)
                        if data:
                            results.append(data)
                    
                    progress_bar.empty()
                    
                    if results:
                        st.session_state.extracted_data = results
                        st.success(f"✅ Successfully extracted data from {len(results)} invoices!")
                    else:
                        st.error("❌ Could not extract data from any images. Please check image quality.")

with col2:
    st.header("📊 Quick Stats")
    
    if st.session_state.extracted_data:
        total_invoices = len(st.session_state.extracted_data)
        
        st.metric("Total Invoices Processed", total_invoices)
        st.metric("Fields Extracted", len(selected_fields))
        
        # Show sample data
        if st.session_state.extracted_data:
            st.subheader("Latest Invoice")
            latest = st.session_state.extracted_data[-1]
            for key, value in latest.items():
                if key not in ['Filename', 'Extraction Date']:
                    st.text(f"{key}: {value}")

# Results section
if st.session_state.extracted_data:
    st.markdown("---")
    st.header("📋 Extracted Data")
    
    # Convert to dataframe
    df = pd.DataFrame(st.session_state.extracted_data)
    
    # Display data
    st.dataframe(df, use_container_width=True, height=400)
    
    # Download buttons
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Download as Excel
        excel_data = convert_df_to_excel(df)
        st.download_button(
            label="📥 Download Excel",
            data=excel_data,
            file_name=f"extracted_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # Download as CSV
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"extracted_invoices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col3:
        # Clear data
        if st.button("🗑️ Clear Data", use_container_width=True):
            st.session_state.extracted_data = []
            st.rerun()

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 2rem;'>
    <p>Invoice Data Extractor v1.0 | Optimized for Indian Invoices</p>
    <p>Supports JPG, PNG, TIFF, BMP formats | Powered by Tesseract OCR</p>
</div>
""", unsafe_allow_html=True)
