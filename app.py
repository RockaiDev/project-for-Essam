
import streamlit as st
import pandas as pd
import json
import os
import io
import zipfile

from generate_invoices import create_invoice_pdf, extract_metadata

# --- Local Storage Simulation ---
DB_FILE = "invoices_db.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, 'r') as f:
                return json.load(f)
        except:
            return []
    return []

def save_db(data):
    with open(DB_FILE, 'w') as f:
        json.dump(data, f, indent=4)

# --- PDF Generation Helper ---
def generate_pdf_bytes(df, sheet_name):
    """
    Generate PDF and return bytes instead of saving to disk.
    This creates a temp file then reads it back (ReportLab writes to file object).
    """
    buffer = io.BytesIO()
    # Create temp filename
    temp_filename = f"temp_invoice_{sheet_name}.pdf"
    create_invoice_pdf(df, sheet_name, temp_filename)
    
    with open(temp_filename, "rb") as f:
        pdf_bytes = f.read()
    
    # Clean up temp file
    if os.path.exists(temp_filename):
        os.remove(temp_filename)
        
    return pdf_bytes

# --- Main App ---
st.set_page_config(page_title="Rockai Invoice System", layout="wide", page_icon="🚀")

# --- Custom Styling (Rockai Dev Identity) ---
st.markdown("""
<style>
    /* Global Dark Theme */
    .stApp {
        background-color: #0E1117;
        color: #FAFAFA;
    }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #161B22;
        border-right: 1px solid #30363D;
    }
    
    /* Headings */
    h1, h2, h3 {
        color: #58A6FF; /* Rockai Blue */
        font-family: 'Segoe UI', sans-serif;
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        color: #A371F7 !important; /* Rockai Purple */
    }
    
    /* Buttons */
    .stButton > button {
        background-color: #238636;
        color: white;
        border: none;
        border-radius: 6px;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2EA043;
        box-shadow: 0 0 10px rgba(46, 160, 67, 0.5);
    }
    
    /* Tables */
    [data-testid="stDataFrame"] {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #30363D;
    }
    
    /* Custom Header */
    .rockai-header {
        text-align: center;
        padding: 2rem 0;
        background: linear-gradient(90deg, #0E1117 0%, #1F242C 50%, #0E1117 100%);
        border-bottom: 2px solid #58A6FF;
        margin-bottom: 2rem;
    }
    .rockai-logo {
        font-size: 2.5rem;
        font-weight: bold;
        background: -webkit-linear-gradient(45deg, #58A6FF, #A371F7);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
</style>
<div class="rockai-header">
    <div class="rockai-logo">ROCKAI DEV <span style="font-size: 1rem; color: #8B949E; -webkit-text-fill-color: #8B949E;">| INVOICE SYSTEM</span></div>
</div>
""", unsafe_allow_html=True)

# Sidebar with better navigation
st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Page", ["Dashboard", "Upload New Data", "Settings"], label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.markdown("© 2026 Rockai Dev")
st.sidebar.markdown("Powered by *Advanced Agentic AI*")

if page == "Dashboard":
    st.subheader("📊 Invoices Dashboard")
    db = load_db()
    
    if not db:
        st.info("No invoices found. Go to 'Upload New Data' to add some.")
    else:
        # Download All Button
        col1, col2 = st.columns([3, 1])
        with col2:
            if st.button("📥 Download All Invoices", type="primary"):
                # Create ZIP file in memory
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for inv in db:
                        pdf_path = inv.get('pdf_path')
                        if pdf_path and os.path.exists(pdf_path):
                            # Add file to ZIP with its basename
                            zip_file.write(pdf_path, os.path.basename(pdf_path))
                
                zip_buffer.seek(0)
                st.download_button(
                    label="Click to Download ZIP",
                    data=zip_buffer,
                    file_name="all_invoices.zip",
                    mime="application/zip"
                )
        
        # Create DataFrame for display
        display_data = []
        for inv in db:
            meta = inv.get('metadata', {})
            display_data.append({
                "Invoice No": meta.get('Invoice No', '-'),
                "Patient Name": meta.get('Patient Name', '-'),
                "Amount": f"{inv.get('grand_total', 0):,.2f}",
                "Date": meta.get('Admission Date', '-'),
                "Sheet": inv.get('sheet_name', '-'),
                "ID": inv.get('id')
            })
            
        df_display = pd.DataFrame(display_data)
        st.dataframe(df_display, width='stretch')
        
        # Detail View
        st.divider()
        st.subheader("Invoice Details & Action")
        selected_inv_no = st.selectbox("Select Invoice to View", [d['Invoice No'] for d in display_data])
        
        if selected_inv_no:
            # Find the full record
            record = next((item for item in db if item.get('metadata', {}).get('Invoice No') == selected_inv_no), None)
            
            if record:
                col1, col2 = st.columns(2)
                with col1:
                    st.json(record.get('metadata'))
                
                with col2:
                    st.metric("Total Amount", f"{record.get('grand_total', 0):,.2f} EGP")
                    
                    pdf_path = record.get('pdf_path')
                    if pdf_path and os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()
                        
                        st.download_button(
                            label="Download PDF Invoice",
                            data=pdf_bytes,
                            file_name=os.path.basename(pdf_path),
                            mime="application/pdf"
                        )
                    else:
                        st.warning("PDF file not found. Please go to 'Upload New Data' to regenerate.")
                    
                    # Re-upload feature is better for PDF generation unless we store everything.

if page == "Upload New Data":
    st.subheader("Upload Excel File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])
    
    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            st.success(f"Loaded {len(sheet_names)} sheets.")
            
            if st.button("Process & Save to Local Storage"):
                current_db = load_db()
                new_count = 0
                
                progress_bar = st.progress(0)
                
                for idx, sheet in enumerate(sheet_names):
                    df = pd.read_excel(xls, sheet_name=sheet)
                    
                    # 1. Find Header
                    header_row_idx = None
                    for i, row in df.iterrows():
                        vals = [str(x).lower().strip() for x in row.values if pd.notna(x)]
                        if 'description' in vals and ('qty' in vals or 'unit' in vals):
                            header_row_idx = i
                            break
                    
                    if header_row_idx is not None:
                        # 2. Extract Metadata
                        meta = extract_metadata(df, header_row_idx)
                        
                        # 3. Calculate Total
                        grand_total = 0.0
                        
                        # We need header map again to find net column
                        header_map = {}
                        if header_row_idx is not None:
                            header_row = df.iloc[header_row_idx]
                            for col_idx, val in enumerate(header_row):
                                if pd.notna(val):
                                    header_map[str(val).strip()] = col_idx
                                    
                        # iterate rows after header
                        col_desc = header_map.get('Description')
                        col_net = header_map.get('Debit')
                        col_date = header_map.get('Date')
                        
                        for i in range(header_row_idx + 1, len(df)):
                            row = df.iloc[i]
                            
                            def get_val(r, idx):
                                if idx is None: return ""
                                v = r.iloc[idx]
                                return v if pd.notna(v) else ""

                            desc = str(get_val(row, col_desc)).strip()
                            net = get_val(row, col_net)
                            date_col_val = str(get_val(row, col_date)).strip()
        
                            if 'Grand Total' in date_col_val or 'Grand Total' in desc:
                                 try:
                                     grand_total = float(str(net).replace(',',''))
                                 except:
                                     grand_total = 0.0
                                 break
                        
                        # Let's save the metadata
                        inv_record = {
                            "id": f"{sheet}_{meta.get('Invoice No')}",
                            "sheet_name": sheet,
                            "metadata": meta,
                            "grand_total": grand_total
                        }
                        
                        # Filename Logic matching generate_invoices.py
                        visit_no = meta.get('Visit No')
                        inv_num = meta.get('Invoice No')
                        
                        if visit_no and visit_no != '-':
                            filename_base = f"Visit_{visit_no}"
                        elif inv_num and inv_num != '-':
                            filename_base = f"Invoice_{inv_num}"
                        else:
                            filename_base = f"Unknown_Sheet_{sheet}"
                        
                        filename_base = filename_base.replace('/', '-').replace('\\', '-')
                        save_path = f"invoices/{filename_base}.pdf"
                        
                        # Generate PDF immediately to bytes (using new design from generate_invoices)
                        # We can just call create_invoice_pdf directly to file
                        if not os.path.exists('invoices'): os.makedirs('invoices')
                        create_invoice_pdf(df, sheet, save_path)
                            
                        inv_record['pdf_path'] = save_path
                        
                        # Check duplicate
                        # Remove old record with same ID if exists
                        current_db = [d for d in current_db if d['id'] != inv_record['id']]
                        current_db.append(inv_record)
                        new_count += 1
                    
                    progress_bar.progress((idx + 1) / len(sheet_names))
                
                save_db(current_db)
                st.success(f"Successfully processed {new_count} new invoices!")
                
        except Exception as e:
            st.error(f"Error processing file: {e}")

if page == "Settings":
    st.write("Settings")
    if st.button("Clear Invoice Database"):
        save_db([])
        st.success("Database cleared.")

