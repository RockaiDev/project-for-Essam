import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display
import os
import datetime

# --- Configuration ---
# Look for local font first (for deployment)
if os.path.exists('fonts/Arial Unicode.ttf'):
    FONT_PATH = 'fonts/Arial Unicode.ttf'
elif os.path.exists('/Library/Fonts/Arial Unicode.ttf'):
    FONT_PATH = '/Library/Fonts/Arial Unicode.ttf'
else:
    FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'

PRIMARY_COLOR = colors.HexColor('#005596') # Medical Blue
SECONDARY_COLOR = colors.HexColor('#F2F2F2') # Light Gray
TEXT_COLOR = colors.Color(0.2, 0.2, 0.2)
BORDER_COLOR = colors.Color(0.8, 0.8, 0.8)

# --- 1. Font Setup ---
def setup_fonts():
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont('ArialUnicode', FONT_PATH))
        # Register the same font for Bold to support Arabic in bold tags
        pdfmetrics.registerFont(TTFont('ArialUnicode-Bold', FONT_PATH))
        
        from reportlab.lib.fonts import addMapping
        addMapping('ArialUnicode', 0, 0, 'ArialUnicode') # normal
        addMapping('ArialUnicode', 1, 0, 'ArialUnicode-Bold') # bold (1=bold, 0=italic)
        
        return 'ArialUnicode'
    else:
        print("Warning: Arial Unicode font not found. Arabic text will not render correctly.")
        return 'Helvetica'

FONT_NAME = setup_fonts()

# --- 2. Styles ---
def get_stylesheet():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='InvoiceTitle', fontName=FONT_NAME, fontSize=24, leading=28, textColor=PRIMARY_COLOR, spaceAfter=20))
    styles.add(ParagraphStyle(name='CompanyInfo', fontName=FONT_NAME, fontSize=10, leading=12, textColor=TEXT_COLOR))
    styles.add(ParagraphStyle(name='SectionHeader', fontName=FONT_NAME, fontSize=12, leading=14, textColor=PRIMARY_COLOR, spaceAfter=5))
    styles.add(ParagraphStyle(name='NormalText', fontName=FONT_NAME, fontSize=9, leading=11, textColor=TEXT_COLOR))
    styles.add(ParagraphStyle(name='BoldText', fontName=FONT_NAME, fontSize=9, leading=11, textColor=TEXT_COLOR)) # Use bold font if available
    styles.add(ParagraphStyle(name='FooterText', fontName=FONT_NAME, fontSize=8, leading=10, textColor=colors.grey))
    return styles

def process_text(text):
    """Reshape Arabic text and handle direction."""
    if text is None or pd.isna(text):
        return ""
    
    text_str = str(text).strip()
    if not text_str:
        return ""
        
    # Check if text contains Arabic characters
    has_arabic = any('\u0600' <= char <= '\u06FF' for char in text_str)
    
    if has_arabic:
        try:
            # Reshape characters (connecting them)
            reshaped_text = arabic_reshaper.reshape(text_str)
            # Reorder for RTL display
            bidi_text = get_display(reshaped_text)
            return bidi_text
        except Exception as e:
            # Fallback to original text if library fails
            return text_str
            
    return text_str

# --- 4. Data Extraction ---
def extract_metadata(df, stop_row_idx):
    """Scan top rows for metadata using multiple checks."""
    metadata = {}
    
    # Default values
    metadata['Hospital Name'] = 'Andalusia Hospitals Smouha'
    metadata['Address'] = '35 Bahaa El Din Ghatoury St, Smouha, Alexandria'
    metadata['VAT No'] = '202471187'
    
    # Iterate through potential header rows
    # Iterate through potential header rows
    for i, row in df.head(stop_row_idx).iterrows():
        # Scan row for keys dynamically
        # Convert row to list of strings
        row_vals = [str(x).strip() for x in row.values]
        
        # Helper to find value next to key
        def find_val(key_list, search_row_vals):
            for k_idx, cell_val in enumerate(search_row_vals):
                for key in key_list:
                    if key.lower() in cell_val.lower():
                        # Found key at k_idx. Look for value in next few cells.
                        # Usually value is immediate next non-empty, or at specific offset.
                        # Let's look at k_idx + 1 up to k_idx + 5
                        for offset in range(1, 6):
                            if k_idx + offset < len(search_row_vals):
                                candidate = search_row_vals[k_idx + offset]
                                if candidate and candidate.lower() != 'nan' and candidate.strip() != '' and candidate.strip() != '-':
                                     return candidate
            return None

        # Patient Name
        if 'Patient Name' not in metadata:
            val = find_val(['patient name', 'اسم المريض'], row_vals)
            if val: metadata['Patient Name'] = val
            
        # Invoice
        if 'Invoice No' not in metadata:
            val = find_val(['invoice', 'رقم الفاتورة'], row_vals)
            # Ensure it's not the label itself repeated
            if val and 'invoice' not in val.lower(): metadata['Invoice No'] = val
            
        # Visit
        if 'Visit No' not in metadata:
             val = find_val(['visit', 'رقم الزيارة'], row_vals)
             if val: metadata['Visit No'] = val

        # File No
        if 'File No' not in metadata:
             val = find_val(['file no', 'رقم الملف', 'file number'], row_vals)
             if val: metadata['File No'] = val

        # Physician
        if 'Physician' not in metadata:
             val = find_val(['physician', 'doctor', 'الهيب'], row_vals) # 'الهيب' might be typo for tabib? 'الطبية'? 
             # Let's stick to English 'Physician'
             if val: metadata['Physician'] = val

        # Admission
        if 'Admission Date' not in metadata:
             val = find_val(['date of admission', 'admission date', 'تاريخ الدخول'], row_vals)
             if val: metadata['Admission Date'] = val

        # Discharge
        if 'Discharge Date' not in metadata:
             val = find_val(['date of discharge', 'discharge date', 'تاريخ الخروج'], row_vals)
             if val: metadata['Discharge Date'] = val
             
        # Insurer
        if 'Insurer' not in metadata:
             # Insurer often appears twice? Rows 21, 24?
             # 'Insurer' label -> Value
             val = find_val(['insurer', 'المؤمن'], row_vals)
             if val and val.lower() != 'patient': metadata['Insurer'] = val

        # Nationality
        if 'Nationality' not in metadata:
             val = find_val(['nationality', 'الجنسية'], row_vals)
             if val: metadata['Nationality'] = val

        # VAT Number
        if 'VAT No' not in metadata or metadata['VAT No'] == '202471187': # Check if we can find a better one
             val = find_val(['vat no', 'tax id', 'registration no', 'التسجيل الضريبي'], row_vals)
             if val: metadata['VAT No'] = val

    # Hardcoded fixes for known problematic layouts if dynamic fails
    # Based on Audit:
    # Row 17: Invoice (Col 5) -> Val (Col 9)
    # Row 17: Visit (Col 21) -> Val (Col 24)
    # Row 22: Physician (Col 5) if exists?
    
    return metadata

# --- 5. PDF Generation ---
def create_invoice_pdf(df, sheet_name, output_path):
    # Setup Document
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=0.5*cm, leftMargin=0.5*cm,
                            topMargin=0.5*cm, bottomMargin=1*cm,
                            title=f"Invoice {sheet_name}")
    styles = get_stylesheet()
    elements = []
    
    # 5.1 Locate Table Header & Extract Metadata
    header_row_idx = None
    header_map = {}
    for i, row in df.iterrows():
        vals = [str(x).lower().strip() for x in row.values if pd.notna(x)]
        if 'description' in vals and ('qty' in vals or 'unit' in vals):
            header_row_idx = i
            for col_idx, val in enumerate(row):
                if pd.notna(val):
                    header_map[str(val).strip()] = col_idx
            break
            
    if header_row_idx is None:
        print(f"Skipping {sheet_name}: Could not find item table.")
        return

    metadata = extract_metadata(df, header_row_idx)
    inv_no = metadata.get('Invoice No', '-')
    
    # --- UI CONSTANTS ---
    RED_COLOR = colors.HexColor('#A6192E') # Andalusia Red approx
    GRAY_BG = colors.HexColor('#E0E0E0')
    DARK_GRAY_BG = colors.HexColor('#B0B0B0')
    BORDER_COLOR = colors.black
    
    # --- HEADER SECTION ---
    # Layout: [Left Text] [Center Logo] [Right Text]
    
    # Styles
    h_style = ParagraphStyle('H', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10)
    addr_style = ParagraphStyle('A', parent=styles['Normal'], fontName=FONT_NAME, fontSize=9, alignment=2) # Right align
    
    # Logo
    logo_img = []
    if os.path.exists('Picture1.png'):
        logo_img = Image('Picture1.png', width=5*cm, height=2.5*cm, kind='proportional')
        
    # Hospital Name (Left)
    hosp_name = Paragraph(f"<b>{metadata.get('Hospital Name', 'Andalusia Hospitals Smouha')}</b>", h_style)
    
    # Address/VAT (Right)
    vat_txt = Paragraph(f"<b>VAT No:</b> {metadata.get('VAT No', '202471187')}", addr_style)
    addr_txt = Paragraph(process_text(metadata.get('Address Arabic', '')), addr_style)
    
    t_header_top = Table([
        [hosp_name, logo_img, [vat_txt, addr_txt]]
    ], colWidths=[6*cm, 7*cm, 6*cm])
    
    t_header_top.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (1,0), (1,0), 'CENTER'), # Center Logo
        ('ALIGN', (2,0), (2,0), 'RIGHT'),
    ]))
    elements.append(t_header_top)
    
    # Red Horizontal Line
    elements.append(Spacer(1, 0.2*cm))
    elements.append(Table([['']], colWidths=[19.5*cm], rowHeights=[3], style=[('BACKGROUND', (0,0), (-1,-1), RED_COLOR)]))
    
    # Title Bar
    title_para = Paragraph(f"DISCHARGE INVOICE&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{process_text('فاتورة خروج مريض')}", 
                           ParagraphStyle('T', fontName=FONT_NAME, fontSize=12, alignment=1, spaceBefore=4, spaceAfter=4))
    elements.append(title_para)
    
    # --- METADATA GRID ---
    # Columns: [Eng Label (2.5cm), Value (4cm), Arb Label (2.5cm)] x 2 (Left/Right sides)
    
    label_style = ParagraphStyle('L', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8)
    val_style = ParagraphStyle('V', parent=styles['Normal'], fontName=FONT_NAME, fontSize=8)
    arb_label_style = ParagraphStyle('AL', parent=styles['Normal'], fontName=FONT_NAME, fontSize=8, alignment=2)
    
    def meta_row(eng1, val1_key, arb1, eng2, val2_key, arb2):
        v1 = process_text(metadata.get(val1_key, '-'))
        v2 = process_text(metadata.get(val2_key, '-'))
        return [
            Paragraph(eng1, label_style), Paragraph(v1, val_style), Paragraph(process_text(arb1), arb_label_style),
            "", # Spacer col
            Paragraph(eng2, label_style), Paragraph(v2, val_style), Paragraph(process_text(arb2), arb_label_style)
        ]
        
    meta_data = [
        # Headers logic handled by wrapper table borders
        meta_row("Invoice", "Invoice No", "رقم الفاتورة", "Visit", "Visit No", "رقم الزيارة"),
        meta_row("Date of Admission", "Admission Date", "تاريخ الدخول", "Date of Discharge", "Discharge Date", "تاريخ الخروج"),
        meta_row("Patient Name", "Patient Name", "اسم المريض", "File No", "File No", "رقم ملف المريض"), # File No repeated? In image, Arabic Name is separate.
        # Image shows: Patient Name (Eng) | File No. (Eng) || Patient Name (Arb) | File No (Arb Label)
        # We will simplify to fit our data:
        meta_row("Insurance Card", "Insurer Card", "-", "Nationality", "Nationality", "الجنسية"),
        meta_row("Insurer", "Insurer", "المؤمن", "Contract", "Contract", "العقد"),
        meta_row("Physician", "Physician", "الطبيب المعالج", "Department", "Department", "القسم"),
        meta_row("Room No.", "Room No", "رقم الغرفة", "", "", "")
    ]
    
    # Need column widths that sum to ~19.5cm
    # Col 1 (2.5), Col 2 (5.0), Col 3 (2.0) | Gap (0.5) | Col 4 (2.5), Col 5 (5.0), Col 6 (2.0)
    col_widths = [2.5*cm, 5.0*cm, 2.0*cm, 0.5*cm, 2.5*cm, 5.0*cm, 2.0*cm]
    
    t_meta = Table(meta_data, colWidths=col_widths)
    t_meta.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, BORDER_COLOR), # Outer Box
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 0.5*cm))
    
    # --- ITEMS TABLE ---
    # Header Structure matches image
    # Row 1: Spacer | Spacer | Spacer | Spacer | Spacer | Spacer | Patient (2 cols) | Insurer (2 cols)
    # Row 2: Description | Qty | Unit | Date | Total | Discount | Debit | Credit | Debit | Credit
    
    # Our data assumes mostly Patient Debit (Net). We will put 'Net' in 'Patient Debit' and 0 in others for now.
    
    # Define Column Widths
    # Desc(6), Qty(1.2), Unit(1.2), Date(2), Total(2), Disc(1.5), P.Deb(1.5), P.Cred(1.5), I.Deb(1.5), I.Cred(1.5) -> Total ~19.9cm
    item_col_widths = [5.5*cm, 1.2*cm, 1.2*cm, 2.2*cm, 2.0*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm]
    
    # Styles
    th_style = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=8, alignment=1)
    td_style = ParagraphStyle('TD', fontName=FONT_NAME, fontSize=8, alignment=1) # Center align for numbers
    td_desc_style = ParagraphStyle('TDD', fontName=FONT_NAME, fontSize=8, alignment=0) # Left align desc
    
    def p_th(txt): return Paragraph(txt, th_style)
    def p_td(txt, s=td_style): return Paragraph(str(txt), s)

    # 1. Super Header
    # We construct the whole table data list first
    table_data = []
    
    # Row 0: Super Headers
    # We will use spans for Patient and Insurer
    table_data.append([
        "", "", "", "", "", "", 
        p_th("Patient"), "", # Spans next
        p_th("Insurer"), ""  # Spans next
    ])
    
    # Row 1: Main Headers
    table_data.append([
        p_th("Description"), p_th("Qty"), p_th("Unit"), p_th("Date"),
        p_th("Total"), p_th("Discount"),
        p_th("Debit"), p_th("Credit"),
        p_th("Debit"), p_th("Credit")
    ])
    
    # Data Rows
    col_desc = header_map.get('Description')
    col_qty = header_map.get('Qty')
    col_total = header_map.get('Total') # Or Unit Price
    col_disc = header_map.get('Discount')
    col_net = header_map.get('Debit') # Or Net Amount
    col_date = header_map.get('Date')
    col_unit = header_map.get('Unit') # Try to map if exists
    
    grand_total = 0.0
    subtotal_rows = []
    
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        
        def get_v(c): 
             v = row.iloc[c] if c is not None and pd.notna(row.iloc[c]) else ""
             return v
             
        desc_raw = str(get_v(col_desc)).strip()
        date_val = str(get_v(col_date)).strip()
        
        # Stop at Grand Total
        if 'Grand Total' in desc_raw or 'Grand Total' in date_val:
            try:
                grand_total = float(str(get_v(col_net)).replace(',',''))
            except:
                grand_total = 0.0
            break
            
        if not desc_raw and not get_v(col_qty) and not get_v(col_net): continue
        
        # Formatting
        qty = get_v(col_qty)
        qty_str = f"{qty:.3f}" if isinstance(qty, float) else str(qty)
        
        price = get_v(col_total)
        price_str = f"{price:.3f}" if isinstance(price, (int,float)) else str(price)
        
        disc = get_v(col_disc)
        disc_str = f"{disc:.3f}" if isinstance(disc, (int,float)) else str(disc)
        
        net = get_v(col_net)
        net_str = f"{net:.3f}" if isinstance(net, (int,float)) else str(net)
        
        unit = get_v(col_unit) # Likely empty/Each
        
        # Identify Row Type
        is_subtotal = 'total' in desc_raw.lower()
        is_header = (desc_raw and qty == "")
        
        if is_subtotal:
            # Subtotal Row
            row_items = [
                Paragraph(f"<b>{process_text(desc_raw)}</b>", td_desc_style),
                "", "", "", 
                Paragraph(f"<b>{price_str}</b>", td_style),
                "", "", "", "", "" # Just showing total price ideally, or net
            ]
            # If net is available, put it in Patient Debit column?
            if net: row_items[6] = Paragraph(f"<b>{net_str}</b>", td_style)
            
            subtotal_rows.append(len(table_data))
            table_data.append(row_items)
            
        elif is_header:
            # Category Header (e.g. Medical Services)
            # Gray background, bold text
            table_data.append([
                Paragraph(f"<b>{process_text(desc_raw)}</b>", td_desc_style),
                "", "", "", "", "", "", "", "", ""
            ])
            subtotal_rows.append(len(table_data)-1) # Reuse subtotal style for gray bg
            
        else:
            # Standard Item
            table_data.append([
                Paragraph(process_text(desc_raw), td_desc_style),
                p_td(qty_str),
                p_td(unit),
                p_td(date_val),
                p_td(price_str),
                p_td(disc_str),
                p_td(net_str), # Patient Debit
                p_td("0.000"), # Patient Credit
                p_td(""),      # Insurer Debit
                p_td("")       # Insurer Credit (Assumed)
            ])

    # Grand Total Row
    gt_str = f"{grand_total:,.3f}"
    table_data.append([
        Paragraph("<b>Grand Total</b>", ParagraphStyle('GT', alignment=2, fontName='Helvetica-Bold', fontSize=9)),
        "", "", "", "",
        Paragraph(f"<b>{gt_str}</b>", td_style), # Discount column? No, match headers
        Paragraph(f"<b>{gt_str}</b>", td_style), # Debit column
        "", "", ""
    ])
    
    # Build Items Table
    t_items = Table(table_data, colWidths=item_col_widths, repeatRows=2)
    
    # Style
    # Base Style
    item_tbl_style = [
        ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        
        # Super Header Styling
        ('SPAN', (6,0), (7,0)), # Patient spans 2
        ('SPAN', (8,0), (9,0)), # Insurer spans 2
        ('BACKGROUND', (0,0), (-1,1), GRAY_BG), # Headers Gray
        ('ALIGN', (0,0), (-1,1), 'CENTER'),
        
        # Grand Total Styling (Last Row)
        ('SPAN', (0,-1), (4,-1)), # Label spans first 5 cols
        ('BACKGROUND', (0,-1), (-1,-1), DARK_GRAY_BG),
    ]
    
    # Gray backgrounds for subtotals/headers
    for r_idx in subtotal_rows:
        item_tbl_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), GRAY_BG))
        
    t_items.setStyle(TableStyle(item_tbl_style))
    elements.append(t_items)
    
    # Footer
    elements.append(Spacer(1, 1.0*cm))
    
    # Red Line Bottom
    elements.append(Table([['']], colWidths=[19.5*cm], rowHeights=[2], style=[('BACKGROUND', (0,0), (-1,-1), RED_COLOR)]))
    elements.append(Spacer(1, 0.2*cm))
    
    # Bottom Text
    # Layout: [User (Left)] [Arabic Refund (Center)] [Page (Right)]
    footer_data = [
        ["User: Ahmed Essam", process_text("الاسترداد النقدى خلال 48 ساعة من أداء الخدمة من 9ص الى 3م عدا الجمعة والعطلات"), "Page 1 of 1"]
    ]
    t_footer = Table(footer_data, colWidths=[5*cm, 10*cm, 4*cm])
    t_footer.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (0,0), 'LEFT'),   # User Left
        ('ALIGN', (1,0), (1,0), 'CENTER'), # Arabic Center
        ('ALIGN', (2,0), (2,0), 'RIGHT'),  # Page Right
    ]))
    elements.append(t_footer)

    try:
        doc.build(elements)
        print(f"Generated: {output_path}")
    except Exception as e:
        print(f"Error {output_path}: {e}")

# --- 6. Main Execution ---
def main():
    try:
        if not os.path.exists('BulkContractorDetailedInvoice.xlsx'):
            print("Error: Excel file not found.")
            return

        xls = pd.ExcelFile('BulkContractorDetailedInvoice.xlsx')
        
        # Cleanup old output
        output_dir = 'invoices'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        print(f"Found {len(xls.sheet_names)} sheets to process.")
        
        for sheet in xls.sheet_names:
            print(f"Processing {sheet}...")
            df = pd.read_excel(xls, sheet_name=sheet)
            
            # Use extract_metadata helper to get Visit No
            # We need to find header row first
            header_row_idx = None
            for i, row in df.iterrows():
                vals = [str(x).lower().strip() for x in row.values if pd.notna(x)]
                if 'description' in vals and ('qty' in vals or 'unit' in vals):
                    header_row_idx = i
                    break
            
            metadata = {}
            if header_row_idx is not None:
                metadata = extract_metadata(df, header_row_idx)
            
            # Filename Logic: Use Visit No if available, else Invoice No, else Unknown
            visit_no = metadata.get('Visit No')
            inv_num = metadata.get('Invoice No')
            
            if visit_no and visit_no != '-':
                filename_base = f"Visit_{visit_no}"
            elif inv_num and inv_num != '-':
                filename_base = f"Invoice_{inv_num}"
            else:
                filename_base = f"Unknown_Sheet_{sheet}"
                
            # Clean filename
            filename_base = filename_base.replace('/', '-').replace('\\', '-')
            
            output_name = os.path.join(output_dir, f"{filename_base}.pdf")
            create_invoice_pdf(df, sheet, output_name)
            
        print("\nAll Invoices generated in 'invoices/' folder.")
            
    except Exception as e:
        print(f"Critical Error: {e}")

if __name__ == "__main__":
    main()
