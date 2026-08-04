import io
import re
import time
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from pptx import Presentation
import streamlit as st

# Page Setup
st.set_page_config(page_title="Pro PPT & Excel Matcher", page_icon="⚡", layout="centered")

# --- PROFESSIONAL UI & CUSTOM CSS ---
st.markdown("""
    <style>
    [data-testid="stHeader"] { display: none !important; }
    #MainMenu { visibility: hidden; }
    header { visibility: hidden; }
    footer { visibility: hidden; }
    div[class*="stAppHeader"] { display: none !important; }

    .main { background-color: #0e1117; }
    
    .css-card {
        background: linear-gradient(135deg, #1e2638 0%, #111827 100%);
        border: 1px solid #374151;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }
    
    .title-text {
        font-family: 'Inter', sans-serif;
        font-size: 28px;
        font-weight: 700;
        color: #f9fafb;
        margin-bottom: 8px;
    }
    .subtitle-text {
        font-size: 14px;
        color: #9ca3af;
        margin-bottom: 24px;
    }
    
    .metric-container {
        display: flex;
        justify-content: space-between;
        gap: 12px;
        margin-top: 15px;
        margin-bottom: 15px;
    }
    .metric-box {
        background: #1f2937;
        border-radius: 8px;
        padding: 12px 16px;
        text-align: center;
        flex: 1;
        border: 1px solid #374151;
    }
    .metric-value {
        font-size: 20px;
        font-weight: 700;
    }
    .metric-label {
        font-size: 12px;
        color: #9ca3af;
        margin-top: 4px;
    }
    
    .stDownloadButton > button {
        width: 100%;
        border-radius: 8px;
        font-weight: 600;
        height: 48px;
    }
    </style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---

def extract_text_from_slide(slide):
    text_runs = []
    def process_shape(shape):
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if paragraph.text:
                    text_runs.append(paragraph.text.strip())
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text:
                        text_runs.append(cell.text.strip())
        elif shape.shape_type == 6:  # Group Shape
            for sub_shape in shape.shapes:
                process_shape(sub_shape)

    for shape in slide.shapes:
        process_shape(shape)
    return "\n".join(text_runs)

def update_slide_size(slide, new_size_str):
    """Excel ki Nayi Size ko PPT Slide me Replace karne ka Logic"""
    if not new_size_str:
        return

    def replace_in_text_frame(tf):
        for paragraph in tf.paragraphs:
            p_text = paragraph.text
            if any(k in p_text.lower() for k in ["size", "qty", "type", "width", "height", "dimension"]):
                if re.search(r'Size\s*[:\-]\s*[^\n\r]+', p_text, re.IGNORECASE):
                    paragraph.text = re.sub(r'(Size\s*[:\-]\s*)[^\n\r]+', rf'\1{new_size_str}', p_text, flags=re.IGNORECASE)
                elif re.search(r'\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:ft|f|in)?\b', p_text, re.IGNORECASE):
                    paragraph.text = re.sub(r'\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:ft|f|in)?\b', new_size_str, p_text, flags=re.IGNORECASE)

    def process_shape_update(shape):
        if shape.has_text_frame:
            replace_in_text_frame(shape.text_frame)
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame:
                        replace_in_text_frame(cell.text_frame)
        elif shape.shape_type == 6:  # Group Shape
            for sub_shape in shape.shapes:
                process_shape_update(sub_shape)

    for shape in slide.shapes:
        process_shape_update(shape)

def extract_details_from_text(text):
    """PPT ke Text me se Details Extractions"""
    name = ""
    name_match = re.search(r'(?:Outlet\s*Name|Party\s*Name|Customer\s*Name|Dealer\s*Name|Shop\s*Name|Store\s*Name|Name)\s*[:\-]\s*([^\n\r]+)', text, re.IGNORECASE)
    
    if name_match:
        name = name_match.group(1).strip()
        name = re.split(r'(?:Address|Contact|City|Date|Qty|Size|Type|Width|Height)', name, flags=re.IGNORECASE)[0].strip()
    else:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        for l in lines:
            if not any(k in l.lower() for k in ['qty', 'size', 'type', 'contact', 'address', 'city', 'date', 'width', 'height']):
                name = l[:40].strip()
                break

    mob_match = re.search(r'\b[6-9]\d{9}\b', text)
    mobile = mob_match.group(0) if mob_match else ""

    sap_match = re.search(r'\b\d{6,10}\b', text)
    sap = sap_match.group(0) if sap_match and sap_match.group(0) != mobile else ""

    dim_match = re.search(r'(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    w = dim_match.group(1) if dim_match else ""
    h = dim_match.group(2) if dim_match else ""

    return name, mobile, sap, w, h

def process_files_with_progress(pptx_bytes, excel_bytes, progress_bar, status_text):
    status_text.markdown("⏳ **Reading Excel Data...**")
    progress_bar.progress(10)
    df = pd.read_excel(excel_bytes, dtype=str)
    
    # Column Identification
    col_mobile = next((c for c in df.columns if any(k in str(c).lower() for k in ['mobile', 'contact', 'phone', 'num'])), None)
    col_sap = next((c for c in df.columns if any(k in str(c).lower() for k in ['sap', 'code', 'dealer_code', 'id'])), None)
    col_name = next((c for c in df.columns if any(k in str(c).lower() for k in ['outlet', 'party', 'customer', 'dealer', 'shop', 'store', 'name'])), None)
    col_city = next((c for c in df.columns if any(k in str(c).lower() for k in ['city', 'location', 'address', 'town', 'place'])), None)
    col_width = next((c for c in df.columns if 'width' in str(c).lower() or 'w' == str(c).strip().lower()), None)
    col_height = next((c for c in df.columns if 'height' in str(c).lower() or 'h' == str(c).strip().lower()), None)
    col_size = next((c for c in df.columns if 'size' in str(c).lower() or 'dimension' in str(c).lower()), None)

    excel_criteria = []
    for idx, row in df.iterrows():
        mob = str(row[col_mobile]).split('.')[0].strip() if col_mobile and pd.notna(row[col_mobile]) else ""
        sap = str(row[col_sap]).split('.')[0].strip() if col_sap and pd.notna(row[col_sap]) else ""
        name = str(row[col_name]).strip() if col_name and pd.notna(row[col_name]) else ""
        city = str(row[col_city]).strip() if col_city and pd.notna(row[col_city]) else ""
        w = str(row[col_width]).split('.')[0].strip() if col_width and pd.notna(row[col_width]) else ""
        h = str(row[col_height]).split('.')[0].strip() if col_height and pd.notna(row[col_height]) else ""
        
        size_str = ""
        if col_size and pd.notna(row[col_size]):
            size_str = str(row[col_size]).strip()
        elif w or h:
            size_str = f"W-{w} H-{h}" if w and h else (w or h)

        excel_criteria.append({
            'row_idx': idx, 'mobile': mob, 'sap': sap, 'name': name, 'city': city, 'width': w, 'height': h, 'size_str': size_str
        })

    status_text.markdown("⏳ **Analyzing Presentation Slides...**")
    progress_bar.progress(25)
    prs = Presentation(pptx_bytes)
    slides = list(prs.slides)
    total_slides = len(slides)
    
    raw_slide_texts = []
    for idx, slide in enumerate(slides):
        raw_slide_texts.append(extract_text_from_slide(slide))
        curr_p = 25 + int((idx + 1) / total_slides * 20)
        progress_bar.progress(min(curr_p, 45))
        
    slide_texts_lower = [t.lower() for t in raw_slide_texts]

    status_text.markdown("⏳ **Smart Matching & Updating PPT Sizes...**")
    matched_indices = []
    seen_slides = set()
    matched_excel_rows = set()

    total_items = len(excel_criteria)
    for i, item in enumerate(excel_criteria):
        row_idx = item['row_idx']
        mob, sap = item['mobile'].lower(), item['sap'].lower()
        name, city = item['name'].lower(), item['city'].lower()
        w, h = item['width'].lower(), item['height'].lower()
        size_str = item['size_str']

        best_match_idx = None

        # Pass 1: Unique Mobile Match
        if mob and mob != 'nan':
            for idx, text in enumerate(slide_texts_lower):
                if idx not in seen_slides and mob in text:
                    best_match_idx = idx
                    break

        # Pass 2: Unique SAP Match
        if best_match_idx is None and sap and sap != 'nan':
            for idx, text in enumerate(slide_texts_lower):
                if idx not in seen_slides and sap in text:
                    best_match_idx = idx
                    break

        # Pass 3: Name + City/Address Match
        if best_match_idx is None and name and name != 'nan' and len(name) > 2:
            if city and city != 'nan':
                for idx, text in enumerate(slide_texts_lower):
                    if idx not in seen_slides and name in text and city in text:
                        best_match_idx = idx
                        break

        # Pass 4: Name + Dimensions Match
        if best_match_idx is None and name and name != 'nan' and len(name) > 2:
            if w and h:
                for idx, text in enumerate(slide_texts_lower):
                    if idx not in seen_slides and name in text and w in text and h in text:
                        best_match_idx = idx
                        break

        # Pass 5: Fallback Name Match
        if best_match_idx is None and name and name != 'nan' and len(name) > 2:
            for idx, text in enumerate(slide_texts_lower):
                if idx not in seen_slides and name in text:
                    best_match_idx = idx
                    break

        if best_match_idx is not None:
            matched_indices.append(best_match_idx)
            seen_slides.add(best_match_idx)
            matched_excel_rows.add(row_idx)

            # Auto-update size in PPT slide from Excel
            if size_str:
                update_slide_size(slides[best_match_idx], size_str)

        curr_p = 45 + int((i + 1) / total_items * 30)
        progress_bar.progress(min(curr_p, 75))

    status_text.markdown("⏳ **Reordering PPT Slides...**")
    progress_bar.progress(80)
    
    sldIdLst = prs.slides._sldIdLst
    original_sldIds = list(sldIdLst)
    for sldId in original_sldIds:
        sldIdLst.remove(sldId)

    for idx in matched_indices:
        sldIdLst.append(original_sldIds[idx])

    unmatched_slide_indices = []
    for idx, sldId in enumerate(original_sldIds):
        if idx not in matched_indices:
            sldIdLst.append(sldId)
            unmatched_slide_indices.append(idx)

    out_pptx_io = io.BytesIO()
    prs.save(out_pptx_io)
    out_pptx_io.seek(0)

    status_text.markdown("⏳ **Generating Final Excel Report...**")
    progress_bar.progress(90)
    
    wb = openpyxl.load_workbook(excel_bytes)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    remark_col_num = len(headers) + 1
    ws.cell(row=1, column=remark_col_num, value="Remark")

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    green_fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")

    missing_count = 0
    for idx, row in enumerate(df.iterrows()):
        excel_row_num = idx + 2
        if idx not in matched_excel_rows:
            missing_count += 1
            for col_num in range(1, remark_col_num + 1):
                ws.cell(row=excel_row_num, column=col_num).fill = red_fill
            ws.cell(row=excel_row_num, column=remark_col_num, value="Not Found in PPT")

    extra_count = 0
    for s_idx in unmatched_slide_indices:
        extra_count += 1
        slide_txt = raw_slide_texts[s_idx]
        ext_name, ext_mob, ext_sap, ext_w, ext_h = extract_details_from_text(slide_txt)

        new_row_idx = ws.max_row + 1

        if col_name:
            c_idx = df.columns.get_loc(col_name) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_name)
        elif len(headers) >= 1:
            ws.cell(row=new_row_idx, column=1, value=ext_name)

        if col_mobile:
            c_idx = df.columns.get_loc(col_mobile) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_mob)

        if col_sap:
            c_idx = df.columns.get_loc(col_sap) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_sap)

        if col_width:
            c_idx = df.columns.get_loc(col_width) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_w)

        if col_height:
            c_idx = df.columns.get_loc(col_height) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_h)

        ws.cell(row=new_row_idx, column=remark_col_num, value="ye ppt me extra hai")

        for col_num in range(1, remark_col_num + 1):
            ws.cell(row=new_row_idx, column=col_num).fill = green_fill

    out_excel_io = io.BytesIO()
    wb.save(out_excel_io)
    out_excel_io.seek(0)

    progress_bar.progress(100)
    status_text.markdown("✅ **Processing Complete!**")
    time.sleep(0.5)

    return out_pptx_io.getvalue(), out_excel_io.getvalue(), len(matched_indices), missing_count, extra_count

# --- MAIN INTERFACE ---

st.markdown("""
    <div class="css-card">
        <div class="title-text">⚡ Smart PPT & Excel Matcher Pro</div>
        <div class="subtitle-text">Match, reorder, and sync your presentation slides with Excel data seamlessly.</div>
    </div>
""", unsafe_allow_html=True)

uploaded_pptx = st.file_uploader("1. Upload PowerPoint Presentation (.pptx)", type=["pptx"])
uploaded_excel = st.file_uploader("2. Upload Master Excel File (.xlsx)", type=["xlsx", "xls"])

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 Process & Sync Files", type="primary", use_container_width=True):
    if uploaded_pptx and uploaded_excel:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            out_pptx_bytes, out_excel_bytes, matched_cnt, missing_cnt, extra_cnt = process_files_with_progress(
                uploaded_pptx, uploaded_excel, progress_bar, status_text
            )
            
            st.session_state["out_pptx"] = out_pptx_bytes
            st.session_state["out_excel"] = out_excel_bytes
            st.session_state["matched_cnt"] = matched_cnt
            st.session_state["missing_cnt"] = missing_cnt
            st.session_state["extra_cnt"] = extra_cnt
            st.session_state["processed"] = True
            
            status_text.empty()
            progress_bar.empty()
            
        except Exception as e:
            status_text.empty()
            progress_bar.empty()
            st.error(f"❌ Error during processing: {str(e)}")
    else:
        st.warning("⚠️ Kripya dono PowerPoint aur Excel files upload karein!")

# --- DISPLAY RESULTS & DOWNLOADS ---
if st.session_state.get("processed", False):
    st.markdown(f"""
        <div class="metric-container">
            <div class="metric-box">
                <div class="metric-value" style="color: #34d399;">{st.session_state['matched_cnt']}</div>
                <div class="metric-label">Matched Slides</div>
            </div>
            <div class="metric-box">
                <div class="metric-value" style="color: #f87171;">{st.session_state['missing_cnt']}</div>
                <div class="metric-label">Missing (Red)</div>
            </div>
            <div class="metric-box">
                <div class="metric-value" style="color: #60a5fa;">{st.session_state['extra_cnt']}</div>
                <div class="metric-label">Extra PPT (Green)</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📥 Download Sorted PPT",
            data=st.session_state["out_pptx"],
            file_name="Sorted_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="btn_ppt"
        )
    with col2:
        st.download_button(
            label="📥 Download Updated Excel",
            data=st.session_state["out_excel"],
            file_name="Missing_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_excel"
        )
