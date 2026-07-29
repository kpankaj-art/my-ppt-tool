import io
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from pptx import Presentation
import streamlit as st

# Page Setup
st.set_page_config(page_title="Excel & PPT Smart Matcher", page_icon="📊", layout="centered")

# --- HIDE STREAMLIT TOP BAR / FORK / GITHUB LOGO ---
st.markdown("""
    <style>
    /* Top bar ko poori tarah hide karne ke liye */
    [data-testid="stHeader"] {
        display: none !important;
    }
    /* Main menu aur toolbar components hide karne ke liye */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    div[class*="stAppHeader"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# --- FILE PROCESSING FUNCTIONS ---

def extract_text_from_slide(slide):
    text_runs = []
    def process_shape(shape):
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text_runs.append(paragraph.text)
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    text_runs.append(cell.text)
        elif shape.shape_type == 6:  # Group Shape
            for sub_shape in shape.shapes:
                process_shape(sub_shape)

    for shape in slide.shapes:
        process_shape(shape)
    return " ".join(text_runs)

def process_files(pptx_bytes, excel_bytes):
    df = pd.read_excel(excel_bytes, dtype=str)
    
    col_mobile = next((c for c in df.columns if 'mobile' in str(c).lower()), None)
    col_width = next((c for c in df.columns if 'width' in str(c).lower()), None)
    col_height = next((c for c in df.columns if 'height' in str(c).lower()), None)

    if not col_mobile:
        raise ValueError("Excel file me 'MobileNumber' column nahi mila!")

    excel_criteria = []
    for idx, row in df.iterrows():
        mob = str(row[col_mobile]).split('.')[0].strip() if pd.notna(row[col_mobile]) else ""
        w = str(row[col_width]).split('.')[0].strip() if col_width and pd.notna(row[col_width]) else ""
        h = str(row[col_height]).split('.')[0].strip() if col_height and pd.notna(row[col_height]) else ""
        
        if mob and mob.lower() != 'nan':
            excel_criteria.append({'row_idx': idx, 'mobile': mob, 'width': w, 'height': h})

    prs = Presentation(pptx_bytes)
    slides = list(prs.slides)
    slide_texts = [extract_text_from_slide(slide).lower() for slide in slides]

    matched_indices = []
    seen_slides = set()
    matched_excel_rows = set()

    for item in excel_criteria:
        row_idx = item['row_idx']
        mob = item['mobile'].lower()
        w = item['width'].lower()
        h = item['height'].lower()

        best_match_idx = None
        for idx, text in enumerate(slide_texts):
            if idx in seen_slides:
                continue
            if mob in text:
                if w and h:
                    if w in text and h in text:
                        best_match_idx = idx
                        break
                    elif not best_match_idx:
                        best_match_idx = idx
                else:
                    best_match_idx = idx
                    break

        if best_match_idx is not None:
            matched_indices.append(best_match_idx)
            seen_slides.add(best_match_idx)
            matched_excel_rows.add(row_idx)

    # Reorder slides
    sldIdLst = prs.slides._sldIdLst
    original_sldIds = list(sldIdLst)
    for sldId in original_sldIds:
        sldIdLst.remove(sldId)

    for idx in matched_indices:
        sldIdLst.append(original_sldIds[idx])

    for idx, sldId in enumerate(original_sldIds):
        if idx not in matched_indices:
            sldIdLst.append(sldId)

    out_pptx_io = io.BytesIO()
    prs.save(out_pptx_io)
    out_pptx_io.seek(0)

    # Highlight Excel
    wb = openpyxl.load_workbook(excel_bytes)
    ws = wb.active
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    unmatched_count = 0
    for idx, row in enumerate(df.iterrows()):
        excel_row_num = idx + 2
        if idx not in matched_excel_rows:
            unmatched_count += 1
            for col_num in range(1, ws.max_column + 1):
                ws.cell(row=excel_row_num, column=col_num).fill = red_fill

    out_excel_io = io.BytesIO()
    wb.save(out_excel_io)
    out_excel_io.seek(0)

    return out_pptx_io.getvalue(), out_excel_io.getvalue(), len(matched_indices), unmatched_count

# --- MAIN DASHBOARD UI ---
st.title("📊 Excel & PPT Smart Matcher")
st.write("Apni PPT aur Excel file upload karke match aur reorder karein.")

uploaded_pptx = st.file_uploader("1. Select PowerPoint File (.pptx)", type=["pptx"])
uploaded_excel = st.file_uploader("2. Select Excel File (.xlsx)", type=["xlsx", "xls"])

if st.button("🚀 Process & Reorder", type="primary"):
    if uploaded_pptx and uploaded_excel:
        with st.spinner("Processing Files... Kripya wait karein..."):
            try:
                out_pptx_bytes, out_excel_bytes, matched_cnt, missing_cnt = process_files(uploaded_pptx, uploaded_excel)
                
                st.session_state["out_pptx"] = out_pptx_bytes
                st.session_state["out_excel"] = out_excel_bytes
                st.session_state["matched_cnt"] = matched_cnt
                st.session_state["missing_cnt"] = missing_cnt
                st.session_state["processed"] = True
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
    else:
        st.warning("Pehle dono files upload karein!")

if st.session_state.get("processed", False):
    st.success(f"Complete! Matched Slides: {st.session_state['matched_cnt']} | Missing Excel Rows: {st.session_state['missing_cnt']}")
    
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
            label="📥 Download Highlighted Excel",
            data=st.session_state["out_excel"],
            file_name="Missing_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_excel"
        )
