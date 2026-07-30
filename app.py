import io
import re
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
    [data-testid="stHeader"] {
        display: none !important;
    }
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    div[class*="stAppHeader"] {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---

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

def extract_details_from_text(text):
    """Extra slide me se mobile, sap code, width, height, aur name dhoondne ke liye"""
    # Find 10 digit mobile number
    mob_match = re.search(r'\b\d{10}\b', text)
    mobile = mob_match.group(0) if mob_match else ""

    # Find SAP Code (6 to 10 digit numbers)
    sap_match = re.search(r'\b\d{6,10}\b', text)
    sap = sap_match.group(0) if sap_match and sap_match.group(0) != mobile else ""

    # Find Width x Height (e.g. 10x20 or 10.5x20)
    dim_match = re.search(r'(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)', text, re.IGNORECASE)
    w = dim_match.group(1) if dim_match else ""
    h = dim_match.group(2) if dim_match else ""

    # Name cleanup
    name = text.split("\n")[0][:30].strip() if text else "Extra Slide Item"
    return name, mobile, sap, w, h

def process_files(pptx_bytes, excel_bytes):
    df = pd.read_excel(excel_bytes, dtype=str)
    
    # Identify Column Names Dynamically
    col_mobile = next((c for c in df.columns if 'mobile' in str(c).lower()), None)
    col_sap = next((c for c in df.columns if 'sap' in str(c).lower() or 'code' in str(c).lower()), None)
    col_name = next((c for c in df.columns if 'name' in str(c).lower() or 'party' in str(c).lower() or 'customer' in str(c).lower()), None)
    col_width = next((c for c in df.columns if 'width' in str(c).lower()), None)
    col_height = next((c for c in df.columns if 'height' in str(c).lower()), None)

    excel_criteria = []
    for idx, row in df.iterrows():
        mob = str(row[col_mobile]).split('.')[0].strip() if col_mobile and pd.notna(row[col_mobile]) else ""
        sap = str(row[col_sap]).split('.')[0].strip() if col_sap and pd.notna(row[col_sap]) else ""
        name = str(row[col_name]).strip() if col_name and pd.notna(row[col_name]) else ""
        w = str(row[col_width]).split('.')[0].strip() if col_width and pd.notna(row[col_width]) else ""
        h = str(row[col_height]).split('.')[0].strip() if col_height and pd.notna(row[col_height]) else ""
        
        excel_criteria.append({
            'row_idx': idx, 
            'mobile': mob, 
            'sap': sap, 
            'name': name, 
            'width': w, 
            'height': h
        })

    prs = Presentation(pptx_bytes)
    slides = list(prs.slides)
    raw_slide_texts = [extract_text_from_slide(slide) for slide in slides]
    slide_texts_lower = [t.lower() for t in raw_slide_texts]

    matched_indices = []
    seen_slides = set()
    matched_excel_rows = set()

    # Step 1: Matching Logic (Mobile -> SAP Code -> Name)
    for item in excel_criteria:
        row_idx = item['row_idx']
        mob = item['mobile'].lower()
        sap = item['sap'].lower()
        name = item['name'].lower()
        w = item['width'].lower()
        h = item['height'].lower()

        best_match_idx = None

        for idx, text in enumerate(slide_texts_lower):
            if idx in seen_slides:
                continue

            # 1st Preference: Mobile Number Match
            if mob and mob != 'nan' and mob in text:
                if w and h and (w in text and h in text):
                    best_match_idx = idx
                    break
                elif not best_match_idx:
                    best_match_idx = idx

            # 2nd Preference: SAP Code Match
            elif sap and sap != 'nan' and sap in text:
                if w and h and (w in text and h in text):
                    best_match_idx = idx
                    break
                elif not best_match_idx:
                    best_match_idx = idx

            # 3rd Preference: Name Match
            elif name and name != 'nan' and len(name) > 2 and name in text:
                best_match_idx = idx
                break

        if best_match_idx is not None:
            matched_indices.append(best_match_idx)
            seen_slides.add(best_match_idx)
            matched_excel_rows.add(row_idx)

    # Step 2: Reorder slides in PPT
    # Matched slides upar aayenge, remaining extra slides niche automatic chali jayengi
    sldIdLst = prs.slides._sldIdLst
    original_sldIds = list(sldIdLst)
    for sldId in original_sldIds:
        sldIdLst.remove(sldId)

    # Add matched slides first
    for idx in matched_indices:
        sldIdLst.append(original_sldIds[idx])

    # Unmatched / Extra slides in PPT (added at bottom)
    unmatched_slide_indices = []
    for idx, sldId in enumerate(original_sldIds):
        if idx not in matched_indices:
            sldIdLst.append(sldId)
            unmatched_slide_indices.append(idx)

    out_pptx_io = io.BytesIO()
    prs.save(out_pptx_io)
    out_pptx_io.seek(0)

    # Step 3: Excel Modification (Red for Missing, Green for Extra PPT Slides)
    wb = openpyxl.load_workbook(excel_bytes)
    ws = wb.active

    # Check or create Remark Column
    headers = [cell.value for cell in ws[1]]
    remark_col_num = len(headers) + 1
    ws.cell(row=1, column=remark_col_num, value="Remark")

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    green_fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")

    # Highlight Missing Excel Rows in Red
    missing_count = 0
    for idx, row in enumerate(df.iterrows()):
        excel_row_num = idx + 2
        if idx not in matched_excel_rows:
            missing_count += 1
            for col_num in range(1, remark_col_num + 1):
                ws.cell(row=excel_row_num, column=col_num).fill = red_fill
            ws.cell(row=excel_row_num, column=remark_col_num, value="Not Found in PPT")

    # Append Extra PPT Slides at Bottom of Excel with Green Highlight
    extra_count = 0
    for s_idx in unmatched_slide_indices:
        extra_count += 1
        slide_txt = raw_slide_texts[s_idx]
        ext_name, ext_mob, ext_sap, ext_w, ext_h = extract_details_from_text(slide_txt)

        new_row_idx = ws.max_row + 1

        # Populate Extracted Data into correct columns if available
        if col_name:
            c_idx = df.columns.get_loc(col_name) + 1
            ws.cell(row=new_row_idx, column=c_idx, value=ext_name)
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

        # Color entire new row green
        for col_num in range(1, remark_col_num + 1):
            ws.cell(row=new_row_idx, column=col_num).fill = green_fill

    out_excel_io = io.BytesIO()
    wb.save(out_excel_io)
    out_excel_io.seek(0)

    return out_pptx_io.getvalue(), out_excel_io.getvalue(), len(matched_indices), missing_count, extra_count

# --- MAIN DASHBOARD UI ---
st.title("📊 Excel & PPT Smart Matcher")
st.write("Apni PPT aur Excel file upload karke match aur reorder karein.")

uploaded_pptx = st.file_uploader("1. Select PowerPoint File (.pptx)", type=["pptx"])
uploaded_excel = st.file_uploader("2. Select Excel File (.xlsx)", type=["xlsx", "xls"])

if st.button("🚀 Process & Reorder", type="primary"):
    if uploaded_pptx and uploaded_excel:
        with st.spinner("Processing Files... Kripya wait karein..."):
            try:
                out_pptx_bytes, out_excel_bytes, matched_cnt, missing_cnt, extra_cnt = process_files(uploaded_pptx, uploaded_excel)
                
                st.session_state["out_pptx"] = out_pptx_bytes
                st.session_state["out_excel"] = out_excel_bytes
                st.session_state["matched_cnt"] = matched_cnt
                st.session_state["missing_cnt"] = missing_cnt
                st.session_state["extra_cnt"] = extra_cnt
                st.session_state["processed"] = True
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
    else:
        st.warning("Pehle dono files upload karein!")

if st.session_state.get("processed", False):
    st.success(f"Complete! Matched: {st.session_state['matched_cnt']} | Missing in PPT (Red): {st.session_state['missing_cnt']} | Extra in PPT (Green): {st.session_state['extra_cnt']}")
    
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
