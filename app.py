import io
import re
import difflib
import math
import time
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from pptx import Presentation
import streamlit as st

# ============================================================
# SMART PPT & EXCEL MATCHER PRO V2
# ============================================================

st.set_page_config(
    page_title="Smart PPT & Excel Matcher Pro",
    page_icon="⚡",
    layout="wide"
)

# -------------------- UI --------------------
st.markdown("""
<style>
.stApp { background-color: #0e1117; }
.stButton > button {
    border-radius: 10px;
    font-weight: 800;
    min-height: 48px;
}
.css-card {
    background: linear-gradient(135deg, #1e2638 0%, #111827 100%);
    border: 1px solid #374151;
    border-radius: 14px;
    padding: 24px;
    margin-bottom: 20px;
}
.title-text { font-size: 30px; font-weight: 800; color: #f9fafb; }
.subtitle-text { font-size: 14px; color: #9ca3af; margin-top: 6px; }
[data-testid="stHeader"] { display: none; }
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# -------------------- Constants --------------------
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F2937")

RED_FONT = Font(color="9C0006")
GREEN_FONT = Font(color="006100")
YELLOW_FONT = Font(color="9C6500")
WHITE_BOLD = Font(color="FFFFFF", bold=True)

REPORT_COLUMNS = [
    "Match Status",
    "Matched PPT Slide",
    "Confidence %",
    "Match Reason",
    "Missing Fields",
    "PPT Extra",
    "Review Required",
    "Remark"
]

# ============================================================
# NORMALIZATION
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value).strip().lower()
    if s in {"nan", "none", "null", "nat"}:
        return ""
    return s


def normalize_text(value):
    """General comparison normalization."""
    s = clean_text(value)
    if not s:
        return ""

    # Common symbols -> spaces
    s = s.replace("&", " and ")
    s = s.replace("×", "x")
    s = s.replace("*", "x")

    # Remove punctuation/whitespace
    s = re.sub(r"[^a-z0-9]+", "", s)

    # Remove trailing .0
    if s.endswith("0") and re.search(r"\d+0$", s):
        pass

    return s


def normalize_phone(value):
    s = clean_text(value)
    digits = re.sub(r"\D", "", s)

    # Handle common Indian formats
    if len(digits) == 12 and digits.startswith("91"):
        digits = digits[2:]
    elif len(digits) == 11 and digits.startswith("0"):
        digits = digits[1:]

    return digits[-10:] if len(digits) >= 10 else digits


def normalize_code(value):
    s = clean_text(value)
    # Excel may turn numeric codes into 12345.0
    s = re.sub(r"\.0+$", "", s)
    return re.sub(r"[^a-z0-9]", "", s)


def normalize_number(value):
    s = clean_text(value).replace(",", "")
    s = s.replace("×", "x").replace("*", "x")
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return ""
    try:
        n = float(m.group())
        return str(int(n)) if n.is_integer() else str(n)
    except Exception:
        return m.group()


def normalize_dimension(value):
    s = clean_text(value)
    s = s.replace("×", "x").replace("*", "x")
    s = re.sub(r"\b(ft|feet|foot|in|inch|inches)\b", "", s)
    m = re.search(r"\d+(?:\.\d+)?", s)
    if not m:
        return ""
    try:
        n = float(m.group())
        return str(int(n)) if n.is_integer() else str(n)
    except Exception:
        return m.group()


def normalize_size_pair(w, h):
    nw = normalize_dimension(w)
    nh = normalize_dimension(h)
    if nw and nh:
        return f"{nw}x{nh}"
    return ""


# ============================================================
# EXCEL COLUMN DETECTION
# ============================================================

def find_column(columns, keywords, exclude=None):
    exclude = exclude or []
    scored = []

    for col in columns:
        c = clean_text(col)
        if any(x in c for x in exclude):
            continue

        score = 0
        for kw in keywords:
            if c == kw:
                score += 100
            elif kw in c:
                score += 30

        if score:
            scored.append((score, col))

    if not scored:
        return None

    scored.sort(reverse=True, key=lambda x: x[0])
    return scored[0][1]


def detect_excel_columns(df):
    cols = list(df.columns)

    return {
        "name": find_column(
            cols,
            ["outlet name", "dealer name", "party name", "customer name",
             "shop name", "store name", "outlet", "dealer", "party",
             "customer", "shop", "store", "name"]
        ),
        "mobile": find_column(
            cols,
            ["mobile no", "mobile number", "contact no", "contact number",
             "phone no", "phone number", "mobile", "contact", "phone", "mob"]
        ),
        "sap": find_column(
            cols,
            ["sap code", "dealer code", "dealer id", "sap", "code", "dealer_id", "id"]
        ),
        "address": find_column(
            cols,
            ["full address", "address", "addr"]
        ),
        "city": find_column(
            cols,
            ["district", "city", "location", "town", "place", "area"]
        ),
        "width": find_column(
            cols,
            ["width", "width ft", "width(ft)", "w(ft)", "w"]
        ),
        "height": find_column(
            cols,
            ["height", "height ft", "height(ft)", "h(ft)", "h"]
        ),
        "size": find_column(
            cols,
            ["size", "dimension", "dimensions", "size ft"]
        ),
        "quantity": find_column(
            cols,
            ["quantity", "qty", "qnty", "units", "no of", "number of"]
        ),
        "media_type": find_column(
            cols,
            ["media type", "media", "type", "display type", "sign type"]
        )
    }


# ============================================================
# FILE READERS
# ============================================================

def load_excel(uploaded_file):
    filename = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if filename.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data), dtype=str).fillna("")

    if filename.endswith(".xls"):
        return pd.read_excel(io.BytesIO(data), dtype=str, engine="xlrd").fillna("")

    return pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl").fillna("")


def dataframe_to_workbook(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sync Report"

    for col_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(1, col_idx, str(col))
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row_idx, col_idx, value)

    return wb


# ============================================================
# PPT EXTRACTION
# ============================================================

def extract_text_from_shape(shape, output):
    try:
        if getattr(shape, "has_text_frame", False):
            txt = shape.text_frame.text
            if txt and txt.strip():
                output.append(txt.strip())

        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                vals = []
                for cell in row.cells:
                    vals.append(cell.text.strip())
                if any(vals):
                    output.append(" | ".join(vals))

        # Group shape
        if hasattr(shape, "shapes"):
            for sub in shape.shapes:
                extract_text_from_shape(sub, output)
    except Exception:
        pass


def extract_text_from_slide(slide):
    parts = []
    for shape in slide.shapes:
        extract_text_from_shape(shape, parts)
    return "\n".join(parts)


def clean_lines(text):
    return [x.strip() for x in text.splitlines() if x.strip()]


def value_after_label(text, labels):
    """
    Finds values in formats:
    Label: Value
    Label - Value
    Label Value
    """
    escaped = "|".join(re.escape(x) for x in labels)

    patterns = [
        rf"(?:{escaped})\s*[:\-]\s*([^\n\r|]+)",
        rf"(?:{escaped})\s+([^\n\r|]+)"
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            value = m.group(1).strip()
            value = re.split(
                r"\b(?:address|mobile|contact|phone|sap|dealer|city|district|"
                r"qty|quantity|size|width|height|type|media)\b",
                value,
                flags=re.IGNORECASE
            )[0].strip(" :-|")
            if value:
                return value
    return ""


def extract_slide_details(text):
    details = {
        "name": "",
        "mobile": "",
        "sap": "",
        "address": "",
        "city": "",
        "width": "",
        "height": "",
        "size": "",
        "quantity": "",
        "media_type": ""
    }

    details["name"] = value_after_label(text, [
        "Outlet Name", "Dealer Name", "Party Name", "Customer Name",
        "Shop Name", "Store Name", "Outlet", "Dealer", "Party",
        "Customer", "Shop", "Store", "Name"
    ])

    details["mobile"] = normalize_phone(
        value_after_label(text, [
            "Mobile No", "Mobile Number", "Mobile", "Contact No",
            "Contact Number", "Contact", "Phone No", "Phone Number", "Phone"
        ])
    )

    details["sap"] = normalize_code(
        value_after_label(text, [
            "SAP Code", "SAP", "Dealer Code", "Dealer ID", "Dealer Id", "Code"
        ])
    )

    details["address"] = value_after_label(text, [
        "Full Address", "Address", "Addr"
    ])

    details["city"] = value_after_label(text, [
        "District", "City", "Location", "Town", "Place", "Area"
    ])

    details["quantity"] = normalize_number(
        value_after_label(text, [
            "Quantity", "Qty", "Qnty", "Units", "No of", "Number of"
        ])
    )

    details["media_type"] = value_after_label(text, [
        "Media Type", "Media", "Display Type", "Sign Type", "Type"
    ])

    # Size: first look for W x H
    dim = re.search(
        r"(\d+(?:\.\d+)?)\s*(?:ft|feet)?\s*[x×*]\s*"
        r"(\d+(?:\.\d+)?)\s*(?:ft|feet)?",
        text,
        re.IGNORECASE
    )

    if dim:
        details["width"] = normalize_dimension(dim.group(1))
        details["height"] = normalize_dimension(dim.group(2))
        details["size"] = normalize_size_pair(
            details["width"], details["height"]
        )
    else:
        details["width"] = normalize_dimension(
            value_after_label(text, ["Width", "W", "Width(ft)", "W(ft)"])
        )
        details["height"] = normalize_dimension(
            value_after_label(text, ["Height", "H", "Height(ft)", "H(ft)"])
        )
        details["size"] = normalize_size_pair(
            details["width"], details["height"]
        )

    # Fallback: find a likely 10 digit Indian mobile anywhere
    if not details["mobile"]:
        m = re.search(r"(?<!\d)(?:\+?91[\s\-]?)?[6-9]\d{9}(?!\d)", text)
        if m:
            details["mobile"] = normalize_phone(m.group())

    # Fallback: name from first useful line
    if not details["name"]:
        bad = [
            "qty", "quantity", "size", "type", "mobile", "contact", "phone",
            "address", "city", "district", "width", "height", "sap", "code",
            "date", "media"
        ]
        for line in clean_lines(text):
            if len(line) >= 3 and not any(x in line.lower() for x in bad):
                if not re.fullmatch(r"[\d\s:/.,xX\-]+", line):
                    details["name"] = line[:100]
                    break

    # Generic SAP fallback: 6-12 digit standalone code, excluding mobile
    if not details["sap"]:
        for candidate in re.findall(r"(?<!\d)\d{6,12}(?!\d)", text):
            c = normalize_code(candidate)
            if c and c != details["mobile"]:
                details["sap"] = c
                break

    return details


# ============================================================
# MATCH ENGINE
# ============================================================

def token_similarity(a, b):
    a = normalize_text(a)
    b = normalize_text(b)

    if not a or not b:
        return 0.0

    if a == b:
        return 1.0

    if len(a) >= 4 and (a in b or b in a):
        return min(len(a), len(b)) / max(len(a), len(b))

    return difflib.SequenceMatcher(None, a, b).ratio()


def field_match(a, b, field):
    if field == "mobile":
        a = normalize_phone(a)
        b = normalize_phone(b)
        return 1.0 if a and b and a == b else 0.0

    if field == "sap":
        a = normalize_code(a)
        b = normalize_code(b)
        return 1.0 if a and b and a == b else 0.0

    if field in {"width", "height", "quantity"}:
        a = normalize_number(a)
        b = normalize_number(b)
        return 1.0 if a and b and a == b else 0.0

    if field == "size":
        return 1.0 if (
            normalize_text(a) and
            normalize_text(b) and
            normalize_text(a) == normalize_text(b)
        ) else token_similarity(a, b)

    return token_similarity(a, b)


FIELD_WEIGHTS = {
    "mobile": 40,
    "sap": 35,
    "name": 22,
    "address": 12,
    "city": 10,
    "size": 15,
    "width": 7,
    "height": 7,
    "quantity": 6,
    "media_type": 5
}


def build_excel_record(row, mapping, row_idx):
    rec = {"row_idx": row_idx}

    for field, col in mapping.items():
        rec[field] = clean_text(row[col]) if col else ""

    rec["mobile"] = normalize_phone(rec.get("mobile", ""))
    rec["sap"] = normalize_code(rec.get("sap", ""))
    rec["name_norm"] = normalize_text(rec.get("name", ""))
    rec["city_norm"] = normalize_text(rec.get("city", ""))
    rec["address_norm"] = normalize_text(rec.get("address", ""))

    if rec.get("size"):
        rec["size_norm"] = normalize_text(rec["size"])
    else:
        rec["size_norm"] = normalize_text(
            normalize_size_pair(rec.get("width", ""), rec.get("height", ""))
        )

    return rec


def candidate_score(excel_rec, ppt_rec):
    """
    Weighted score. Strong identifiers get priority.
    Conflicts reduce the score to prevent dangerous false matches.
    """
    score = 0.0
    max_possible = 0.0
    reasons = []
    conflicts = []

    # Strong identifiers
    if excel_rec["mobile"] and ppt_rec["mobile"]:
        max_possible += FIELD_WEIGHTS["mobile"]
        if excel_rec["mobile"] == ppt_rec["mobile"]:
            score += FIELD_WEIGHTS["mobile"]
            reasons.append("Mobile")
        else:
            conflicts.append("Mobile conflict")
            score -= 35

    if excel_rec["sap"] and ppt_rec["sap"]:
        max_possible += FIELD_WEIGHTS["sap"]
        if excel_rec["sap"] == ppt_rec["sap"]:
            score += FIELD_WEIGHTS["sap"]
            reasons.append("SAP")
        else:
            conflicts.append("SAP conflict")
            score -= 30

    # Other fields
    for field in ["name", "address", "city", "size", "width", "height",
                  "quantity", "media_type"]:
        ev = excel_rec.get(field, "")
        pv = ppt_rec.get(field, "")

        if field == "size":
            ev = excel_rec.get("size_norm", "")
            pv = normalize_text(ppt_rec.get("size", ""))

        if not ev or not pv:
            continue

        weight = FIELD_WEIGHTS[field]
        sim = field_match(ev, pv, field)

        # Don't let tiny similarities count
        if sim >= 0.88:
            score += weight
            max_possible += weight
            reasons.append(field.replace("_", " ").title())
        elif sim >= 0.65 and field in {"name", "address", "city"}:
            score += weight * sim
            max_possible += weight
            reasons.append(f"{field.title()} similar")

    # If no comparable data exists, candidate is unusable
    if max_possible <= 0:
        return 0.0, reasons, conflicts

    # Normalize against a useful baseline.
    confidence = max(0.0, min(100.0, (score / max_possible) * 100))

    # Bonus when multiple independent identifiers agree.
    if len(reasons) >= 3:
        confidence = min(100.0, confidence + 5)

    # Strong ID + size/name gives a very strong match.
    if "Mobile" in reasons and ("Size" in reasons or "Name" in reasons):
        confidence = min(100.0, confidence + 5)

    if "SAP" in reasons and ("Size" in reasons or "Name" in reasons):
        confidence = min(100.0, confidence + 5)

    # Penalize explicit conflicts.
    confidence = max(0.0, confidence + (sum(-8 for _ in conflicts)))

    return confidence, reasons, conflicts


def get_missing_fields(excel_rec, ppt_rec):
    missing = []

    for field, label in [
        ("name", "Name"),
        ("mobile", "Mobile"),
        ("sap", "SAP"),
        ("address", "Address"),
        ("city", "City/District"),
        ("size", "Size"),
        ("quantity", "Quantity"),
        ("media_type", "Media Type")
    ]:
        ev = excel_rec.get(field, "")
        if field == "size":
            ev = excel_rec.get("size_norm", "")

        pv = ppt_rec.get(field, "")
        if field == "size":
            pv = normalize_text(ppt_rec.get("size", ""))

        if ev and not pv:
            missing.append(label)

    return ", ".join(missing)


def match_all(excel_records, ppt_records):
    """
    Global greedy assignment:
    1. Calculate every candidate score.
    2. Sort strongest candidates first.
    3. Assign each Excel row and PPT slide only once.
    This prevents two Excel rows from consuming the same slide.
    """
    candidates = []

    for e_idx, e in enumerate(excel_records):
        for p_idx, p in enumerate(ppt_records):
            confidence, reasons, conflicts = candidate_score(e, p)
            if confidence > 0:
                candidates.append(
                    (confidence, e_idx, p_idx, reasons, conflicts)
                )

    candidates.sort(reverse=True, key=lambda x: x[0])

    assigned_excel = set()
    assigned_ppt = set()
    results = {}

    for confidence, e_idx, p_idx, reasons, conflicts in candidates:
        if e_idx in assigned_excel or p_idx in assigned_ppt:
            continue

        # Minimum confidence gate.
        # Strong IDs can pass with 70+, otherwise require 78+.
        strong_id = (
            "Mobile" in reasons or
            "SAP" in reasons
        )

        threshold = 70 if strong_id else 78

        if confidence < threshold:
            continue

        # Check whether the next-best candidate is too close.
        # If ambiguous, leave for review instead of forcing a match.
        nearby = [
            c for c in candidates
            if c[1] == e_idx and c[2] != p_idx and c[0] >= confidence - 7
        ]

        if nearby and confidence < 90 and not strong_id:
            continue

        assigned_excel.add(e_idx)
        assigned_ppt.add(p_idx)
        results[e_idx] = {
            "ppt_idx": p_idx,
            "confidence": round(confidence),
            "reasons": reasons,
            "conflicts": conflicts
        }

    return results, assigned_excel, assigned_ppt


# ============================================================
# EXCEL REPORT
# ============================================================

def autosize_worksheet(ws):
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)

        for cell in col_cells:
            try:
                value = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(value))
            except Exception:
                pass

        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def append_report_columns(ws):
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

    # Remove duplicate report columns if re-run logic is ever reused
    existing = {str(x).strip() for x in headers if x is not None}

    for col in REPORT_COLUMNS:
        if col not in existing:
            ws.cell(1, ws.max_column + 1, col)

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    col_map = {str(v): i for i, v in enumerate(headers, 1)}

    for c in range(1, ws.max_column + 1):
        ws.cell(1, c).fill = HEADER_FILL
        ws.cell(1, c).font = WHITE_BOLD
        ws.cell(1, c).alignment = Alignment(horizontal="center")

    return col_map


def fill_entire_row(ws, row_num, fill, font):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row_num, c)
        cell.fill = fill
        cell.font = font


def add_extra_ppt_row(ws, col_map, mapping, details, slide_number):
    row_num = ws.max_row + 1

    # Put extracted PPT data into matching existing Excel columns.
    for field, col in mapping.items():
        if not col:
            continue

        value = details.get(field, "")

        if field == "mobile":
            value = normalize_phone(value)
        elif field == "sap":
            value = normalize_code(value)
        elif field in {"width", "height", "quantity"}:
            value = normalize_number(value)

        # Size-only Excel columns
        if field == "size" and not value:
            value = details.get("size", "")

        excel_col = col_map.get(str(col))
        if excel_col:
            ws.cell(row_num, excel_col, value)

    # If no dedicated size column but width/height exist, those are already filled.
    # Add report fields.
    ws.cell(row_num, col_map["Match Status"], "PPT Extra")
    ws.cell(row_num, col_map["Matched PPT Slide"], slide_number)
    ws.cell(row_num, col_map["Confidence %"], 0)
    ws.cell(row_num, col_map["Match Reason"], "No matching Excel record")
    ws.cell(row_num, col_map["Missing Fields"], "")
    ws.cell(row_num, col_map["PPT Extra"], "Yes")
    ws.cell(row_num, col_map["Review Required"], "Yes")
    ws.cell(row_num, col_map["Remark"], "Ye PPT me extra hai")

    fill_entire_row(ws, row_num, GREEN_FILL, GREEN_FONT)

    return row_num


# ============================================================
# MAIN PROCESS
# ============================================================

def set_progress(progress, value):
    """Safe no-op progress hook for Streamlit Cloud compatibility."""
    if progress is not None:
        try:
            progress.progress(value)
        except Exception:
            pass


def process_files(pptx_file, excel_file, progress, status):
    status.write("⏳ Excel read ho rahi hai...")
    set_progress(progress, 8)

    df = load_excel(excel_file)
    mapping = detect_excel_columns(df)

    status.write("🔎 Excel columns identify ho rahe hain...")
    set_progress(progress, 15)

    excel_records = [
        build_excel_record(row, mapping, idx)
        for idx, (_, row) in enumerate(df.iterrows())
    ]

    status.write("📊 PowerPoint slides read ho rahi hain...")
    set_progress(progress, 22)

    prs = Presentation(pptx_file)
    slides = list(prs.slides)

    raw_slides = []
    ppt_records = []

    total = max(1, len(slides))

    for i, slide in enumerate(slides):
        raw = extract_text_from_slide(slide)
        raw_slides.append(raw)
        ppt_records.append(extract_slide_details(raw))
        set_progress(progress, 22 + int((i + 1) / total * 23))

    status.write("🧠 Smart matching engine candidates calculate kar raha hai...")
    set_progress(progress, 48)

    results, matched_excel, matched_ppt = match_all(
        excel_records, ppt_records
    )

    set_progress(progress, 70)
    status.write("📝 Excel report generate ho rahi hai...")

    # Make a fresh workbook from original dataframe.
    wb = dataframe_to_workbook(df)
    ws = wb.active

    col_map = append_report_columns(ws)

    # Update original Excel rows.
    for e_idx, e in enumerate(excel_records):
        excel_row = e_idx + 2

        if e_idx in results:
            r = results[e_idx]
            p_idx = r["ppt_idx"]
            p = ppt_records[p_idx]

            ws.cell(
                excel_row,
                col_map["Match Status"],
                "Matched"
            )
            ws.cell(
                excel_row,
                col_map["Matched PPT Slide"],
                p_idx + 1
            )
            ws.cell(
                excel_row,
                col_map["Confidence %"],
                r["confidence"]
            )

            reason = " + ".join(r["reasons"]) if r["reasons"] else "Fields matched"
            if r["conflicts"]:
                reason += " | " + ", ".join(r["conflicts"])

            ws.cell(
                excel_row,
                col_map["Match Reason"],
                reason
            )
            ws.cell(
                excel_row,
                col_map["Missing Fields"],
                get_missing_fields(e, p)
            )
            ws.cell(
                excel_row,
                col_map["PPT Extra"],
                "No"
            )

            review = "Yes" if r["confidence"] < 90 or r["conflicts"] else "No"
            ws.cell(
                excel_row,
                col_map["Review Required"],
                review
            )
            ws.cell(
                excel_row,
                col_map["Remark"],
                "Matched"
            )

            # Low confidence = yellow, otherwise normal
            if r["confidence"] < 90 or r["conflicts"]:
                fill_entire_row(ws, excel_row, YELLOW_FILL, YELLOW_FONT)

        else:
            ws.cell(
                excel_row,
                col_map["Match Status"],
                "Missing"
            )
            ws.cell(
                excel_row,
                col_map["Matched PPT Slide"],
                ""
            )
            ws.cell(
                excel_row,
                col_map["Confidence %"],
                0
            )
            ws.cell(
                excel_row,
                col_map["Match Reason"],
                "No sufficiently reliable PPT match"
            )
            ws.cell(
                excel_row,
                col_map["Missing Fields"],
                ""
            )
            ws.cell(
                excel_row,
                col_map["PPT Extra"],
                "No"
            )
            ws.cell(
                excel_row,
                col_map["Review Required"],
                "Yes"
            )
            ws.cell(
                excel_row,
                col_map["Remark"],
                "Ye PPT me missing hai"
            )

            # 🔴 Entire Excel row RED
            fill_entire_row(ws, excel_row, RED_FILL, RED_FONT)

    set_progress(progress, 82)

    # Add PPT extra slides as GREEN rows.
    extra_count = 0

    for p_idx, details in enumerate(ppt_records):
        if p_idx not in matched_ppt:
            add_extra_ppt_row(
                ws,
                col_map,
                mapping,
                details,
                p_idx + 1
            )
            extra_count += 1

    set_progress(progress, 92)

    # ---------------- PPT reorder ----------------
    matched_slide_indices = [
        results[e_idx]["ppt_idx"]
        for e_idx in range(len(excel_records))
        if e_idx in results
    ]

    ordered_indices = matched_slide_indices + [
        i for i in range(len(slides))
        if i not in matched_ppt
    ]

    sld_ids = prs.slides._sldIdLst
    original_ids = list(sld_ids)

    for sid in list(sld_ids):
        sld_ids.remove(sid)

    for idx in ordered_indices:
        sld_ids.append(original_ids[idx])

    out_ppt = io.BytesIO()
    prs.save(out_ppt)
    out_ppt.seek(0)

    # ---------------- Excel formatting ----------------
    autosize_worksheet(ws)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

    ws.row_dimensions[1].height = 28

    out_excel = io.BytesIO()
    wb.save(out_excel)
    out_excel.seek(0)

    set_progress(progress, 100)
    status.write("✅ Processing Complete!")

    matched_count = len(results)
    missing_count = len(excel_records) - matched_count

    return (
        out_ppt.getvalue(),
        out_excel.getvalue(),
        matched_count,
        missing_count,
        extra_count
    )


# ============================================================
# UI
# ============================================================

st.markdown("""
<div class="css-card">
    <div class="title-text">⚡ Smart PPT & Excel Matcher Pro V2</div>
    <div class="subtitle-text">
        Format-independent matching • Confidence score • Missing/Extra detection
        • Automatic PPT reorder • Professional Excel report
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------- Upload UI --------------------
st.markdown("""
<style>
.upload-label-red {
    color: #ff4b4b !important;
    font-size: 16px;
    font-weight: 800;
    margin-bottom: 6px;
}
.upload-label-green {
    color: #22c55e !important;
    font-size: 16px;
    font-weight: 800;
    margin-bottom: 6px;
}
.upload-box {
    border-radius: 10px;
    padding: 4px;
}
</style>
""", unsafe_allow_html=True)

left, right = st.columns(2)

with left:
    st.markdown(
        '<div class="upload-label-red">1️⃣ Upload PowerPoint</div>',
        unsafe_allow_html=True
    )
    uploaded_pptx = st.file_uploader(
        "PowerPoint",
        type=["pptx"],
        label_visibility="collapsed",
        key="ppt_upload"
    )

with right:
    st.markdown(
        '<div class="upload-label-green">2️⃣ Upload Master Excel</div>',
        unsafe_allow_html=True
    )
    uploaded_excel = st.file_uploader(
        "Master Excel",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
        key="excel_upload"
    )

st.caption(
    "Supported matching: Name, Mobile, SAP, Address, City/District, "
    "Size, Width, Height, Quantity and Media Type."
)

if st.button(
    "🚀 Process & Sync Files",
    type="primary",
    use_container_width=True
):
    if not uploaded_pptx or not uploaded_excel:
        st.warning("⚠️ Kripya PowerPoint aur Excel dono upload karein.")
    else:
        # Streamlit Cloud compatibility: use a status placeholder instead
        # of st.progress(), which can fail in some widget execution contexts.
        progress = None
        status = st.empty()

        try:
            (
                out_ppt,
                out_excel,
                matched,
                missing,
                extra
            ) = process_files(
                uploaded_pptx,
                uploaded_excel,
                progress,
                status
            )

            st.session_state["out_ppt"] = out_ppt
            st.session_state["out_excel"] = out_excel
            st.session_state["matched"] = matched
            st.session_state["missing"] = missing
            st.session_state["extra"] = extra
            st.session_state["done"] = True

            time.sleep(0.4)
            if progress is not None:
                try:
                    progress.empty()
                except Exception:
                    pass
            status.empty()

        except Exception as e:
            if progress is not None:
                try:
                    progress.empty()
                except Exception:
                    pass
            status.empty()
            st.error(f"❌ Processing error: {e}")
            st.caption("Tip: Agar error repeat ho, PPT/Excel file ka format check karein.")

if st.session_state.get("done"):
    st.markdown("### 📊 Sync Result")

    c1, c2, c3 = st.columns(3)

    c1.metric("Matched", st.session_state["matched"])
    c2.metric("Missing in PPT", st.session_state["missing"])
    c3.metric("Extra PPT Slides", st.session_state["extra"])

    st.markdown("---")

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "📥 Download Sorted PPT",
            data=st.session_state["out_ppt"],
            file_name="Sorted_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

    with d2:
        st.download_button(
            "📥 Download Final Excel Report",
            data=st.session_state["out_excel"],
            file_name="PPT_Excel_Final_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.info(
        "🔴 Red = Excel record PPT me missing | "
        "🟢 Green = PPT slide Excel me extra | "
        "🟡 Yellow = Low-confidence / manual review"
    )
