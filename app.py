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

# Excel/OpenXML does not allow these control characters.
def safe_excel_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    # Remove XML 1.0 forbidden control characters, including vertical tab (\x0b).
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", str(value))


def sanitize_dataframe_for_excel(df):
    df = df.copy()
    # Clean illegal characters from headers and make blank/duplicate headers safe.
    new_cols = []
    used = {}
    for i, col in enumerate(df.columns, 1):
        name = safe_excel_text(col).strip() or f"Column {i}"
        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1
        new_cols.append(name)
    df.columns = new_cols

    for col in df.columns:
        df[col] = df[col].map(safe_excel_text)
    return df


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
    """Safe column detection. Uses word/token matching so 'id' cannot match 'width'."""
    exclude = exclude or []
    scored = []

    def tokens(s):
        return set(re.findall(r"[a-z0-9]+", clean_text(s)))

    for col in columns:
        c = clean_text(col)
        if any(x in c for x in exclude):
            continue

        ct = tokens(c)
        score = 0

        for kw in keywords:
            kt = tokens(kw)
            if c == clean_text(kw):
                score += 200
            elif kt and kt.issubset(ct):
                score += 80
            elif clean_text(kw) in ct:
                score += 60
            elif len(clean_text(kw).replace(" ", "")) >= 4 and clean_text(kw).replace(" ", "") in re.sub(r"[^a-z0-9]", "", c):
                # Allow MobileNumber, ContactNumber, SAPCode, etc.
                score += 45

        if score:
            scored.append((score, col))

    if not scored:
        return None

    scored.sort(reverse=True, key=lambda x: x[0])
    return scored[0][1]

def detect_excel_columns(df):
    """Detect the business fields without confusing helper columns such as Area with City."""
    cols = list(df.columns)

    def pick_exact(names):
        for wanted in names:
            for c in cols:
                if clean_text(c) == clean_text(wanted):
                    return c
        return None

    name = pick_exact([
        "Outlet Name", "Retailer Name", "Dealer Name", "Customer Name",
        "Name of firm", "DEALER NAME", "DEALER  NAME"
    ]) or find_column(
        cols,
        ["outlet name", "retailer name", "dealer name", "customer name",
         "name of firm", "dealer name", "outlet"]
    )

    mobile = pick_exact([
        "MobileNumber", "Mobile", "Contact No. of Party", "Contact Number",
        "DEALER MOBILE NO", "Dealer Mobile No"
    ]) or find_column(
        cols,
        ["mobile number", "mobile", "contact number", "contact no of party",
         "dealer mobile no", "phone number"]
    )

    sap = pick_exact([
        "SAP Code", "SAP", "Dealer Code", "Retailer Code",
        "PARTY CODE", "Party Code"
    ]) or find_column(
        cols,
        ["sap code", "dealer code", "retailer code", "party code"]
    )

    address = pick_exact([
        "Address", "ADDRESS", "DEALER ADDRESS", "Dealer Address"
    ]) or find_column(cols, ["address", "dealer address"])

    city = pick_exact([
        "City", "CITY", "City / Town", "District"
    ])
    if not city:
        city = find_column(cols, ["city", "district", "city town", "town"])

    # Never treat Area(Sq.FT.) / Total / allocation columns as city.
    if city and re.search(r"area|sq\.?\s*ft|allocation|total", clean_text(city)):
        city = None

    width = pick_exact([
        "Width (inches)", "W", "Widht", "Width", "WIDTH"
    ]) or find_column(cols, ["width", "widht"])

    height = pick_exact([
        "Height (inches)", "H", "Hight", "Height", "HEIGHT"
    ]) or find_column(cols, ["height", "hight"])

    size = pick_exact(["Size", "Dimension", "Dimensions"])

    quantity = pick_exact([
        "Qty", "Qyt", "QTY", "Quantity", "board_qty"
    ]) or find_column(cols, ["qty", "quantity", "qyt"])

    media_type = pick_exact([
        "Media Type", "TYPE", "Type", "type"
    ]) or find_column(cols, ["media type", "type"])

    return {
        "name": name,
        "mobile": mobile,
        "sap": sap,
        "address": address,
        "city": city,
        "width": width,
        "height": height,
        "size": size,
        "quantity": quantity,
        "media_type": media_type
    }

def load_excel(uploaded_file):
    filename = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if filename.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(data), dtype=str).fillna("")
    elif filename.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(data), dtype=str, engine="xlrd").fillna("")
    else:
        df = pd.read_excel(io.BytesIO(data), dtype=str, engine="openpyxl").fillna("")

    return sanitize_dataframe_for_excel(df)



def safe_excel_value(value):
    """Remove characters forbidden by openpyxl/Excel XML."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value)
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)

def dataframe_to_workbook(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sync Report"

    for col_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(1, col_idx, safe_excel_value(col))
        cell.fill = HEADER_FILL
        cell.font = WHITE_BOLD
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            ws.cell(row_idx, col_idx, safe_excel_value(value))

    return wb

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



def extract_phones(text):
    """Return all valid Indian 10-digit mobile numbers found in text."""
    text = clean_text(text)
    found = re.findall(r"(?<!\d)(?:\+?91[\s\-]?)?[6-9]\d{9}(?!\d)", text)
    result = []
    for x in found:
        d = re.sub(r"\D", "", x)
        if len(d) == 12 and d.startswith("91"):
            d = d[2:]
        if len(d) == 10 and d not in result:
            result.append(d)
    return result


def normalize_phones(value):
    return extract_phones(str(value)) if value else []


def extract_size_candidates(text):
    """
    Extract dimension pairs from a slide. Supports:
    12x3, 12 X 3, 12 × 3 and layouts where x is a separate text box.
    """
    t = str(text).replace("×", "x").replace("*", "x")
    pairs = []

    # Normal inline dimensions
    for m in re.finditer(
        r"(?<!\d)(\d+(?:\.\d+)?)\s*(?:ft|feet)?\s*x\s*"
        r"(\d+(?:\.\d+)?)(?:\s*(?:ft|feet))?(?!\d)",
        t, re.I
    ):
        pairs.append((normalize_dimension(m.group(1)), normalize_dimension(m.group(2))))

    # Newline / separate-box layout: number, number, x
    nums = re.findall(r"(?<!\d)(\d+(?:\.\d+)?)(?!\d)", t)
    # Don't blindly pair all numbers because mobile/SAP can appear.
    # Pairs are handled separately by the spatial parser below.
    return [(a,b) for a,b in pairs if a and b]


def extract_slide_details(text, slide=None):
    """Extract recce fields from both normal and the special Walkaroo template."""
    details = {
        "name": "", "mobile": "", "sap": "", "address": "", "city": "",
        "width": "", "height": "", "size": "", "quantity": "", "media_type": ""
    }

    raw = str(text or "").replace("\x0b", "\n").replace("\x0c", "\n")
    raw = re.sub(r"[\x00-\x08\x0e-\x1f]", "", raw)

    # Determine special Walkaroo layout from its labels.
    is_walkaroo = bool(re.search(r"\b(?:dealer[_ ]code|board[_ ]qty)\b", raw, re.I))

    if slide is not None and is_walkaroo:
        shapes = []
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False) and sh.text.strip():
                shapes.append({
                    "text": sh.text.strip(),
                    "x": sh.left,
                    "y": sh.top
                })

        # Top-area value boxes.
        top = [
            s for s in shapes
            if s["y"] < 3_000_000
            and not re.match(
                r"^(?:outlet name|address|city|contact no|mobile|installation date|dealer_code)\b",
                s["text"].strip(), re.I
            )
            and s["text"].strip().lower() not in {"far view", "close view"}
        ]

        # Name/address are the two non-phone, non-numeric value boxes.
        value_boxes = [
            s for s in top
            if not extract_phones(s["text"])
            and not re.fullmatch(r"[\d\s.,:/\-]+", s["text"].strip())
            and len(s["text"].strip()) >= 3
        ]
        value_boxes.sort(key=lambda s: (s["y"], s["x"]))

        if value_boxes:
            details["name"] = value_boxes[0]["text"].strip()
        if len(value_boxes) > 1:
            details["address"] = value_boxes[1]["text"].strip()

        phones = []
        for s in shapes:
            phones.extend(extract_phones(s["text"]))
        if phones:
            details["mobile"] = phones[0]

        # Type is the short alphabetic value in the Type area.
        for s in shapes:
            t = s["text"].strip()
            if 9_000_000 <= s["y"] <= 12_000_000 and s["x"] < 5_500_000:
                if re.fullmatch(r"[A-Za-z]{2,6}", t) and t.lower() not in {
                    "qty", "size", "type", "remarks", "board_qty"
                }:
                    details["media_type"] = t
                    break

        # Width/height are the two numeric boxes around the Size label.
        nums = []
        for s in shapes:
            t = s["text"].strip()
            if 9_000_000 <= s["y"] <= 12_000_000:
                if re.fullmatch(r"\d+(?:\.\d+)?", t):
                    if 6_000_000 <= s["x"] <= 9_800_000:
                        nums.append((s["x"], normalize_dimension(t)))

        nums.sort(key=lambda z: z[0])
        if len(nums) >= 2:
            details["width"] = nums[0][1]
            details["height"] = nums[-1][1]
            details["size"] = f"{details['width']}x{details['height']}"

        # Keep Qty conservative; board_qty is NOT automatically Qty.
        # If a numeric box is immediately in the Qty area, use it.
        for s in shapes:
            t = s["text"].strip()
            if 9_000_000 <= s["y"] <= 12_000_000 and 10_000_000 <= s["x"] <= 16_000_000:
                if re.fullmatch(r"\d+(?:\.\d+)?", t):
                    # Walkaroo files have a stray extra number; only accept 1-99.
                    n = normalize_number(t)
                    if n and float(n) <= 99:
                        details["quantity"] = n
                        break

        return details

    # ---------------- Normal label/value templates ----------------
    lines = [x.strip() for x in raw.splitlines() if x.strip()]

    def labeled_value(pattern):
        m = re.search(pattern, raw, re.I)
        return m.group(1).strip() if m else ""

    details["name"] = labeled_value(
        r"Outlet[ \t]*Name[ \t]*:[ \t]*([^\n\r]+)"
    )
    details["address"] = labeled_value(
        r"Address[ \t]*:[ \t]*([^\n\r]+)"
    )

    phones = extract_phones(raw)
    if phones:
        details["mobile"] = phones[0]

    details["sap"] = normalize_code(labeled_value(
        r"(?:SAP[ \t]*Code|Dealer[ \t]*Code|Retailer[ \t]*Code|SAP)[ \t]*:[ \t]*([A-Za-z0-9]+)"
    ))

    details["media_type"] = labeled_value(
        r"(?:Media[ \t]*Type|Display[ \t]*Type|Sign[ \t]*Type)[ \t]*:[ \t]*([A-Za-z0-9]+)"
    )

    details["quantity"] = normalize_number(labeled_value(
        r"(?:Qty|Quantity|Qnty|Board[ _\t]*Qty)[ \t]*:[ \t]*(\d+(?:\.\d+)?)"
    ))

    pairs = extract_size_candidates(raw)
    if pairs:
        details["width"], details["height"] = pairs[0]
        details["size"] = f"{details['width']}x{details['height']}"

    # Some templates have the label in one line and the value in the next.
    if not details["name"]:
        for i, line in enumerate(lines):
            if re.fullmatch(r"Outlet[ \t]*Name[ \t]*:?", line, re.I):
                for nxt in lines[i+1:]:
                    if not re.match(
                        r"^(Address|City|Mobile|Contact|Installation|Size|Media Type|Remarks|Qty|Type|SAP|Dealer Code)\b",
                        nxt, re.I
                    ) and not extract_phones(nxt):
                        details["name"] = nxt
                        break
                    if details["name"]:
                        break

    if not details["address"]:
        for i, line in enumerate(lines):
            if re.fullmatch(r"Address[ \t]*:?", line, re.I):
                for nxt in lines[i+1:]:
                    if (
                        nxt.strip() == details["name"].strip()
                        or extract_phones(nxt)
                        or re.fullmatch(r"[\d\s.,:/\-]+", nxt)
                        or re.match(
                            r"^(City|Mobile|Contact|Installation|Size|Media Type|Remarks|Qty|Type|SAP|Dealer Code)\b",
                            nxt, re.I
                        )
                    ):
                        continue
                    details["address"] = nxt
                    break

    # Adani's third sample has an invalid "Mobile : brand"; use any valid
    # 10-digit phone from the whole slide instead.
    if not details["mobile"] and phones:
        details["mobile"] = phones[0]

    # Fallback name from obvious value lines.
    if not details["name"]:
        blocked = (
            "outlet name", "address", "mobile", "contact", "installation",
            "size", "media type", "remarks", "qty", "quantity", "type",
            "sap code", "dealer code", "speaker notes", "s_no"
        )
        for line in lines:
            low = line.lower().strip(" :.-")
            if (
                len(line) >= 3
                and low not in blocked
                and not extract_phones(line)
                and not re.fullmatch(r"[\d\s.,:/\-]+", line)
                and low not in {"far view", "close view", "ok"}
            ):
                details["name"] = line
                break

    if not details["address"]:
        # Prefer the first long non-label line that is NOT the outlet name.
        for line in lines:
            low = line.lower().strip(" :.-")
            if (
                line.strip() != details["name"].strip()
                and len(line) >= 5
                and not extract_phones(line)
                and not re.fullmatch(r"[\d\s.,:/\-]+", line)
                and low not in {"far view", "close view", "ok", "remarks"}
                and not re.match(
                    r"^(Outlet Name|Address|Mobile|Contact|Installation|City|Size|Media Type|Remarks|Qty|Type|SAP|Dealer Code|Bangur|speaker notes|s_no)\b",
                    line, re.I
                )
            ):
                details["address"] = line
                break

    return details

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
        aa = normalize_phones(a)
        bb = normalize_phones(b)
        return 1.0 if set(aa).intersection(bb) else 0.0

    if field == "sap":
        aa = normalize_code(a)
        bb = normalize_code(b)
        return 1.0 if aa and bb and aa == bb else 0.0

    if field in {"width", "height", "quantity"}:
        aa = normalize_number(a)
        bb = normalize_number(b)
        return 1.0 if aa and bb and aa == bb else 0.0

    if field == "size":
        return size_similarity(a, b)

    return token_similarity(a, b)


def size_similarity(a, b):
    def pair(v):
        s = str(v).lower().replace("×", "x").replace("*", "x")
        m = re.search(r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)", s)
        if not m:
            return None
        return float(m.group(1)), float(m.group(2))

    pa, pb = pair(a), pair(b)
    if not pa or not pb:
        return 0.0

    # Allow W/H reversal.
    candidates = [
        abs(pa[0]-pb[0]) / max(pa[0], pb[0], 1),
        abs(pa[1]-pb[1]) / max(pa[1], pb[1], 1)
    ]
    direct = sum(candidates) / 2

    reverse = (
        abs(pa[0]-pb[1]) / max(pa[0], pb[1], 1) +
        abs(pa[1]-pb[0]) / max(pa[1], pb[0], 1)
    ) / 2

    error = min(direct, reverse)

    if error == 0:
        return 1.0
    if error <= 0.08:
        return 0.95
    if error <= 0.15:
        return 0.90
    if error <= 0.25:
        return 0.78
    if error <= 0.40:
        return 0.55
    return 0.0

def build_excel_record(row, mapping, row_idx):
    rec = {"row_idx": row_idx}

    for field, col in mapping.items():
        rec[field] = clean_text(row[col]) if col else ""

    # Keep ALL mobile numbers (some dealers have 2 contact numbers).
    rec["mobile"] = ",".join(normalize_phones(rec.get("mobile", "")))
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
    Match score designed for real recce files:
    - Mobile/SAP are strong identifiers.
    - Name/address identify the outlet.
    - Size/media/qty distinguish multiple boards for the same outlet.
    - Size tolerates small recce/OCR differences.
    """
    score = 0.0
    max_possible = 0.0
    reasons = []
    conflicts = []

    # Mobile
    em = normalize_phones(excel_rec.get("mobile", ""))
    pm = normalize_phones(ppt_rec.get("mobile", ""))
    if em and pm:
        max_possible += 40
        if set(em).intersection(pm):
            score += 40
            reasons.append("Mobile")
        else:
            conflicts.append("Mobile conflict")

    # SAP
    es = normalize_code(excel_rec.get("sap", ""))
    ps = normalize_code(ppt_rec.get("sap", ""))
    if es and ps:
        max_possible += 35
        if es == ps:
            score += 35
            reasons.append("SAP")
        else:
            conflicts.append("SAP conflict")

    # Name
    en = excel_rec.get("name", "")
    pn = ppt_rec.get("name", "")
    if en and pn:
        sim = token_similarity(en, pn)
        max_possible += 25
        if sim >= 0.90:
            score += 25
            reasons.append("Name")
        elif sim >= 0.72:
            score += 25 * sim
            reasons.append("Name similar")
        else:
            conflicts.append("Name conflict")

    # Address
    ea = excel_rec.get("address", "")
    pa = ppt_rec.get("address", "")
    if ea and pa:
        sim = token_similarity(ea, pa)
        max_possible += 12
        if sim >= 0.80:
            score += 12 * sim
            reasons.append("Address")
        elif sim >= 0.55:
            score += 5
            reasons.append("Address partial")

    # City
    ec = excel_rec.get("city", "")
    pc = ppt_rec.get("city", "")
    if ec and pc:
        sim = token_similarity(ec, pc)
        max_possible += 8
        if sim >= 0.75:
            score += 8 * sim
            reasons.append("City/District")

    # Media type
    et = normalize_text(excel_rec.get("media_type", ""))
    pt = normalize_text(ppt_rec.get("media_type", ""))
    if et and pt:
        max_possible += 8
        if et == pt:
            score += 8
            reasons.append("Media Type")

    # Size
    esz = excel_rec.get("size_norm", "")
    psz = ppt_rec.get("size", "")
    if esz and psz:
        sim = size_similarity(esz, psz)
        max_possible += 18
        if sim >= 0.90:
            score += 18
            reasons.append("Size")
        elif sim >= 0.75:
            score += 12
            reasons.append("Size close")

    # Quantity
    eq = normalize_number(excel_rec.get("quantity", ""))
    pq = normalize_number(ppt_rec.get("quantity", ""))
    if eq and pq:
        max_possible += 5
        if eq == pq:
            score += 5
            reasons.append("Qty")

    if max_possible <= 0:
        return 0.0, reasons, conflicts

    confidence = max(0.0, min(100.0, score / max_possible * 100))

    # Important: strong identifiers should not be diluted by noisy fields.
    if "Mobile" in reasons and ("Name" in reasons or "Size" in reasons):
        confidence = max(confidence, 92)

    if "SAP" in reasons and ("Name" in reasons or "Size" in reasons):
        confidence = max(confidence, 94)

    # Name + mobile is highly reliable even if PPT has bad size text.
    if "Mobile" in reasons and "Name" in reasons:
        confidence = max(confidence, 96)

    return round(confidence, 2), reasons, conflicts

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
    Two-stage matching:
    Stage 1: exact strong identifiers.
    Stage 2: global scored matching.
    Duplicate outlet names are separated using size/media/qty where possible.
    """
    results = {}
    assigned_excel = set()
    assigned_ppt = set()

    # ---------- Stage 1: exact mobile ----------
    mobile_map = {}
    for p_idx, p in enumerate(ppt_records):
        for mob in normalize_phones(p.get("mobile", "")):
            mobile_map.setdefault(mob, []).append(p_idx)

    for e_idx, e in enumerate(excel_records):
        em = normalize_phones(e.get("mobile", ""))
        if not em:
            continue

        candidate_idxs = []
        for mob in em:
            candidate_idxs.extend(mobile_map.get(mob, []))

        candidate_idxs = [x for x in dict.fromkeys(candidate_idxs) if x not in assigned_ppt]
        if len(candidate_idxs) == 1:
            p_idx = candidate_idxs[0]
            conf, reasons, conflicts = candidate_score(e, ppt_records[p_idx])
            if conf >= 70:
                results[e_idx] = {
                    "ppt_idx": p_idx,
                    "confidence": max(conf, 92),
                    "reasons": reasons or ["Mobile"],
                    "conflicts": conflicts
                }
                assigned_excel.add(e_idx)
                assigned_ppt.add(p_idx)

    # ---------- Stage 2: global scoring ----------
    candidates = []

    for e_idx, e in enumerate(excel_records):
        if e_idx in assigned_excel:
            continue

        for p_idx, p in enumerate(ppt_records):
            if p_idx in assigned_ppt:
                continue

            conf, reasons, conflicts = candidate_score(e, p)
            if conf > 0:
                candidates.append((conf, e_idx, p_idx, reasons, conflicts))

    candidates.sort(reverse=True, key=lambda x: x[0])

    for conf, e_idx, p_idx, reasons, conflicts in candidates:
        if e_idx in assigned_excel or p_idx in assigned_ppt:
            continue

        # Strong exact mobile/SAP can be accepted lower.
        strong = "Mobile" in reasons or "SAP" in reasons
        threshold = 70 if strong else 78

        if conf < threshold:
            continue

        # Avoid ambiguous weak matches.
        alternatives = [
            c for c in candidates
            if c[1] == e_idx and c[2] != p_idx and c[0] >= conf - 8
        ]
        if alternatives and conf < 88 and not strong:
            continue

        results[e_idx] = {
            "ppt_idx": p_idx,
            "confidence": conf,
            "reasons": reasons,
            "conflicts": conflicts
        }
        assigned_excel.add(e_idx)
        assigned_ppt.add(p_idx)

    return results, assigned_excel, assigned_ppt

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
            ws.cell(row_num, excel_col, safe_excel_text(value))

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
        ppt_records.append(extract_slide_details(raw, slide))
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

    # Final safety pass: remove any illegal XML/control characters from ALL cells.
    for row_cells in ws.iter_rows():
        for cell in row_cells:
            if isinstance(cell.value, str):
                cell.value = safe_excel_value(cell.value)

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
