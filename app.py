import io
import json
import re
from pathlib import Path

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Border, Font, Side

st.set_page_config(page_title="Diary -> Quotation", page_icon=":camera:", layout="centered",
                   initial_sidebar_state="collapsed")
MODEL = "gemini-2.5-flash"
TITLE = "(MODULAR) FURNITURE (CENTURY CLUB PRIME WITH LAMINATE)"
DEFAULT_FREIGHT = 1800
OUTPUT_DIR = Path("quotations")
OUTPUT_DIR.mkdir(exist_ok=True)
TERMS = [
    "THE QUOTATION DOES NOT INCLUDE GST.",
    "THE QUOTATION IS VALID FOR 15 DAYS FROM THE DATE OF QUOTATION.",
    "FURNITURE WILL BE FITTED STRICTLY AS PER APPROVED DESIGN/DRAWING.",
    "DEPTH OF THE ALL WALL CABINET IS 300mm INCLUDING SHUTTER",
    "CENTURY CLUB PRIME WITH LAMINATE WILL BE USED AS A MATERIAL",
    "ALL HARDWARE (HINGE AND CHANNEL) SOURCED FROM HETTICH (GERMAN MAKER)",
    "THE PLYWOOD CONTAINS A WARRANTY OF 30YRS AGAINST TERMITE AND WATER DAMAGE",
]
PROMPT = r'''Read this handwritten quotation diary page for an Indian modular kitchen or furniture business.
Return ONLY valid JSON, with no markdown or explanation:
{"client_name":"","quotation_type":"kitchen"|"furniture","quotation_id":"","customer_id":"","measurement_date":"","measurement_taken_by":"","date_of_submission":"","confirmed_on":"","freight":0,"items":[{"sno":1,"description":"","code":"","qty":1,"amount":0}],"accessories":[{"item_sno":1,"description":"","code":"","qty":1,"amount":0}],"moulding_bottom":0,"moulding_top":0,"skirting":0,"notes":""}
Use kitchen when the page contains SINK, TROLLEY, CUTLERY, THALI, BASKET, CYLINDER, HOB, CHIMNEY, TALL UNIT, or WALL STORAGE. Copy descriptions and dimensions exactly. Expand 96k, 1.14L, and 66/-. Default qty to 1. Put right-hand baskets, holders, and trolleys in accessories. Never invent values.'''

ONES = ["", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"]
TENS = ["", "", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"]


def _two(number):
    return ONES[number] if number < 20 else TENS[number // 10] + (" " + ONES[number % 10] if number % 10 else "")


def _three(number):
    return _two(number) if number < 100 else ONES[number // 100] + " HUNDRED" + (" " + _two(number % 100) if number % 100 else "")


def num_to_words_indian(number):
    number = int(round(float(number)))
    if not number:
        return "ZERO"
    parts = []
    crore, number = divmod(number, 10000000)
    lakh, number = divmod(number, 100000)
    thousand, rest = divmod(number, 1000)
    if crore:
        parts.append(_three(crore) + " CRORE")
    if lakh:
        parts.append(_two(lakh) + " LAKH")
    if thousand:
        parts.append(_two(thousand) + " THOUSAND")
    if rest:
        parts.append(_three(rest))
    return parts[0] if len(parts) == 1 else " ".join(parts[:-1]) + " AND " + parts[-1]


def file_stem(client_name, quotation_type):
    words = re.findall(r"[A-Za-z]+", client_name or "")
    initials = "_".join(word[0].upper() for word in words) or "CLIENT"
    return f"{initials}({'k' if quotation_type == 'kitchen' else 'f'})"


def extract_quotation(images, api_key):
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=api_key)
    parts = [PROMPT] + [types.Part.from_bytes(data=data, mime_type=mime) for data, mime in images]
    response = client.models.generate_content(model=MODEL, contents=parts)
    text = re.sub(r"^```(?:json)?|```$", "", (response.text or "").strip(), flags=re.M).strip()
    return json.loads(text)


def compute_total(data):
    total = sum(float(item.get("amount") or 0) for item in data.get("items", []))
    total += sum(float(item.get("amount") or 0) for item in data.get("accessories", []))
    return total + sum(float(data.get(key) or 0) for key in ("moulding_bottom", "moulding_top", "skirting", "freight"))


def build_workbook(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    for column, width in {"A": 6, "B": 62, "C": 10, "D": 8, "E": 13, "F": 32, "G": 10, "H": 8, "I": 13}.items():
        sheet.column_dimensions[column].width = width
    fields = {"B1": "QUOTATION FOR", "C1": TITLE, "G1": "Quotation ID", "H1": data.get("quotation_id", ""), "B2": "PARTY NAME", "C2": data.get("client_name", ""), "G2": "Customer ID", "H2": data.get("customer_id", ""), "B3": "Measurement Date", "C3": data.get("measurement_date", ""), "F3": "Measurement Taken By", "G3": data.get("measurement_taken_by", ""), "B4": "DATE OF SUBMISSION", "C4": data.get("date_of_submission", ""), "F4": "CONFIRMED ON", "G4": data.get("confirmed_on", "")}
    for cell, value in fields.items():
        sheet[cell] = value
    for cell in ("B1", "B2", "B3", "B4", "F3", "F4", "G1", "G2"):
        sheet[cell].font = Font(bold=True)
    headers = ["S.NO", "DESCRIPTION", "CODE", "QNTY", "AMOUNT", "DESCRIPTION", "CODE", "QNTY", "AMOUNT"]
    for column, header in enumerate(headers, 1):
        sheet.cell(6, column, header).font = Font(bold=True)
    items, accessories = data.get("items", []) or [], data.get("accessories", []) or []
    start, end = 7, 7 + max(23, len(items)) - 1
    for index, item in enumerate(items):
        row = start + index
        for column, value in enumerate((f"{index + 1})", item.get("description", ""), item.get("code", ""), item.get("qty", 1), item.get("amount", 0)), 1):
            sheet.cell(row, column, value)
    used, next_row = set(), start
    for accessory in accessories:
        try:
            row = start + int(accessory.get("item_sno")) - 1
        except (TypeError, ValueError):
            row = None
        if row is None or row in used or row > end:
            while next_row in used:
                next_row += 1
            row = next_row
        used.add(row)
        for column, value in enumerate((accessory.get("description", ""), accessory.get("code", ""), accessory.get("qty", 1), accessory.get("amount", 0)), 6):
            sheet.cell(row, column, value)
    accessories_row = end + 1
    sheet.cell(accessories_row, 6, "ACCESSORIES TOTAL").font = Font(bold=True)
    sheet.cell(accessories_row, 9, f"=SUM(I{start}:I{end})")
    bottom, top, skirting, module = (accessories_row + offset for offset in range(1, 5))
    for row, label, key in ((bottom, "Moulding - Bottom", "moulding_bottom"), (top, "Moulding -Top", "moulding_top"), (skirting, "Skirting -  18 X 100K", "skirting")):
        sheet.cell(row, 2, label)
        if float(data.get(key) or 0):
            sheet.cell(row, 5, float(data[key]))
    sheet.cell(module, 2, "Module Total (A)").font = Font(bold=True)
    sheet.cell(module, 5, f"=SUM(E{start}:E{module - 1})")
    freight = module + 3
    sheet.cell(freight, 2, "FREIGHT  (C)")
    sheet.cell(freight, 5, float(data.get("freight") or DEFAULT_FREIGHT))
    grand = freight + 2
    sheet.cell(grand, 2, "Grand Total (A+B+C)").font = Font(bold=True)
    sheet.cell(grand, 5, f"=E{module}+I{accessories_row}+E{freight}")
    sheet.cell(grand + 1, 2, "ADVANCE PAYABLE ALONG WITH ORDER")
    sheet.cell(grand + 1, 5, f"=E{grand}/2")
    sheet.cell(grand + 2, 2, "AMOUNT PAYABLE BEFORE DELIVERY")
    sheet.cell(grand + 2, 5, f"=E{grand + 1}")
    total = compute_total(data)
    sheet.cell(grand + 3, 2, f"GRAND TOTAL : RUPEES {num_to_words_indian(total)} ONLY").font = Font(bold=True)
    row = grand + 5
    sheet.cell(row, 1, "DELIVERY TIME").font = Font(bold=True)
    sheet.cell(row + 1, 1, "TERMS AND CONDITIONS :").font = Font(bold=True)
    sheet.cell(row + 1, 3, "KITCHEN DELIVERY 45 DAYS")
    for offset, term in enumerate(TERMS, 2):
        sheet.cell(row + offset, 1, term)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row_cells in sheet.iter_rows(min_row=6, max_row=module, min_col=1, max_col=9):
        for cell in row_cells:
            cell.border = border
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue(), total


st.markdown("## :camera: Diary -> Quotation Converter")
st.caption("Upload photos of the handwritten diary page. The app reads them and produces a formatted Excel quotation.")
try:
    default_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    default_key = ""
api_key = default_key or st.text_input("Gemini API Key", type="password", placeholder="AIza...", help="Get a key at aistudio.google.com")
uploaded = st.file_uploader("Upload Diary Photos", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
if uploaded:
    columns = st.columns(min(3, len(uploaded)))
    for index, file in enumerate(uploaded):
        with columns[index % len(columns)]:
            st.image(file, caption=file.name, use_container_width=True)

if st.button("Generate Quotation", type="primary", use_container_width=True):
    if not api_key:
        st.error("Please enter a Gemini API key.")
    elif not uploaded:
        st.error("Please upload at least one diary photo.")
    else:
        try:
            with st.spinner("Reading handwriting..."):
                data = extract_quotation([(file.getvalue(), file.type or "image/jpeg") for file in uploaded], api_key)
            with st.spinner("Building Excel quotation..."):
                xlsx, total = build_workbook(data)
                filename = f"{file_stem(data.get('client_name', ''), data.get('quotation_type', 'furniture'))}.xlsx"
                (OUTPUT_DIR / filename).write_bytes(xlsx)
            st.session_state.update(xlsx=xlsx, filename=filename, data=data, total=total)
        except json.JSONDecodeError:
            st.error("Gemini returned an unreadable response. Try clearer photos.")
        except Exception as error:
            st.error(f"Error: {error}")

if "xlsx" in st.session_state:
    data, total = st.session_state["data"], st.session_state["total"]
    st.success(f"**Client:** {data.get('client_name', '-')}  \n**Type:** {data.get('quotation_type', '-').title()}  \n**Grand Total:** INR {total:,.0f}  \n*Rupees {num_to_words_indian(total).title()} Only*")
    st.download_button("Download Excel Quotation", data=st.session_state["xlsx"], file_name=st.session_state["filename"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    rows = [{"S.No": index, "Description": item.get("description", ""), "Qty": item.get("qty", 1), "Amount": item.get("amount", 0)} for index, item in enumerate(data.get("items", []), 1)]
    rows += [{"S.No": "-", "Description": item.get("description", ""), "Qty": item.get("qty", 1), "Amount": item.get("amount", 0)} for item in data.get("accessories", [])]
    st.dataframe(rows, use_container_width=True, hide_index=True)
st.markdown("---")
st.caption("Files are named by client initials, e.g. A_K_S(k).xlsx for a kitchen quote.")
