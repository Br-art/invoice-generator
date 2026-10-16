import os
import io
import json
import re
import tempfile
import subprocess
from datetime import datetime, date
from typing import List, Tuple, Optional

import streamlit as st
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.styles import Alignment, Border, Side

# =========================
# Fixed cell mapping (template)
# =========================
DATE_CELL = "I6"
DOCNO_CELL = "I12"

# Deliver to multi-row block (Rows 18 to 25, Columns G to I)
DELIVER_TO_START_ROW = 18
DELIVER_TO_END_ROW = 25
DELIVER_TO_COL_START = 7  # Column G
DELIVER_TO_COL_END = 9    # Column I

# Flight sector rows (E35 to E38 for routes, F35 to F38 for refs, I35 to I38 for prices)
ROUTE_START_ROW = 35
ROUTE_END_ROW = 38

# Currency symbols
CURRENCY_SYMBOL = {
    "ZMW": "K",
    "ZAR": "R",
    "USD": "$",
    "GBP": "£",
    "EUR": "€",
}

# Persistence
SETTINGS_FILE = "settings.json"
TEMPLATE_FILE = "template.xlsx"

# =========================
# Settings helpers
# =========================
def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_settings(d: dict) -> None:
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f)

settings = load_settings()

# =========================
# Parsing helpers
# =========================
def parse_reference(txt: str) -> Optional[str]:
    pattern = re.compile(r"\b(?=[A-Z0-9]{5,6}\b)(?=.*[A-Z])[A-Z0-9]{5,6}\b")
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    tail = "\n".join(lines[-12:]) if lines else txt

    m_last = None
    for m in pattern.finditer(tail):
        m_last = m
    if m_last:
        return m_last.group(0)

    m2_last = None
    for m in pattern.finditer(txt):
        m2_last = m
    return m2_last.group(0) if m2_last else None

def parse_total_currency(txt: str) -> Tuple[Optional[str], Optional[float]]:
    m = re.search(r"Total\s+([A-Z]{3})\s+([0-9]+(?:\.[0-9]{1,2})?)", txt)
    if not m:
        return None, None
    return m.group(1), float(m.group(2))

def parse_flights(txt: str) -> List[Tuple[str, str]]:
    flights: List[Tuple[str, str]] = []
    for line in txt.splitlines():
        m = re.search(r"(\d{2}[A-Z]{3}\d{2})\s*([A-Z]{3})\s+([A-Z]{3})", line)
        if m:
            ddmmmyy, orig, dest = m.groups()
            try:
                d = datetime.strptime(ddmmmyy, "%d%b%y").date()
                date_str = d.strftime("%d.%m.%Y")
            except Exception:
                date_str = ddmmmyy
            flights.append((f"{orig}-{dest}", date_str))
    return flights

def pretty_name(slash_name: str) -> str:
    parts = slash_name.split("/")
    if len(parts) != 2:
        return slash_name.title()

    last = parts[0].title()
    given_raw = parts[1]

    ch_match = re.search(r'\.CH(\d+)$', given_raw)
    in_match = re.search(r'\.IN(\d+)$', given_raw)

    suffix_str = ""
    if ch_match:
        age = ch_match.group(1)
        suffix_str = f" {age}yrs"
        given_raw = re.sub(r'\.CH\d+$', '', given_raw)
    elif in_match:
        age = in_match.group(1)
        suffix_str = f" IN {age}"
        given_raw = re.sub(r'\.IN\d+$', '', given_raw)

    title_map = {
        "MR": "Mr",
        "MS": "Ms",
        "MRS": "Mrs",
        "MSTR": "Mstr",
        "MISS": "Miss",
    }
    
    m = re.match(r"^([A-Z]+?)(MR|MS|MRS|MSTR|MISS)?$", given_raw)
    title = "Mr/Ms"
    given = given_raw.title()

    if m:
        core, t = m.groups()
        given = core.title()
        if t and t in title_map:
            title = title_map[t]

    return f"{title} {given} {last}{suffix_str}"

def parse_passengers_for_chunk(txt: str) -> List[str]:
    # Check for group booking headers (e.g. 1.13TAMEDIA*)
    group_pax = []
    for line in txt.splitlines():
        m_group = re.search(r'\b\d+\.(\d+)([A-Z0-9]+)\b', line)
        if m_group:
            count_str, name_part = m_group.groups()
            if count_str.isdigit() and int(count_str) > 1 and name_part.isalpha():
                group_pax.append(f"{count_str} {name_part.title()} (group {count_str} passengers)")

    if group_pax:
        return group_pax

    pax_raw = []
    for line in txt.splitlines():
        m_line = re.match(r"^\s*\d+\.\d+([A-Z]+/[A-Z\.0-9]+)", line)
        if m_line:
            pax_raw.append(m_line.group(1))

    if not pax_raw:
        for m in re.finditer(r"TICKET ISSUED .*?\s([A-Z]+/[A-Z]+[A-Z\.0-9]*)", txt):
            pax_raw.append(m.group(1))

    seen_base_names = set()
    out = []

    for raw in pax_raw:
        parts = raw.split('/')
        if len(parts) == 2:
            last = parts[0].strip().upper()
            given_part = parts[1].strip().upper()
            given_clean = re.sub(r'(MR|MS|MRS|MSTR|MISS|\.CH\d+|\.IN\d+)', '', given_part).strip()
            
            base_key = f"{last}/{given_clean}"
            if base_key not in seen_base_names:
                seen_base_names.add(base_key)
                out.append(pretty_name(raw))
        else:
            formatted = pretty_name(raw)
            if formatted not in out:
                out.append(formatted)

    return out

def split_tickets(big_text: str) -> List[str]:
    s = big_text.strip()
    if not s:
        return []

    parts = re.split(r"\n\s*---\s*\n", s)
    if len(parts) > 1:
        return [p.strip() for p in parts if p.strip()]

    parts = re.split(r"(?:\r?\n){2,}", s)
    return [p.strip() for p in parts if p.strip()]

def parse_one_ticket(txt: str) -> dict:
    ref = parse_reference(txt) or "REFXXX"
    ccy, total = parse_total_currency(txt)
    flights = parse_flights(txt)
    pax = parse_passengers_for_chunk(txt)

    return {
        "ref": ref,
        "currency": ccy or "ZMW",
        "total": float(total or 0.0),
        "flights": flights,
        "pax": pax,
    }

def ticket_route_string(flights: List[Tuple[str, str]]) -> str:
    if not flights:
        return "N/A"

    first = flights[0][0].split("-")
    chain = [first[0], first[1]] if len(first) == 2 else [flights[0][0]]

    for seg, _ in flights[1:]:
        parts = seg.split("-")
        chain.append(parts[1] if len(parts) == 2 else seg)

    return "-".join(chain)

# =========================
# Excel writers & Format Fixes
# =========================
def remove_payment_gridlines(ws) -> None:
    for r in range(63, 72):
        for col_letter in ['D', 'G']:
            cell = ws[f'{col_letter}{r}']
            current_border = cell.border
            cell.border = Border(
                left=current_border.left if current_border else None,
                right=Side(style=None),
                top=current_border.top if current_border else None,
                bottom=current_border.bottom if current_border else None
            )

    for coord in ['D67', 'D68', 'H68', 'I67', 'I68']:
        cell = ws[coord]
        curr = cell.border
        cell.border = Border(
            left=Side(style=None),
            right=Side(style=None),
            top=curr.top if curr else None,
            bottom=curr.bottom if curr else None
        )

    for coord in ['E67', 'E68']:
        cell = ws[coord]
        cell.border = Border(
            left=Side(style=None),
            right=Side(style='thin'),
            top=Side(style=None),
            bottom=Side(style=None)
        )

def write_deliver_to(ws, customer: str) -> None:
    for r in range(DELIVER_TO_START_ROW, DELIVER_TO_END_ROW + 1):
        for merged_range in list(ws.merged_cells.ranges):
            if f"G{r}" in merged_range or f"H{r}" in merged_range or f"I{r}" in merged_range:
                ws.unmerge_cells(str(merged_range))
        
        for c in range(DELIVER_TO_COL_START, DELIVER_TO_COL_END + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = ""
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for r in range(DELIVER_TO_START_ROW, DELIVER_TO_END_ROW + 1):
        ws.merge_cells(start_row=r, start_column=DELIVER_TO_COL_START, end_row=r, end_column=DELIVER_TO_COL_END)

    chars_per_line = 35 
    words = customer.strip().split()
    
    lines = []
    current_line = ""
    for word in words:
        if current_line:
            test_line = f"{current_line} {word}"
        else:
            test_line = word
            
        if len(test_line) <= chars_per_line:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    if current_line:
        lines.append(current_line)

    for i, line_text in enumerate(lines[:(DELIVER_TO_END_ROW - DELIVER_TO_START_ROW + 1)]):
        r = DELIVER_TO_START_ROW + i
        cell = ws.cell(row=r, column=DELIVER_TO_COL_START)
        cell.value = line_text
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")

def write_flight_sectors_and_summary(ws, ticket_items: List[dict], adjusted_data: List[tuple]) -> None:
    for r in range(ROUTE_START_ROW, ROUTE_END_ROW + 1):
        ws[f"E{r}"] = ""
        ws[f"F{r}"] = ""
        ws[f"I{r}"] = ""

    max_slots = ROUTE_END_ROW - ROUTE_START_ROW + 1
    
    if len(ticket_items) < max_slots:
        for idx, t in enumerate(ticket_items):
            r = ROUTE_START_ROW + idx
            sym = CURRENCY_SYMBOL.get(t["currency"], "K")
            ws[f"E{r}"] = ticket_route_string(t["flights"])
            ws[f"F{r}"] = t["ref"]
            
            adj_total, label = adjusted_data[idx] if idx < len(adjusted_data) else (t["total"], "")
            if label:
                ws[f"I{r}"] = f"{sym}{adj_total:,.2f} ({label})"
            else:
                ws[f"I{r}"] = f"{sym}{adj_total:,.2f}"
    else:
        for idx in range(max_slots - 1):
            r = ROUTE_START_ROW + idx
            t = ticket_items[idx]
            sym = CURRENCY_SYMBOL.get(t["currency"], "K")
            ws[f"E{r}"] = ticket_route_string(t["flights"])
            ws[f"F{r}"] = t["ref"]
            
            adj_total, label = adjusted_data[idx] if idx < len(adjusted_data) else (t["total"], "")
            if label:
                ws[f"I{r}"] = f"{sym}{adj_total:,.2f} ({label})"
            else:
                ws[f"I{r}"] = f"{sym}{adj_total:,.2f}"
            
        overflow_items = ticket_items[max_slots - 1:]
        r = ROUTE_END_ROW
        
        ws[f"E{r}"] = "(+more)"
        
        ref_list = [x["ref"] for x in overflow_items]
        if len(ref_list) > 4:
            remaining_refs = ", ".join(ref_list[:4]) + ", +"
        else:
            remaining_refs = ", ".join(ref_list)
            
        ws[f"F{r}"] = remaining_refs
        
        overflow_sum = sum(adjusted_data[i][0] for i in range(max_slots - 1, len(adjusted_data))) if len(adjusted_data) >= max_slots else sum(x["total"] for x in overflow_items)
        first_ccy = overflow_items[0]["currency"] if overflow_items else "ZMW"
        sym = CURRENCY_SYMBOL.get(first_ccy, "K")
        ws[f"I{r}"] = f"{sym}{overflow_sum:,.2f}"

def write_flight_dates(ws, ticket_items: List[dict]) -> None:
    for r in range(40, 45):
        for col_letter in ["B", "C", "D", "E", "F", "G", "H", "I"]:
            ws[f"{col_letter}{r}"] = ""

    all_flights = []
    for t in ticket_items:
        if t["flights"]:
            all_flights.extend(t["flights"])
        else:
            all_flights.append(("N/A", ""))

    if len(all_flights) > 0:
        ws["B40"] = "Flight dates:"
        for idx, (route, d) in enumerate(all_flights[:4]):
            r = 41 + idx
            ws[f"B{r}"] = route
            ws[f"C{r}"] = d

    if len(all_flights) > 4:
        ws["E40"] = "Flight dates:"
        for idx, (route, d) in enumerate(all_flights[4:8]):
            r = 41 + idx
            ws[f"E{r}"] = route
            ws[f"F{r}"] = d

    if len(all_flights) > 8:
        ws["H40"] = "Flight dates:"
        for idx, (route, d) in enumerate(all_flights[8:11]):
            r = 41 + idx
            ws[f"H{r}"] = route
            ws[f"I{r}"] = d
        
        if len(all_flights) > 11:
            ws["H44"] = "(+more)"
            ws["I44"] = "(+more)"

def write_passengers(ws, pax: List[str]) -> None:
    for r in range(45, 56):
        for col in ["B", "C", "D", "E", "F", "G", "H", "I"]:
            ws[f"{col}{r}"] = ""

    if not pax:
        return

    ws["B45"] = "Passengers:"
    for idx, name in enumerate(pax[:9]):
        ws[f"B{46 + idx}"] = name

    if len(pax) > 9:
        ws["E45"] = "Passengers:"
        for idx, name in enumerate(pax[9:18]):
            ws[f"E{46 + idx}"] = name

    if len(pax) > 18:
        ws["H45"] = "Passengers:"
        for idx, name in enumerate(pax[18:26]):
            ws[f"H{46 + idx}"] = name
            
        remaining_count = len(pax) - 26
        if remaining_count > 0:
            ws["H54"] = f"+{remaining_count} more"

# =========================
# Page setup for PDF export
# =========================
def configure_print_settings(ws) -> None:
    ws.print_area = "A1:I75"

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False

    ws.page_margins.left = 0.15
    ws.page_margins.right = 0.15
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0

# =========================
# Apply: single/multi-ticket
# =========================
def apply_to_workbook(wb: Workbook, ticket_text: str, customer: str, adjusted_data: List[tuple]) -> Workbook:
    ws = wb.active

    chunks = split_tickets(ticket_text)
    items = [parse_one_ticket(c) for c in chunks if c]
    if not items:
        items = [parse_one_ticket(ticket_text)]

    first_ref = items[0]["ref"] if items else "REFXXX"
    ws[DOCNO_CELL] = f"{first_ref}{date.today():%d%m%y}"
    write_deliver_to(ws, customer)
    ws[DATE_CELL] = date.today()

    ref_list = ", ".join([x["ref"] for x in items])
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "ABBNWU" in cell.value:
                cell.value = cell.value.replace("ABBNWU", ref_list)

    write_flight_sectors_and_summary(ws, items, adjusted_data)
    write_flight_dates(ws, items)

    union = []
    for t in items:
        for p in t["pax"]:
            if p not in union:
                union.append(p)
    write_passengers(ws, union)

    grand = sum(item[0] for item in adjusted_data)
    first_ccy = items[0]["currency"] if items else "ZMW"
    sym = CURRENCY_SYMBOL.get(first_ccy, "K")
    amt = f"{sym}{grand:,.2f}"

    ws["H56"] = "Sub Total"
    ws["H59"] = "Total Payable"
    ws["I56"] = amt
    ws["I59"] = amt

    remove_payment_gridlines(ws)
    configure_print_settings(ws)
    return wb

# =========================
# PDF export from Excel file
# =========================
def get_soffice_path() -> Optional[str]:
    candidates = [
        os.environ.get("SOFFICE_PATH", ""),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",
    ]
    for path in candidates:
        if not path:
            continue
        if path == "soffice":
            return path
        if os.path.exists(path):
            return path
    return None

def convert_excel_to_pdf(xlsx_path: str) -> bytes:
    soffice = get_soffice_path()
    if not soffice:
        raise RuntimeError("LibreOffice not found. Install LibreOffice or set SOFFICE_PATH.")

    outdir = tempfile.mkdtemp(prefix="invoice_pdf_")

    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", outdir,
        xlsx_path,
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip() or "PDF conversion failed."
        raise RuntimeError(err)

    pdf_path = os.path.join(
        outdir,
        os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf"
    )

    if not os.path.exists(pdf_path):
        raise RuntimeError("PDF file was not created.")

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    if not pdf_bytes:
        raise RuntimeError("PDF file is empty.")

    return pdf_bytes

# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Invoice Generator", page_icon="📄")
st.title("Invoice Generator")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, TEMPLATE_FILE)

if not os.path.exists(TEMPLATE_PATH):
    st.error("template.xlsx is missing from the app folder.")
    st.stop()

if "deliver_to" not in st.session_state:
    st.session_state.deliver_to = settings.get("last_deliver_to", "")

customer = st.text_input(
    "Deliver To (customer/company)",
    value=st.session_state.deliver_to
)
st.session_state.deliver_to = customer

ticket_text = st.text_area(
    "Paste ticket text (one or multiple tickets — separate multiple with a blank line or a line with ---)",
    height=280
)

adjusted_data = []

if ticket_text.strip():
    chunks = split_tickets(ticket_text)
    items = [parse_one_ticket(c) for c in chunks if c]
    if not items:
        items = [parse_one_ticket(ticket_text)]

    st.markdown("### 📋 Booking Fare Adjustments")
    
    for idx, t in enumerate(items):
        # Determine effective passenger count (if group booking, check group size or len(t['pax']))
        pax_count = len(t["pax"])
        for p in t["pax"]:
            m_grp = re.match(r"^(\d+)\s+", p)
            if m_grp:
                pax_count = max(pax_count, int(m_grp.group(1)))

        base_total = t["total"]
        sym = CURRENCY_SYMBOL.get(t["currency"], "K")
        ref = t["ref"]
        
        if pax_count >= 10:
            st.warning(f"⚠️ **Booking #{idx+1} (Ref: {ref})** has **{pax_count} passengers** (>= 10). Total Fare: {sym}{base_total:,.2f}")
            
            option = st.selectbox(
                f"Select calculation basis for Booking #{idx+1} ({ref}):",
                ("100% (Full Fare)", "75% Deposit", "25% Deposit", "Custom Advised Amount"),
                key=f"opt_{idx}"
            )
            
            if option == "75% Deposit":
                adj = base_total * 0.75
                label = "75%"
            elif option == "25% Deposit":
                adj = base_total * 0.25
                label = "25%"
            elif option == "Custom Advised Amount":
                adj = st.number_input(
                    f"Enter advised amount for Booking #{idx+1} ({ref}) [{sym}]:",
                    min_value=0.0,
                    value=float(base_total),
                    step=100.0,
                    key=f"cust_{idx}"
                )
                label = "Custom"
            else:
                adj = base_total
                label = ""
            
            adjusted_data.append((adj, label))
        else:
            st.info(f"ℹ️ **Booking #{idx+1} (Ref: {ref})** has **{pax_count} passengers** (<= 9). Fare remains at 100% ({sym}{base_total:,.2f}).")
            adjusted_data.append((base_total, ""))

if st.button("Generate Invoice", type="primary"):
    if not customer.strip():
        st.error("Please enter the Deliver To (customer).")
        st.stop()

    if not ticket_text.strip():
        st.error("Please paste the ticket text.")
        st.stop()

    wb = load_workbook(TEMPLATE_PATH)
    wb = apply_to_workbook(wb, ticket_text, customer, adjusted_data)

    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", customer.strip())
    base_name = f"Invoice_{safe}_{date.today():%Y%m%d}"

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_bytes = excel_buffer.getvalue()

    pdf_bytes = None
    pdf_error = None

    try:
        workdir = tempfile.mkdtemp(prefix="invoice_work_")
        xlsx_path = os.path.join(workdir, f"{base_name}.xlsx")

        with open(xlsx_path, "wb") as f:
            f.write(excel_bytes)

        pdf_bytes = convert_excel_to_pdf(xlsx_path)

    except Exception as e:
        pdf_error = repr(e)

    st.success("Invoice generated ✅")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "⬇️ Download Excel Invoice",
            data=excel_bytes,
            file_name=f"{base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col2:
        if pdf_bytes:
            st.download_button(
                "⬇️ Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"{base_name}.pdf",
                mime="application/pdf",
            )
        else:
            st.warning(f"PDF not available: {pdf_error}")

    save_settings({"last_deliver_to": customer})
