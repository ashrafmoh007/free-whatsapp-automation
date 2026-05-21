#!/usr/bin/env python3
"""
Generates Falcon_Garage_Job_Tracker.xlsx  (v3.0 — 34 columns A-AH)
Upload to Google Sheets, then paste Falcon_Garage_Tracker.gs
into Extensions → Apps Script to activate full automation.
"""

import datetime
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment, GradientFill
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ──────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────
SHEET_NAME  = "JOB UPDATE TRACKER"
TITLE_ROW   = 1
SUB_ROW     = 2
GROUP_ROW   = 3
HEADER_ROW  = 4
DATA_START  = 5
TOTAL_COLS  = 34

GARAGE_NAME  = "Falcon Garage"
GARAGE_AR    = "ورشة فالكون"
GARAGE_TEL   = "+974 4486 2018 | 7072 5811 | 7003 3723 | 7731 8043"
GARAGE_WEB   = "www.falqatar.com"
GARAGE_EMAIL = "info@falqatar.com"
GARAGE_ADDR  = "Street 47, Bldg 190, Zone 57, Industrial Area, Doha - Qatar"

# Column index map (0-based for list access, 1-based for openpyxl)
C = dict(
    JOB_ID=1, DATE_IN=2, TIME_IN=3,
    STATUS=4, PRIORITY=5, JOB_TYPE=6, TECHNICIAN=7,
    CUST_NAME=8, PHONE=9, EMAIL=10, CUST_ID=11,
    PLATE=12, MAKE=13, MODEL=14, YEAR=15, VIN=16, ODOMETER=17, VEH_TYPE=18,
    COMPLAINT=19, WORK_DONE=20, PARTS_USED=21,
    PARTS_QAR=22, LABOUR_QAR=23, DISCOUNT=24, VAT_QAR=25, TOTAL_QAR=26,
    PAYMENT=27, PAY_STATUS=28,
    EST_REF=29, INV_REF=30, DATE_OUT=31, APPROVED_BY=32,
    WA_LINK=33, ALERT_SENT=34,
)

HEADERS = [
    "Job ID","Date In","Time In",
    "Status","Priority","Job Type","Technician",
    "Customer Name","Phone","Email","Customer ID",
    "Plate No.","Make","Model","Year","VIN","Odometer","Veh Type",
    "Complaint / Fault","Work Done","Parts Used",
    "Parts (QAR)","Labour (QAR)","Discount (QAR)","VAT (QAR)","Total (QAR)",
    "Payment Method","Pay Status",
    "Estimate Ref","Invoice Ref","Date Out","Approved By",
    "WhatsApp","Alert Sent",
]

GROUPS = [
    dict(label="🔧  JOB INFORMATION",            s=1,  e=7,  bg="1C2839", fg="F59E0B"),
    dict(label="👤  CUSTOMER  · العميل",          s=8,  e=11, bg="0D47A1", fg="FFFFFF"),
    dict(label="🚗  VEHICLE  · المركبة",          s=12, e=18, bg="0891B2", fg="FFFFFF"),
    dict(label="🔧  COMPLAINT & WORK · الشكوى",  s=19, e=21, bg="7C3AED", fg="FFFFFF"),
    dict(label="💰  FINANCIALS · المالية (QAR)", s=22, e=28, bg="065F46", fg="F59E0B"),
    dict(label="📋  ADMIN",                       s=29, e=32, bg="374151", fg="FFFFFF"),
    dict(label="📱  COMMUNICATION",               s=33, e=34, bg="166534", fg="FFFFFF"),
]

COL_WIDTHS = {
    1:13, 2:11, 3:8,  4:14, 5:12, 6:14, 7:13,
    8:17, 9:15, 10:20,11:12,
    12:11,13:11,14:11,15:7, 16:18,17:10,18:10,
    19:28,20:28,21:22,
    22:12,23:12,24:10,25:10,26:12,27:14,28:12,
    29:12,30:12,31:11,32:13,
    33:17,34:12,
}

STATUS_COLOURS = {
    "Open"       : ("FF4444","FFFFFF"),
    "In Progress": ("1565C0","FFFFFF"),
    "Body Shop"  : ("5D4037","FFFFFF"),
    "QC"         : ("6A1B9A","FFFFFF"),
    "Ready"      : ("2E7D32","FFFFFF"),
    "Delivered"  : ("004D40","CCFFCC"),
    "On Hold"    : ("37474F","CFD8DC"),
}


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(color="E2E8F0", size=9, bold=False, name="Calibri"):
    return Font(name=name, size=size, bold=bold, color=color)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def thin_border():
    s = Side(border_style="thin", color="374151")
    return Border(left=s, right=s, top=s, bottom=s)

def col_letter(n):
    return get_column_letter(n)

def fmt_qar(v):
    return f"QAR {v:,.2f}"

def vat_formula(row):
    v, w, x = col_letter(C["PARTS_QAR"]), col_letter(C["LABOUR_QAR"]), col_letter(C["DISCOUNT"])
    return f'=IF({v}{row}+{w}{row}>0,ROUND(MAX(0,{v}{row}+{w}{row}-{x}{row})*0.05,2),0)'

def total_formula(row):
    v, w, x, y = (col_letter(C["PARTS_QAR"]), col_letter(C["LABOUR_QAR"]),
                  col_letter(C["DISCOUNT"]),   col_letter(C["VAT_QAR"]))
    return f'=IF({v}{row}+{w}{row}>0,ROUND(MAX(0,{v}{row}+{w}{row}-{x}{row})+{y}{row},2),0)'

def wa_formula(row):
    ph = col_letter(C["PHONE"])
    jc = col_letter(C["JOB_ID"])
    nm = col_letter(C["CUST_NAME"])
    mk = col_letter(C["MAKE"])
    md = col_letter(C["MODEL"])
    pl = col_letter(C["PLATE"])
    st = col_letter(C["STATUS"])
    tot= col_letter(C["TOTAL_QAR"])
    return (
        f'=IF({ph}{row}="","",'
        f'HYPERLINK("https://wa.me/974"&REGEXREPLACE({ph}{row},"[^0-9]","")&"?text="&'
        f'ENCODEURL("Dear "&{nm}{row}&", your vehicle ("&{mk}{row}&" "&{md}{row}&" | Plate: "&{pl}{row}&") — Job: "&{jc}{row}&" | Status: "&{st}{row}&" | Total: QAR "&{tot}{row}&". Thank you — Falcon Garage, Doha +974 4486 2018"),'
        f'"📱 WhatsApp"))'
    )


# ──────────────────────────────────────────────────────────────
# SAMPLE DATA (5 rows — JC-2026-001 to JC-2026-005)
# ──────────────────────────────────────────────────────────────
TODAY = datetime.date.today().strftime("%d/%m/%Y")

SAMPLE_ROWS = [
    # JOB_ID     DATE_IN  TIME_IN  STATUS            PRIORITY      JOB_TYPE            TECHNICIAN
    # CUST_NAME            PHONE           EMAIL                 CUST_ID
    # PLATE      MAKE      MODEL       YEAR   VIN                    ODOMETER   VEH_TYPE
    # COMPLAINT                WORK_DONE                  PARTS_USED
    # PARTS  LABOUR  DISC  VAT     TOTAL   PAYMENT          PAY_STATUS
    # EST_REF  INV_REF  DATE_OUT  APPROVED_BY  WA_LINK  ALERT_SENT
    [
        "JC-2026-001", TODAY, "08:30", "✅ Delivered", "🟢 NORMAL", "Oil Change", "Ahmed Al-Rashidi",
        "Mohammed Al-Hamad", "+974 5512 3456", "m.alhamad@email.com", "QID-28843721",
        "A 12345", "Toyota", "Land Cruiser", "2022", "JTMCV02J204045321", "62000 km", "SUV",
        "Routine oil change & filter", "Completed oil change, replaced filter", "Engine oil 5L, oil filter",
        350, 150, 0, None, None, "Cash", "Pending",
        "EST-001", "INV-001", TODAY, "Ahmed Al-Rashidi", None, None,
    ],
    [
        "JC-2026-002", TODAY, "09:15", "🟢 Ready", "🔴 URGENT", "AC Repair", "Khalid Al-Sayed",
        "Sarah Johnson", "+974 6623 4567", "s.johnson@gmail.com", "QID-39124568",
        "B 67890", "Nissan", "Patrol", "2020", "JN8AY2ND5L9760412", "45200 km", "SUV",
        "AC not cooling at all", "Replaced compressor & recharged", "AC Compressor, refrigerant gas",
        1200, 450, 0, None, None, "Card", "Pending",
        "EST-002", "INV-002", None, "Khalid Al-Sayed", None, None,
    ],
    [
        "JC-2026-003", TODAY, "10:00", "🔵 In Progress", "🟡 HIGH", "Body Repair", "Yousuf Ibrahim",
        "Ali Karimi", "+974 7734 5678", "ali.karimi@work.qa", "QID-45231890",
        "C 11111", "BMW", "5 Series", "2021", "WBA13BJ08MCF62831", "31500 km", "Sedan",
        "Front bumper damage", "Dent removal in progress", "Bumper paint, filler compound",
        800, 600, 50, None, None, "Bank Transfer", "Pending",
        "EST-003", None, None, "Yousuf Ibrahim", None, None,
    ],
    [
        "JC-2026-004", TODAY, "11:30", "🔴 Open", "🟢 NORMAL", "General Service", "Fatima Al-Dosari",
        "Rania Khalil", "+974 3345 6789", "rania.k@hotmail.com", "QID-56342901",
        "D 22222", "Mercedes", "C-Class", "2023", "WDD2050422R456789", "15000 km", "Sedan",
        "Full service check", "Awaiting technician assignment", "Service kit, brake pads",
        500, 300, 25, None, None, "Cash", "Pending",
        "EST-004", None, None, None, None, None,
    ],
    [
        "JC-2026-005", TODAY, "13:00", "🔧 Body Shop", "🟡 HIGH", "Tyres / Alignment", "Hassan Al-Mansoor",
        "James O'Brien", "+974 5567 8901", "james.ob@corp.net", "QID-67453012",
        "E 33333", "Ford", "F-150", "2019", "1FTEW1E55KKC24680", "78900 km", "Pick-up",
        "Wheel alignment & new tyres", "Alignment done, fitting tyres", "4x Michelin LTX tyres",
        1800, 250, 100, None, None, "Credit", "Pending",
        "EST-005", None, None, "Hassan Al-Mansoor", None, None,
    ],
]


# ──────────────────────────────────────────────────────────────
# BUILD WORKBOOK
# ──────────────────────────────────────────────────────────────
def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # ── Row 1: Main Title ──────────────────────────────────
    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=TOTAL_COLS)
    c = ws.cell(TITLE_ROW, 1)
    c.value = f"🦅  {GARAGE_NAME}  |  {GARAGE_AR}  —  Job Update Tracker"
    c.fill = fill("111827")
    c.font = font("F59E0B", 18, True)
    c.alignment = center()
    ws.row_dimensions[TITLE_ROW].height = 38

    # ── Row 2: Subtitle ────────────────────────────────────
    ws.merge_cells(start_row=SUB_ROW, start_column=1, end_row=SUB_ROW, end_column=TOTAL_COLS)
    c = ws.cell(SUB_ROW, 1)
    c.value = (f"📞 {GARAGE_TEL}   🌐 {GARAGE_WEB}   ✉ {GARAGE_EMAIL}   📍 {GARAGE_ADDR}")
    c.fill = fill("1F2937")
    c.font = font("9CA3AF", 9)
    c.alignment = center()
    ws.row_dimensions[SUB_ROW].height = 20

    # ── Row 3: Section Group Headers ───────────────────────
    ws.row_dimensions[GROUP_ROW].height = 20
    for g in GROUPS:
        if g["e"] > g["s"]:
            ws.merge_cells(start_row=GROUP_ROW, start_column=g["s"], end_row=GROUP_ROW, end_column=g["e"])
        c = ws.cell(GROUP_ROW, g["s"])
        c.value = g["label"]
        c.fill = fill(g["bg"])
        c.font = font(g["fg"], 9, True)
        c.alignment = center()

    # ── Row 4: Column Headers ──────────────────────────────
    ws.row_dimensions[HEADER_ROW].height = 32
    for i, h in enumerate(HEADERS):
        c = ws.cell(HEADER_ROW, i + 1)
        c.value = h
        c.fill = fill("0F172A")
        c.font = font("E2E8F0", 9, True)
        c.alignment = center(wrap=True)
        c.border = thin_border()

    # ── Sample Data Rows ───────────────────────────────────
    for i, row_data in enumerate(SAMPLE_ROWS):
        r = DATA_START + i
        for j, val in enumerate(row_data):
            if val is not None:
                ws.cell(r, j + 1).value = val

        # Status colour
        status_val = row_data[C["STATUS"] - 1]
        bg, fg = "1A2332", "D1D5DB"
        for kw, colours in STATUS_COLOURS.items():
            if kw in status_val:
                bg, fg = colours
                break

        for col_idx in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, col_idx)
            cell.fill = fill(bg)
            cell.font = font(fg, 9)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = thin_border()

        # Formulas
        ws.cell(r, C["VAT_QAR"  ]).value = vat_formula(r)
        ws.cell(r, C["TOTAL_QAR"]).value = total_formula(r)
        ws.cell(r, C["WA_LINK"  ]).value = wa_formula(r)

        ws.row_dimensions[r].height = 20

    # ── Zebra stripe for remaining empty rows (up to row 54) ──
    for r in range(DATA_START + len(SAMPLE_ROWS), DATA_START + 50):
        bg = "1A2332" if r % 2 == 0 else "111827"
        for col_idx in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, col_idx)
            cell.fill = fill(bg)
            cell.font = font("D1D5DB", 9)
        ws.row_dimensions[r].height = 18

    # ── Conditional Formatting (row-level status colours) ──
    status_col_letter = col_letter(C["STATUS"])
    for kw, (bg, fg) in STATUS_COLOURS.items():
        formula = [
            f'NOT(ISBLANK(${status_col_letter}{DATA_START}))'
            f'*ISNUMBER(SEARCH("{kw}",${status_col_letter}{DATA_START}))'
        ]
        dxf = DifferentialStyle(
            fill=PatternFill(bgColor=bg),
            font=Font(color=fg),
        )
        rule = Rule(type="expression", formula=formula, dxf=dxf)
        data_range = f"A{DATA_START}:{col_letter(TOTAL_COLS)}{DATA_START + 499}"
        ws.conditional_formatting.add(data_range, rule)

    # ── Data Validation dropdowns ──────────────────────────
    dv_range = f"{DATA_START}:{DATA_START + 499}"

    def add_dv(col_idx, formula1):
        dv = DataValidation(type="list", formula1=formula1, allow_blank=True, showDropDown=False)
        ws.add_data_validation(dv)
        dv.add(f"{col_letter(col_idx)}{DATA_START}:{col_letter(col_idx)}{DATA_START+499}")

    add_dv(C["STATUS"],    '"🔴 Open,🔵 In Progress,🔧 Body Shop,🔍 QC,🟢 Ready,✅ Delivered,⏸ On Hold"')
    add_dv(C["PRIORITY"],  '"🔴 URGENT,🟡 HIGH,🟢 NORMAL,🔵 LOW"')
    add_dv(C["JOB_TYPE"],  '"General Service,Oil Change,Body Repair,Electrical,AC Repair,Tyres / Alignment,Inspection,Custom / Other"')
    add_dv(C["VEH_TYPE"],  '"Sedan,SUV,Pick-up,Van,Truck,Motorcycle,Other"')
    add_dv(C["PAY_STATUS"],"\"Cash,Card,Bank Transfer,Credit,Pending\"")

    # ── Column Widths ──────────────────────────────────────
    for col_idx, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter(col_idx)].width = width

    # ── Freeze panes (freeze rows 1-4 + col A) ────────────
    ws.freeze_panes = ws.cell(HEADER_ROW + 1, 2)

    # ── Tab colour ─────────────────────────────────────────
    ws.sheet_properties.tabColor = "F59E0B"

    # ── Instructions sheet ─────────────────────────────────
    instr = wb.create_sheet("📋 Setup Guide")
    _build_instructions(instr)

    wb.save("Falcon_Garage_Job_Tracker.xlsx")
    print("✅  Falcon_Garage_Job_Tracker.xlsx generated successfully (v3.0 — 34 columns)")


# ──────────────────────────────────────────────────────────────
# INSTRUCTIONS SHEET
# ──────────────────────────────────────────────────────────────
def _build_instructions(ws):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 55

    def row(r, a="", b="", c="", bg="111827", fg="E2E8F0", bold=False, size=10):
        for col, val in [(1, a), (2, b), (3, c)]:
            cell = ws.cell(r, col)
            cell.value = val
            cell.fill = fill(bg)
            cell.font = font(fg, size, bold)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 20

    row(1, "", "🦅 FALCON GARAGE — SETUP GUIDE", "", bg="111827", fg="F59E0B", bold=True, size=14)
    row(2, "", "Google Sheets + Apps Script  v3.0", "", bg="1F2937", fg="9CA3AF")
    row(3)
    row(4,  "", "STEP 1: Upload this file to Google Sheets", "", bg="0D47A1", fg="FFFFFF", bold=True)
    row(5,  "", "1a. Go to sheets.new", "")
    row(6,  "", "1b. File → Import → Upload → select Falcon_Garage_Job_Tracker.xlsx", "")
    row(7,  "", "1c. Choose 'Replace spreadsheet' → Import data", "")
    row(8)
    row(9,  "", "STEP 2: Add the Apps Script", "", bg="065F46", fg="F59E0B", bold=True)
    row(10, "", "2a. Extensions → Apps Script", "")
    row(11, "", "2b. Delete all existing code in Code.gs", "")
    row(12, "", "2c. Paste the full content of Falcon_Garage_Tracker.gs", "")
    row(13, "", "2d. Save (Ctrl+S / Cmd+S)", "")
    row(14)
    row(15, "", "STEP 3: Run Setup", "", bg="374151", fg="FFFFFF", bold=True)
    row(16, "", "3a. Back in the spreadsheet, refresh the page (F5)", "")
    row(17, "", "3b. Click  🦅 Falcon Garage  in the menu bar", "")
    row(18, "", "3c. Click  ⚙ Setup Sheet (first-time)", "")
    row(19, "", "3d. Authorise the script when prompted", "")
    row(20)
    row(21, "", "DAILY USAGE", "", bg="0891B2", fg="FFFFFF", bold=True)
    row(22, "", "• Menu → New Job Card      → auto-generates JC-YYYY-NNN", "")
    row(23, "", "• Status dropdown          → row colour changes instantly", "")
    row(24, "", "• 📱 WhatsApp link (col AG)→ opens WhatsApp Web with pre-filled message", "")
    row(25, "", "• Menu → Email             → branded HTML update to customer", "")
    row(26, "", "• Menu → Monthly PDF       → saved to Google Drive", "")
    row(27)
    row(28, "", "COLUMNS  A–AH  (34 total)", "", bg="7C3AED", fg="FFFFFF", bold=True)
    row(29, "", "A  Job ID        B  Date In      C  Time In      D  Status       E  Priority", "")
    row(30, "", "F  Job Type      G  Technician   H  Cust Name    I  Phone        J  Email", "")
    row(31, "", "K  Cust ID       L  Plate        M  Make         N  Model        O  Year", "")
    row(32, "", "P  VIN           Q  Odometer     R  Veh Type     S  Complaint    T  Work Done", "")
    row(33, "", "U  Parts Used    V  Parts QAR    W  Labour QAR   X  Discount     Y  VAT QAR", "")
    row(34, "", "Z  Total QAR     AA Payment      AB Pay Status   AC Est Ref      AD Inv Ref", "")
    row(35, "", "AE Date Out      AF Approved By  AG WhatsApp     AH Alert Sent", "")
    row(36)
    row(37, "", f"Contact: {GARAGE_EMAIL}  |  {GARAGE_WEB}", "", bg="0F172A", fg="6B7280")

    ws.sheet_properties.tabColor = "374151"


# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    build()
