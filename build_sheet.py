#!/usr/bin/env python3
"""
Generates Falcon_Garage_Job_Tracker.xlsx
Upload to Google Sheets, then paste Falcon_Garage_Tracker.gs
into Extensions → Apps Script to activate full automation.
"""

import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

OUTPUT = "/home/user/free-whatsapp-automation/Falcon_Garage_Job_Tracker.xlsx"

# ── Palette ────────────────────────────────────────────────────────────────
TITLE_BG   = "0D2137"
HEADER_BG  = "1A3A5C"
BAND_BG    = "2E86AB"
WHITE      = "FFFFFF"
ROW_A      = "F0F7FF"
ROW_B      = "FFFFFF"
BORDER_C   = "BDC3C7"
ACCENT     = "1A6FAE"
LINK_C     = "1D6A39"

RED_BG, RED_FG     = "FADBD8", "7B241C"
ORG_BG, ORG_FG     = "FAE5D3", "784212"
BLU_BG, BLU_FG     = "D6EAF8", "1A5276"
GRN_BG, GRN_FG     = "D5F5E3", "1D6A39"
GRY_BG, GRY_FG     = "EAECEE", "5D6D7E"

def fill(hex_c):
    return PatternFill(start_color=hex_c, end_color=hex_c, fill_type="solid")

def font(bold=False, color=None, size=10, italic=False, underline=None):
    kw = dict(name="Arial", size=size, bold=bold, italic=italic)
    if color:   kw["color"] = color
    if underline: kw["underline"] = underline
    return Font(**kw)

def border(color=BORDER_C, style="thin"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

# ── Column definitions (21 cols A–U) ──────────────────────────────────────
COLUMNS = [
    ("A", "Job Card No",                    14),
    ("B", "Date Opened",                    13),
    ("C", "Customer Name",                  20),
    ("D", "Mobile Number",                  17),
    ("E", "Vehicle Make & Model",           24),
    ("F", "Plate Number",                   14),
    ("G", "VIN / Chassis No",              21),
    ("H", "Mileage (KM)",                  13),
    ("I", "Customer Complaint",             32),
    ("J", "Diagnosis Details",              32),
    ("K", "Work Done / Repairs Performed",  36),
    ("L", "Parts Changed / Replaced",       32),
    ("M", "Current Status",                 24),
    ("N", "Ready for Collection",           18),
    ("O", "Technician Name",               19),
    ("P", "Service Advisor",               19),
    ("Q", "Last Update Date",              18),
    ("R", "WhatsApp Update Link",          21),
    ("S", "Customer Alert Sent",           18),
    ("T", "Remarks / Notes",               32),
    ("U", "Customer Email",                23),
]

STATUS_LIST = [
    "Vehicle Received", "Diagnosis Started", "Waiting Customer Approval",
    "Parts Ordered", "Repair In Progress", "Additional Work Required",
    "Work Finished", "Ready for Collection", "Delivered",
]

SAMPLE = [
    {
        "A": "FG-2024-001",
        "B": datetime.date.today() - datetime.timedelta(days=2),
        "C": "Ahmed Al Rashidi",
        "D": "+971501234567",
        "E": "Toyota Land Cruiser 2020",
        "F": "DXB-A-12345",
        "G": "JTMHX02J504012345",
        "H": 85000,
        "I": "Engine overheating. AC not cooling.",
        "J": "Coolant leak at upper hose. AC compressor low pressure.",
        "K": "Replaced upper radiator hose. Re-gassed AC system. Road test passed.",
        "L": "Upper radiator hose × 1, AC refrigerant R134a 800 g",
        "M": "Work Finished",
        "N": "YES",
        "O": "Mohammed Hassan",
        "P": "Khalid Mansoor",
        "Q": datetime.datetime.now(),
        "R": "",
        "S": "YES",
        "T": "Customer confirmed collection tomorrow morning.",
        "U": "ahmed.rashidi@email.com",
    },
    {
        "A": "FG-2024-002",
        "B": datetime.date.today() - datetime.timedelta(days=1),
        "C": "Sara Al Mansoori",
        "D": "+971502345678",
        "E": "Nissan Patrol 2019",
        "F": "AUH-B-67890",
        "G": "JN8AZ2KR5BT012345",
        "H": 120000,
        "I": "Gearbox slipping on 2nd gear. Whining noise at speed.",
        "J": "Gearbox oil burnt black. Solenoid pack faulty. Torque converter suspect.",
        "K": "",
        "L": "",
        "M": "Waiting Customer Approval",
        "N": "NO",
        "O": "Faisal Al Zaabi",
        "P": "Khalid Mansoor",
        "Q": datetime.datetime.now() - datetime.timedelta(days=1),
        "R": "",
        "S": "NO",
        "T": "Gearbox overhaul quote: AED 4,200. Awaiting customer go-ahead.",
        "U": "sara.mansoori@email.com",
    },
    {
        "A": "FG-2024-003",
        "B": datetime.date.today(),
        "C": "James Wilson",
        "D": "+971503456789",
        "E": "BMW X5 xDrive40i 2021",
        "F": "SHJ-C-11111",
        "G": "5UXKR6C56F0K12345",
        "H": 45000,
        "I": "Check engine light on. Steering vibration above 80 km/h.",
        "J": "Fault P0138: O2 sensor bank 1. Front brake discs warped.",
        "K": "Replacing front brake discs and pads. Fitting new O2 sensor bank 1.",
        "L": "Front brake discs × 2, front brake pads set, O2 sensor B1S2",
        "M": "Parts Ordered",
        "N": "NO",
        "O": "Mohammed Hassan",
        "P": "David Chen",
        "Q": datetime.datetime.now(),
        "R": "",
        "S": "YES",
        "T": "Parts ETA 2 business days. Customer informed via WhatsApp.",
        "U": "james.wilson@email.com",
    },
    {
        "A": "FG-2024-004",
        "B": datetime.date.today(),
        "C": "Fatima Al Zahra",
        "D": "+971504567890",
        "E": "Mercedes GLC 300 2022",
        "F": "DXB-D-55555",
        "G": "WDC0G8EB3KF123456",
        "H": 28000,
        "I": "Suspension noise over bumps. Tyre wear uneven.",
        "J": "Front left control arm bush worn. Wheel alignment out.",
        "K": "Replacing front left control arm. Full 4-wheel alignment.",
        "L": "Control arm assembly LH × 1",
        "M": "Repair In Progress",
        "N": "NO",
        "O": "Faisal Al Zaabi",
        "P": "David Chen",
        "Q": datetime.datetime.now(),
        "R": "",
        "S": "YES",
        "T": "ETA completion: today 5 PM.",
        "U": "fatima.alzahra@email.com",
    },
    {
        "A": "FG-2024-005",
        "B": datetime.date.today() - datetime.timedelta(days=3),
        "C": "Robert Khalil",
        "D": "+971505678901",
        "E": "Ford F-150 Raptor 2020",
        "F": "AJM-A-99999",
        "G": "1FTFW1RG9LFA12345",
        "H": 62000,
        "I": "Battery draining overnight. Starter motor slow.",
        "J": "Battery 3 years old, 60% capacity. Starter brushes worn.",
        "K": "Replaced battery and starter motor. Full electrical system check.",
        "L": "AGM battery 800CCA × 1, Starter motor × 1",
        "M": "Ready for Collection",
        "N": "YES",
        "O": "Mohammed Hassan",
        "P": "Khalid Mansoor",
        "Q": datetime.datetime.now(),
        "R": "",
        "S": "YES",
        "T": "Vehicle ready. Customer notified by WhatsApp and email.",
        "U": "robert.khalil@email.com",
    },
]


def wa_formula(row: int) -> str:
    """Google Sheets HYPERLINK+ENCODEURL formula for WhatsApp link."""
    g = "Falcon Garage"
    r = row
    return (
        f'=IF(A{r}="","",'
        f'HYPERLINK('
        f'"https://wa.me/"'
        f'&REGEXREPLACE(D{r},"[^0-9]","")'
        f'&"?text="'
        f'&ENCODEURL('
        f'"Dear "&C{r}&","'
        f'&CHAR(10)&CHAR(10)'
        f'&"Vehicle Update – {g}"'
        f'&CHAR(10)&CHAR(10)'
        f'&"Job Card: "&A{r}'
        f'&CHAR(10)'
        f'&"Vehicle: "&E{r}&" | "&F{r}'
        f'&CHAR(10)&CHAR(10)'
        f'&"Work Completed:"'
        f'&CHAR(10)&IF(K{r}="","N/A",K{r})'
        f'&CHAR(10)&CHAR(10)'
        f'&"Parts Changed:"'
        f'&CHAR(10)&IF(L{r}="","N/A",L{r})'
        f'&CHAR(10)&CHAR(10)'
        f'&"Current Status:"'
        f'&CHAR(10)&IF(M{r}="","Not Set",M{r})'
        f'&CHAR(10)&CHAR(10)'
        f'&"Ready for Collection:"'
        f'&CHAR(10)&IF(N{r}="","Pending",N{r})'
        f'&CHAR(10)&CHAR(10)'
        f'&"Thank you,"'
        f'&CHAR(10)&"{g}"'
        f'),'
        f'"📱 Send Update"))'
    )


def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JOB UPDATE TRACKER"
    ws.sheet_properties.tabColor = HEADER_BG

    # ── Row 1: Title ───────────────────────────────────────────────
    ws.row_dimensions[1].height = 52
    ws.merge_cells("A1:U1")
    c = ws["A1"]
    c.value = "🔧   FALCON GARAGE  |  JOB UPDATE TRACKER"
    c.font      = font(bold=True, color=WHITE, size=20)
    c.fill      = fill(TITLE_BG)
    c.alignment = align("center", "center")

    # ── Row 2: Subtitle ────────────────────────────────────────────
    ws.row_dimensions[2].height = 26
    ws.merge_cells("A2:U2")
    c = ws["A2"]
    c.value = (
        "Manage Job Cards  ·  Track Repairs  ·  WhatsApp & Email Updates  ·  "
        "Monthly PDF Reports  ·  Permanent History"
    )
    c.font      = font(italic=True, color="AACCEE", size=10)
    c.fill      = fill(HEADER_BG)
    c.alignment = align("center", "center")

    # ── Row 3: Column headers ──────────────────────────────────────
    ws.row_dimensions[3].height = 42
    hdr_border = Border(
        left   = Side(style="thin",   color=TITLE_BG),
        right  = Side(style="thin",   color=TITLE_BG),
        top    = Side(style="thin",   color=TITLE_BG),
        bottom = Side(style="medium", color=TITLE_BG),
    )
    for col_letter, label, width in COLUMNS:
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        c = ws.cell(row=3, column=col_idx, value=label)
        c.font      = font(bold=True, color=WHITE, size=10)
        c.fill      = fill(BAND_BG)
        c.alignment = align("center", "center", wrap=True)
        c.border    = hdr_border
        ws.column_dimensions[col_letter].width = width

    # ── Data validation (dropdowns) ────────────────────────────────
    status_dv = DataValidation(
        type="list",
        formula1=f'"{",".join(STATUS_LIST)}"',
        allow_blank=True, showDropDown=False,
    )
    status_dv.promptTitle = "Job Status"
    status_dv.prompt      = "Select the current repair status"
    status_dv.errorTitle  = "Invalid Status"
    status_dv.error       = "Please choose a value from the list"
    ws.add_data_validation(status_dv)
    status_dv.add("M4:M1000")

    for col_range in ("N4:N1000", "S4:S1000"):
        yn_dv = DataValidation(
            type="list", formula1='"YES,NO"',
            allow_blank=True, showDropDown=False,
        )
        yn_dv.errorTitle = "Invalid"
        yn_dv.error      = "Select YES or NO"
        ws.add_data_validation(yn_dv)
        yn_dv.add(col_range)

    # ── Conditional formatting (status row colours) ─────────────────
    cf_rules = [
        ("Waiting Customer Approval", RED_BG, RED_FG),
        ("Parts Ordered",             ORG_BG, ORG_FG),
        ("Additional Work Required",  ORG_BG, ORG_FG),
        ("Repair In Progress",        BLU_BG, BLU_FG),
        ("Work Finished",             GRN_BG, GRN_FG),
        ("Ready for Collection",      GRN_BG, GRN_FG),
        ("Delivered",                 GRY_BG, GRY_FG),
    ]
    for status, bg, fg in cf_rules:
        dxf = DifferentialStyle(
            font=Font(color=fg, bold=True),
            fill=PatternFill(start_color=bg, end_color=bg, fill_type="solid"),
        )
        rule = Rule(type="expression", formula=[f'$M4="{status}"'], dxf=dxf)
        ws.conditional_formatting.add("A4:U1000", rule)

    # ── Data rows ──────────────────────────────────────────────────
    cell_border = border()
    COL_LETTERS = [col for col, _, _ in COLUMNS]

    def write_row(ws, row_num, data: dict, is_sample=False):
        bg = ROW_A if (row_num - 4) % 2 == 0 else ROW_B
        ws.row_dimensions[row_num].height = 22 if not is_sample else 24

        for col_letter in COL_LETTERS:
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            c = ws.cell(row=row_num, column=col_idx)
            c.fill   = fill(bg)
            c.border = cell_border
            c.font   = font(size=10)
            c.alignment = align("left", "center")

        for col_letter, value in data.items():
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            c = ws.cell(row=row_num, column=col_idx)
            c.value = value

        # Column-specific styles
        # A: job card bold+accent
        a = ws.cell(row=row_num, column=1)
        a.font = font(bold=True, color=ACCENT, size=10)
        a.alignment = align("center", "center")

        # B: date format
        b = ws.cell(row=row_num, column=2)
        b.number_format = "DD-MMM-YYYY"
        b.alignment = align("center", "center")

        # C: customer name bold
        ws.cell(row=row_num, column=3).font = font(bold=True, size=10)

        # D: mobile center
        ws.cell(row=row_num, column=4).alignment = align("center", "center")

        # F: plate center
        ws.cell(row=row_num, column=6).alignment = align("center", "center")

        # H: mileage
        h = ws.cell(row=row_num, column=8)
        h.number_format = "#,##0"
        h.alignment = align("center", "center")

        # I, J, K, L, T: wrap text
        for col in (9, 10, 11, 12, 20):
            ws.cell(row=row_num, column=col).alignment = align("left", "top", wrap=True)
            ws.row_dimensions[row_num].height = 36 if is_sample else 22

        # M: status bold center
        m = ws.cell(row=row_num, column=13)
        m.font = font(bold=True, size=10)
        m.alignment = align("center", "center")

        # N: ready center bold
        n = ws.cell(row=row_num, column=14)
        n.font = font(bold=True, size=10)
        n.alignment = align("center", "center")

        # Q: update date
        q = ws.cell(row=row_num, column=17)
        q.number_format = "DD-MMM-YYYY HH:mm"
        q.alignment = align("center", "center")

        # R: WhatsApp formula
        r_cell = ws.cell(row=row_num, column=18)
        r_cell.value = wa_formula(row_num)
        r_cell.font  = font(bold=True, color=LINK_C, size=10, underline="single")
        r_cell.alignment = align("center", "center")

        # S: alert center
        ws.cell(row=row_num, column=19).alignment = align("center", "center")

        # U: email color
        u = ws.cell(row=row_num, column=21)
        u.font = font(color=ACCENT, size=10)
        u.alignment = align("center", "center")

    # Write 5 sample rows
    for idx, row_data in enumerate(SAMPLE):
        write_row(ws, 4 + idx, row_data, is_sample=True)

    # Write blank formatted rows 9–300 with WhatsApp formula
    for row_num in range(9, 301):
        write_row(ws, row_num, {}, is_sample=False)

    # ── Freeze panes at row 4 (title + subtitle + header frozen) ──
    ws.freeze_panes = "A4"

    # ── Auto-filter on header row ──────────────────────────────────
    ws.auto_filter.ref = "A3:U3"

    # ── Page setup ─────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:3"
    ws.sheet_view.showGridLines = True

    wb.save(OUTPUT)
    print(f"✅  Saved: {OUTPUT}")


if __name__ == "__main__":
    build()
