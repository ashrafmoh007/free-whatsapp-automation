/**
 * ================================================================
 *  FALCON GARAGE — JOB UPDATE TRACKER  v3.0
 *  Google Sheets + Apps Script  |  فالكون للسيارات — قطر
 * ================================================================
 *
 *  QUICK SETUP (one-time, ~60 seconds):
 *  ─────────────────────────────────────
 *  1. Open a NEW blank Google Sheet
 *  2. Extensions → Apps Script → paste this file → Save
 *  3. Back in the sheet: Falcon Garage → ⚙ Setup Sheet
 *  4. Authorise when prompted — done!
 *
 *  HOW TO USE:
 *  ─────────────────────────────────────
 *  • Add a job row → Status / Priority dropdowns auto-appear
 *  • Change Status → row colour updates instantly (onEdit)
 *  • Click 📱 WhatsApp link in column AG to open WA Web
 *  • Menu → New Job Card → auto-generates next JC-YYYY-NNN ID
 *  • Menu → Email → send branded HTML alerts
 *  • Menu → Monthly PDF → export to Google Drive
 *
 *  COLUMNS A–AH  (34 total):
 *  ─────────────────────────────────────
 *  A  Job ID          B  Date In         C  Time In
 *  D  Status          E  Priority        F  Job Type        G  Technician
 *  H  Cust Name       I  Phone           J  Email           K  Cust ID
 *  L  Plate           M  Make            N  Model           O  Year
 *  P  VIN             Q  Odometer        R  Veh Type
 *  S  Complaint       T  Work Done       U  Parts Used
 *  V  Parts QAR       W  Labour QAR      X  Discount QAR    Y  VAT QAR
 *  Z  Total QAR       AA Payment Method  AB Pay Status
 *  AC Est Ref         AD Inv Ref         AE Date Out        AF Approved By
 *  AG WA Link         AH Alert Sent
 * ================================================================
 */

// ──────────────────────────────────────────────────────────────
//  CONFIGURATION
// ──────────────────────────────────────────────────────────────
const CFG = {
  GARAGE_NAME : "Falcon Garage",
  GARAGE_AR   : "ورشة فالكون",
  GARAGE_TEL  : "+974 4486 2018 | 7072 5811 | 7003 3723 | 7731 8043",
  GARAGE_WEB  : "www.falqatar.com",
  GARAGE_EMAIL: "info@falqatar.com",
  GARAGE_ADDR : "Street 47, Bldg 190, Zone 57, Industrial Area, Doha - Qatar",

  SHEET_NAME  : "JOB UPDATE TRACKER",
  TITLE_ROW   : 1,
  SUB_ROW     : 2,
  GROUP_ROW   : 3,
  HEADER_ROW  : 4,
  DATA_START  : 5,
  MAX_ROWS    : 500,
  JOB_PREFIX  : "JC",
  TOTAL_COLS  : 34,
  CURRENCY    : "QAR",
  VAT_RATE    : 0.05,
};

// Column index map (1-based)
const C = {
  JOB_ID    : 1,  DATE_IN   : 2,  TIME_IN   : 3,
  STATUS    : 4,  PRIORITY  : 5,  JOB_TYPE  : 6,  TECHNICIAN: 7,
  CUST_NAME : 8,  PHONE     : 9,  EMAIL     : 10, CUST_ID   : 11,
  PLATE     : 12, MAKE      : 13, MODEL     : 14, YEAR      : 15,
  VIN       : 16, ODOMETER  : 17, VEH_TYPE  : 18,
  COMPLAINT : 19, WORK_DONE : 20, PARTS_USED: 21,
  PARTS_QAR : 22, LABOUR_QAR: 23, DISCOUNT  : 24, VAT_QAR   : 25,
  TOTAL_QAR : 26, PAYMENT   : 27, PAY_STATUS: 28,
  EST_REF   : 29, INV_REF   : 30, DATE_OUT  : 31, APPROVED_BY: 32,
  WA_LINK   : 33, ALERT_SENT: 34,
};

const STATUS_LIST   = ["🔴 Open","🔵 In Progress","🔧 Body Shop","🔍 QC","🟢 Ready","✅ Delivered","⏸ On Hold"];
const PRIORITY_LIST = ["🔴 URGENT","🟡 HIGH","🟢 NORMAL","🔵 LOW"];
const JOB_TYPE_LIST = ["General Service","Oil Change","Body Repair","Electrical","AC Repair","Tyres / Alignment","Inspection","Custom / Other"];
const VEH_TYPE_LIST = ["Sedan","SUV","Pick-up","Van","Truck","Motorcycle","Other"];
const PAY_LIST      = ["Cash","Card","Bank Transfer","Credit","Pending"];

// Section group definitions
const GROUPS = [
  { label:"🔧  JOB INFORMATION",               s:1,  e:7,  bg:"#1C2839", fg:"#F59E0B" },
  { label:"👤  CUSTOMER  ·  العميل",            s:8,  e:11, bg:"#0D47A1", fg:"#FFFFFF" },
  { label:"🚗  VEHICLE  ·  المركبة",            s:12, e:18, bg:"#0891B2", fg:"#FFFFFF" },
  { label:"🔧  COMPLAINT & WORK  ·  الشكوى",   s:19, e:21, bg:"#7C3AED", fg:"#FFFFFF" },
  { label:"💰  FINANCIALS  ·  المالية (QAR)",  s:22, e:28, bg:"#065F46", fg:"#F59E0B" },
  { label:"📋  ADMIN",                          s:29, e:32, bg:"#374151", fg:"#FFFFFF" },
  { label:"📱  COMMUNICATION",                  s:33, e:34, bg:"#166534", fg:"#FFFFFF" },
];

// Column header labels (index 0 = col A)
const HEADERS = [
  "Job ID","Date In","Time In",
  "Status","Priority","Job Type","Technician",
  "Customer Name","Phone","Email","Customer ID",
  "Plate No.","Make","Model","Year","VIN","Odometer","Veh Type",
  "Complaint / Fault","Work Done","Parts Used",
  "Parts (QAR)","Labour (QAR)","Discount (QAR)","VAT (QAR)","Total (QAR)",
  "Payment Method","Pay Status",
  "Estimate Ref","Invoice Ref","Date Out","Approved By",
  "WhatsApp","Alert Sent",
];

// Status colour map  { keyword → { bg, fg } }
const STATUS_COLOURS = {
  "Open"       : { bg:"#FF4444", fg:"#FFFFFF" },
  "In Progress": { bg:"#1565C0", fg:"#FFFFFF" },
  "Body Shop"  : { bg:"#5D4037", fg:"#FFFFFF" },
  "QC"         : { bg:"#6A1B9A", fg:"#FFFFFF" },
  "Ready"      : { bg:"#2E7D32", fg:"#FFFFFF" },
  "Delivered"  : { bg:"#004D40", fg:"#CCFFCC" },
  "On Hold"    : { bg:"#37474F", fg:"#CFD8DC" },
};


// ──────────────────────────────────────────────────────────────
//  MENU
// ──────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🦅 Falcon Garage")
    .addItem("⚙ Setup Sheet (first-time)",  "setupSheet")
    .addSeparator()
    .addItem("➕ New Job Card",              "createNewJobCard")
    .addItem("📱 Open WhatsApp (active row)","openWhatsAppForActiveRow")
    .addSeparator()
    .addItem("📧 Email Customer (active row)","emailCustomerUpdate")
    .addItem("📧 Email All Active Jobs",      "emailAllActiveJobs")
    .addItem("📧 Email Ready-for-Pickup",     "emailReadyAlerts")
    .addSeparator()
    .addItem("📄 Generate Monthly PDF",       "generateMonthlyPDF")
    .addSeparator()
    .addItem("🔍 Search Jobs",               "searchJobs")
    .addItem("🔄 Clear Filters",             "clearFilters")
    .addItem("🎨 Refresh Row Colours",       "refreshRowColors")
    .addItem("🔗 Rebuild Formulas",          "rebuildFormulas")
    .addToUi();
}


// ──────────────────────────────────────────────────────────────
//  SETUP
// ──────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CFG.SHEET_NAME);
    const idx = sh.getIndex();
    if (idx !== 1) ss.setActiveSheet(sh), ss.moveActiveSheet(1);
  }

  sh.clearContents();
  sh.clearFormats();

  _buildTitle(sh);
  _buildGroupHeaders(sh);
  _buildHeaders(sh);
  _applyColumnWidths(sh);
  _setupValidation(sh);
  _setupConditionalFormatting(sh);
  _freezeAndFilter(sh);
  _styleDataBand(sh);
  _addSampleData(sh);
  _buildFormulas(sh);

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Falcon Garage Tracker v3.0 ready!", "Setup Complete", 5);
}


// ──────────────────────────────────────────────────────────────
//  TITLE  (rows 1–2)
// ──────────────────────────────────────────────────────────────
function _buildTitle(sh) {
  const nc = CFG.TOTAL_COLS;

  // Row 1 — main title
  sh.getRange(1, 1, 1, nc).merge()
    .setValue(`🦅  ${CFG.GARAGE_NAME}  |  ${CFG.GARAGE_AR}  —  Job Update Tracker`)
    .setBackground("#111827")
    .setFontColor("#F59E0B")
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 48);

  // Row 2 — subtitle / contact bar
  sh.getRange(2, 1, 1, nc).merge()
    .setValue(`📞 ${CFG.GARAGE_TEL}   🌐 ${CFG.GARAGE_WEB}   ✉ ${CFG.GARAGE_EMAIL}   📍 ${CFG.GARAGE_ADDR}`)
    .setBackground("#1F2937")
    .setFontColor("#9CA3AF")
    .setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 28);
}


// ──────────────────────────────────────────────────────────────
//  GROUP HEADERS  (row 3)
// ──────────────────────────────────────────────────────────────
function _buildGroupHeaders(sh) {
  sh.setRowHeight(CFG.GROUP_ROW, 26);
  GROUPS.forEach(g => {
    const r = sh.getRange(CFG.GROUP_ROW, g.s, 1, g.e - g.s + 1);
    if (g.e > g.s) r.merge();
    r.setValue(g.label)
      .setBackground(g.bg).setFontColor(g.fg)
      .setFontSize(10).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
}


// ──────────────────────────────────────────────────────────────
//  COLUMN HEADERS  (row 4)
// ──────────────────────────────────────────────────────────────
function _buildHeaders(sh) {
  sh.setRowHeight(CFG.HEADER_ROW, 40);
  HEADERS.forEach((h, i) => {
    sh.getRange(CFG.HEADER_ROW, i + 1)
      .setValue(h)
      .setBackground("#0F172A")
      .setFontColor("#E2E8F0")
      .setFontSize(9).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrap(true);
  });
}


// ──────────────────────────────────────────────────────────────
//  COLUMN WIDTHS
// ──────────────────────────────────────────────────────────────
function _applyColumnWidths(sh) {
  const WIDTHS = {
    1:100, 2:90,  3:70,  4:110, 5:95,  6:110, 7:100,
    8:130, 9:120, 10:150,11:90,
    12:90, 13:90, 14:90, 15:55, 16:130,17:80, 18:80,
    19:200,20:200,21:150,
    22:90, 23:90, 24:80, 25:80, 26:90, 27:110,28:90,
    29:90, 30:90, 31:90, 32:100,
    33:130,34:90,
  };
  Object.entries(WIDTHS).forEach(([col, w]) => sh.setColumnWidth(Number(col), w));
}


// ──────────────────────────────────────────────────────────────
//  DATA VALIDATION  (dropdowns)
// ──────────────────────────────────────────────────────────────
function _setupValidation(sh) {
  const last = CFG.DATA_START + CFG.MAX_ROWS - 1;

  const dv = (list) =>
    SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(false)
      .build();

  sh.getRange(CFG.DATA_START, C.STATUS,    CFG.MAX_ROWS, 1).setDataValidation(dv(STATUS_LIST));
  sh.getRange(CFG.DATA_START, C.PRIORITY,  CFG.MAX_ROWS, 1).setDataValidation(dv(PRIORITY_LIST));
  sh.getRange(CFG.DATA_START, C.JOB_TYPE,  CFG.MAX_ROWS, 1).setDataValidation(dv(JOB_TYPE_LIST));
  sh.getRange(CFG.DATA_START, C.VEH_TYPE,  CFG.MAX_ROWS, 1).setDataValidation(dv(VEH_TYPE_LIST));
  sh.getRange(CFG.DATA_START, C.PAY_STATUS,CFG.MAX_ROWS, 1).setDataValidation(dv(PAY_LIST));
}


// ──────────────────────────────────────────────────────────────
//  CONDITIONAL FORMATTING  (row-level colours by status)
// ──────────────────────────────────────────────────────────────
function _setupConditionalFormatting(sh) {
  const dataRange = sh.getRange(CFG.DATA_START, 1, CFG.MAX_ROWS, CFG.TOTAL_COLS);
  const statusCol = _colLetter(C.STATUS);
  const rules = [];

  const statuses = [
    ["Open",        "#FF4444","#FFFFFF"],
    ["In Progress", "#1565C0","#FFFFFF"],
    ["Body Shop",   "#5D4037","#FFFFFF"],
    ["QC",          "#6A1B9A","#FFFFFF"],
    ["Ready",       "#2E7D32","#FFFFFF"],
    ["Delivered",   "#004D40","#CCFFCC"],
    ["On Hold",     "#37474F","#CFD8DC"],
  ];

  statuses.forEach(([kw, bg, fg]) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=NOT(ISBLANK($${statusCol}${CFG.DATA_START}))*ISNUMBER(SEARCH("${kw}",$${statusCol}${CFG.DATA_START}))`)
        .setBackground(bg).setFontColor(fg)
        .setRanges([dataRange])
        .build()
    );
  });

  sh.setConditionalFormatRules(rules);
}


// ──────────────────────────────────────────────────────────────
//  FREEZE + FILTER
// ──────────────────────────────────────────────────────────────
function _freezeAndFilter(sh) {
  sh.setFrozenRows(CFG.HEADER_ROW);
  sh.setFrozenColumns(1);
  sh.getRange(CFG.HEADER_ROW, 1, 1, CFG.TOTAL_COLS).createFilter();
}


// ──────────────────────────────────────────────────────────────
//  ALTERNATING ROW BAND  (zebra stripes for empty data rows)
// ──────────────────────────────────────────────────────────────
function _styleDataBand(sh) {
  for (let r = CFG.DATA_START; r < CFG.DATA_START + CFG.MAX_ROWS; r++) {
    const bg = (r % 2 === 0) ? "#1A2332" : "#111827";
    sh.getRange(r, 1, 1, CFG.TOTAL_COLS).setBackground(bg).setFontColor("#D1D5DB").setFontSize(9);
  }
}


// ──────────────────────────────────────────────────────────────
//  FORMULA HELPERS
// ──────────────────────────────────────────────────────────────
function _vatFormula(row) {
  const v = _colLetter(C.PARTS_QAR);
  const w = _colLetter(C.LABOUR_QAR);
  const x = _colLetter(C.DISCOUNT);
  return `=IF(${v}${row}+${w}${row}>0,ROUND(MAX(0,${v}${row}+${w}${row}-${x}${row})*0.05,2),0)`;
}

function _totalFormula(row) {
  const v = _colLetter(C.PARTS_QAR);
  const w = _colLetter(C.LABOUR_QAR);
  const x = _colLetter(C.DISCOUNT);
  const y = _colLetter(C.VAT_QAR);
  return `=IF(${v}${row}+${w}${row}>0,ROUND(MAX(0,${v}${row}+${w}${row}-${x}${row})+${y}${row},2),0)`;
}

function _waFormula(row) {
  const ph = _colLetter(C.PHONE);
  const jc = _colLetter(C.JOB_ID);
  const nm = _colLetter(C.CUST_NAME);
  const mk = _colLetter(C.MAKE);
  const md = _colLetter(C.MODEL);
  const pl = _colLetter(C.PLATE);
  const st = _colLetter(C.STATUS);
  const tot= _colLetter(C.TOTAL_QAR);
  return `=IF(${ph}${row}="","",HYPERLINK("https://wa.me/974"&REGEXREPLACE(${ph}${row},"[^0-9]","")&"?text="&ENCODEURL("Dear "${nm}${row}", your vehicle ("${mk}${row}" "${md}${row}" | Plate: "${pl}${row}") — Job: "${jc}${row}" | Status: "${st}${row}" | Total: QAR "${tot}${row}". Thank you — Falcon Garage, Doha +974 4486 2018"),"📱 WhatsApp"))`;
}


// ──────────────────────────────────────────────────────────────
//  BUILD FORMULAS FOR ALL EXISTING DATA ROWS
// ──────────────────────────────────────────────────────────────
function _buildFormulas(sh) {
  const lastRow = Math.max(sh.getLastRow(), CFG.DATA_START - 1);
  if (lastRow < CFG.DATA_START) return;
  for (let r = CFG.DATA_START; r <= lastRow; r++) {
    sh.getRange(r, C.VAT_QAR  ).setFormula(_vatFormula(r));
    sh.getRange(r, C.TOTAL_QAR).setFormula(_totalFormula(r));
    sh.getRange(r, C.WA_LINK  ).setFormula(_waFormula(r));
  }
}


// ──────────────────────────────────────────────────────────────
//  SAMPLE DATA  (5 rows from JOB DATABASE)
// ──────────────────────────────────────────────────────────────
function _addSampleData(sh) {
  const today = new Date();
  const fmt = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const rows = [
    ["JC-2026-001", fmt(today), "08:30", "✅ Delivered", "🟢 NORMAL", "Oil Change",      "Ahmed Al-Rashidi", "Mohammed Al-Hamad", "+974 5512 3456", "m.alhamad@email.com", "QID-28843721", "A 12345","Toyota","Land Cruiser","2022","JTMCV02J204045321","62000 km","SUV",         "Routine oil change & filter","Completed oil change, replaced filter","Engine oil 5L, oil filter",         350, 150, 0, "", "", "Cash",            "Pending", "EST-001", "INV-001", fmt(today), "Ahmed Al-Rashidi", "", ""],
    ["JC-2026-002", fmt(today), "09:15", "🟢 Ready",     "🔴 URGENT", "AC Repair",       "Khalid Al-Sayed",  "Sarah Johnson",     "+974 6623 4567", "s.johnson@gmail.com","QID-39124568", "B 67890","Nissan","Patrol","2020","JN8AY2ND5L9760412", "45200 km","SUV",         "AC not cooling at all",     "Replaced compressor & recharged","AC Compressor, refrigerant gas",    1200, 450, 0, "", "", "Card",            "Pending", "EST-002", "INV-002", "",        "Khalid Al-Sayed",  "", ""],
    ["JC-2026-003", fmt(today), "10:00", "🔵 In Progress","🟡 HIGH",  "Body Repair",     "Yousuf Ibrahim",   "Ali Karimi",        "+974 7734 5678", "ali.karimi@work.qa", "QID-45231890", "C 11111","BMW","5 Series","2021","WBA13BJ08MCF62831", "31500 km","Sedan",       "Front bumper damage",       "Dent removal in progress",       "Bumper paint, filler compound",     800, 600, 50,"", "", "Bank Transfer",    "Pending", "EST-003", "",        "",        "Yousuf Ibrahim",   "", ""],
    ["JC-2026-004", fmt(today), "11:30", "🔴 Open",      "🟢 NORMAL", "General Service", "Fatima Al-Dosari", "Rania Khalil",      "+974 3345 6789", "rania.k@hotmail.com","QID-56342901", "D 22222","Mercedes","C-Class","2023","WDD2050422R456789", "15000 km","Sedan",       "Full service check",        "Awaiting technician assignment", "Service kit, brake pads",           500, 300, 25,"", "", "Cash",            "Pending", "EST-004", "",        "",        "",                 "", ""],
    ["JC-2026-005", fmt(today), "13:00", "🔧 Body Shop",  "🟡 HIGH",  "Tyres / Alignment","Hassan Al-Mansoor","James O'Brien",     "+974 5567 8901", "james.ob@corp.net",  "QID-67453012", "E 33333","Ford","F-150","2019","1FTEW1E55KKC24680", "78900 km","Pick-up",      "Wheel alignment & new tyres","Alignment done, fitting tyres",   "4x Michelin LTX tyres",             1800, 250, 100,"","", "Credit",          "Pending", "EST-005", "",        "",        "Hassan Al-Mansoor","", ""],
  ];

  rows.forEach((row, i) => {
    const r = CFG.DATA_START + i;
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  });

  // Apply formulas over sample rows
  for (let r = CFG.DATA_START; r < CFG.DATA_START + rows.length; r++) {
    sh.getRange(r, C.VAT_QAR  ).setFormula(_vatFormula(r));
    sh.getRange(r, C.TOTAL_QAR).setFormula(_totalFormula(r));
    sh.getRange(r, C.WA_LINK  ).setFormula(_waFormula(r));
  }
}


// ──────────────────────────────────────────────────────────────
//  onEdit  — live row colouring + stamp update date
// ──────────────────────────────────────────────────────────────
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET_NAME) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < CFG.DATA_START) return;

  if (col === C.STATUS) {
    const val = (e.value || "").toString();
    let matched = false;
    for (const [kw, colours] of Object.entries(STATUS_COLOURS)) {
      if (val.includes(kw)) {
        sh.getRange(row, 1, 1, CFG.TOTAL_COLS)
          .setBackground(colours.bg).setFontColor(colours.fg);
        matched = true;
        break;
      }
    }
    if (!matched) {
      const bg = (row % 2 === 0) ? "#1A2332" : "#111827";
      sh.getRange(row, 1, 1, CFG.TOTAL_COLS).setBackground(bg).setFontColor("#D1D5DB");
    }
    // stamp AE (Date Out) when Delivered
    if (val.includes("Delivered") && !sh.getRange(row, C.DATE_OUT).getValue()) {
      sh.getRange(row, C.DATE_OUT).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
    }
  }

  // Auto-stamp Date In / Time In on new Job ID entry
  if (col === C.JOB_ID && e.value) {
    if (!sh.getRange(row, C.DATE_IN).getValue())
      sh.getRange(row, C.DATE_IN).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
    if (!sh.getRange(row, C.TIME_IN).getValue())
      sh.getRange(row, C.TIME_IN).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"));
  }

  // Ensure formulas present when financial cols or phone is edited
  if ([C.PARTS_QAR, C.LABOUR_QAR, C.DISCOUNT, C.PHONE].includes(col)) {
    sh.getRange(row, C.VAT_QAR  ).setFormula(_vatFormula(row));
    sh.getRange(row, C.TOTAL_QAR).setFormula(_totalFormula(row));
    sh.getRange(row, C.WA_LINK  ).setFormula(_waFormula(row));
  }
}


// ──────────────────────────────────────────────────────────────
//  CREATE NEW JOB CARD
// ──────────────────────────────────────────────────────────────
function createNewJobCard() {
  const sh = _getSheet();
  const newId = _nextJobId(sh);
  const lastRow = sh.getLastRow();
  const newRow = Math.max(lastRow + 1, CFG.DATA_START);
  const now = new Date();
  const tz  = Session.getScriptTimeZone();

  sh.getRange(newRow, C.JOB_ID ).setValue(newId);
  sh.getRange(newRow, C.DATE_IN).setValue(Utilities.formatDate(now, tz, "dd/MM/yyyy"));
  sh.getRange(newRow, C.TIME_IN).setValue(Utilities.formatDate(now, tz, "HH:mm"));
  sh.getRange(newRow, C.STATUS ).setValue("🔴 Open");
  sh.getRange(newRow, C.PRIORITY).setValue("🟢 NORMAL");
  sh.getRange(newRow, C.VAT_QAR  ).setFormula(_vatFormula(newRow));
  sh.getRange(newRow, C.TOTAL_QAR).setFormula(_totalFormula(newRow));
  sh.getRange(newRow, C.WA_LINK  ).setFormula(_waFormula(newRow));

  // style new row
  sh.getRange(newRow, 1, 1, CFG.TOTAL_COLS)
    .setBackground("#FF4444").setFontColor("#FFFFFF").setFontSize(9);

  sh.setActiveRange(sh.getRange(newRow, C.CUST_NAME));
  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ New job created: ${newId}`, "New Job", 4);
}

function _nextJobId(sh) {
  const year  = new Date().getFullYear();
  const prefix = `${CFG.JOB_PREFIX}-${year}-`;
  const data  = sh.getRange(CFG.DATA_START, C.JOB_ID, sh.getLastRow() - CFG.DATA_START + 1, 1).getValues();
  let max = 0;
  data.forEach(([v]) => {
    if (typeof v === "string" && v.startsWith(prefix)) {
      const n = parseInt(v.replace(prefix, ""), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return `${prefix}${String(max + 1).padStart(3, "0")}`;
}


// ──────────────────────────────────────────────────────────────
//  WHATSAPP
// ──────────────────────────────────────────────────────────────
function openWhatsAppForActiveRow() {
  const sh  = _getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < CFG.DATA_START) { SpreadsheetApp.getUi().alert("Please select a data row first."); return; }
  const url = sh.getRange(row, C.WA_LINK).getValue();
  if (!url) { SpreadsheetApp.getUi().alert("No WhatsApp link — check that Phone is filled in."); return; }
  const html = HtmlService.createHtmlOutput(`<script>window.open("${url}","_blank");google.script.host.close();</script>`)
    .setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening WhatsApp…");
  sh.getRange(row, C.ALERT_SENT).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
}

function stampUpdateDate() {
  const sh  = _getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < CFG.DATA_START) return;
  sh.getRange(row, C.ALERT_SENT).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
}


// ──────────────────────────────────────────────────────────────
//  SEARCH / FILTER
// ──────────────────────────────────────────────────────────────
function searchJobs() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt("🔍 Search Jobs", "Enter job ID, customer name, or plate number:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const term = resp.getResponseText().trim().toLowerCase();
  if (!term) return;

  const sh    = _getSheet();
  const last  = sh.getLastRow();
  if (last < CFG.DATA_START) return;

  const data = sh.getRange(CFG.DATA_START, 1, last - CFG.DATA_START + 1, CFG.TOTAL_COLS).getValues();
  const hits = [];
  data.forEach((row, i) => {
    const searchable = [row[C.JOB_ID-1], row[C.CUST_NAME-1], row[C.PLATE-1], row[C.PHONE-1]].join(" ").toLowerCase();
    if (searchable.includes(term)) hits.push(CFG.DATA_START + i);
  });

  if (!hits.length) { ui.alert(`No results for "${term}".`); return; }
  sh.setActiveRange(sh.getRange(hits[0], 1));
  SpreadsheetApp.getActiveSpreadsheet().toast(`Found ${hits.length} result(s) — first match highlighted.`, "Search", 4);
}

function clearFilters() {
  const sh = _getSheet();
  const f  = sh.getFilter();
  if (f) f.remove();
  sh.getRange(CFG.HEADER_ROW, 1, 1, CFG.TOTAL_COLS).createFilter();
  SpreadsheetApp.getActiveSpreadsheet().toast("Filters cleared.", "Done", 3);
}

function refreshRowColors() {
  const sh   = _getSheet();
  const last = sh.getLastRow();
  if (last < CFG.DATA_START) return;
  const statuses = sh.getRange(CFG.DATA_START, C.STATUS, last - CFG.DATA_START + 1, 1).getValues();
  statuses.forEach(([v], i) => {
    const row = CFG.DATA_START + i;
    let matched = false;
    for (const [kw, colours] of Object.entries(STATUS_COLOURS)) {
      if (String(v).includes(kw)) {
        sh.getRange(row, 1, 1, CFG.TOTAL_COLS).setBackground(colours.bg).setFontColor(colours.fg);
        matched = true; break;
      }
    }
    if (!matched) {
      const bg = (row % 2 === 0) ? "#1A2332" : "#111827";
      sh.getRange(row, 1, 1, CFG.TOTAL_COLS).setBackground(bg).setFontColor("#D1D5DB");
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Row colours refreshed.", "Done", 3);
}

function rebuildFormulas() {
  const sh   = _getSheet();
  const last = sh.getLastRow();
  if (last < CFG.DATA_START) return;
  for (let r = CFG.DATA_START; r <= last; r++) {
    sh.getRange(r, C.VAT_QAR  ).setFormula(_vatFormula(r));
    sh.getRange(r, C.TOTAL_QAR).setFormula(_totalFormula(r));
    sh.getRange(r, C.WA_LINK  ).setFormula(_waFormula(r));
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Formulas rebuilt.", "Done", 3);
}


// ──────────────────────────────────────────────────────────────
//  EMAIL — SINGLE ROW
// ──────────────────────────────────────────────────────────────
function emailCustomerUpdate() {
  const sh  = _getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < CFG.DATA_START) { SpreadsheetApp.getUi().alert("Please select a data row first."); return; }

  const vals = sh.getRange(row, 1, 1, CFG.TOTAL_COLS).getValues()[0];
  const email = vals[C.EMAIL - 1];
  if (!_isValidEmail(email)) { SpreadsheetApp.getUi().alert("No valid email address in this row."); return; }

  MailApp.sendEmail({ to: email, subject: _buildEmail(vals).subject, htmlBody: _buildEmail(vals).html });
  sh.getRange(row, C.ALERT_SENT).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
  SpreadsheetApp.getActiveSpreadsheet().toast(`Email sent to ${email}`, "Sent", 4);
}


// ──────────────────────────────────────────────────────────────
//  EMAIL — BULK
// ──────────────────────────────────────────────────────────────
function emailAllActiveJobs() {
  _bulkEmail(["🔴 Open","🔵 In Progress","🔧 Body Shop","🔍 QC"]);
}

function emailReadyAlerts() {
  _bulkEmail(["🟢 Ready"]);
}

function _bulkEmail(statusFilters) {
  const sh   = _getSheet();
  const last = sh.getLastRow();
  if (last < CFG.DATA_START) return;
  const data = sh.getRange(CFG.DATA_START, 1, last - CFG.DATA_START + 1, CFG.TOTAL_COLS).getValues();
  let sent = 0, skipped = 0;
  data.forEach((row, i) => {
    const status = String(row[C.STATUS - 1]);
    const email  = row[C.EMAIL - 1];
    const match  = statusFilters.some(f => status.includes(f.replace(/^[^\s]+ /,"")));
    if (!match || !_isValidEmail(email)) { skipped++; return; }
    const built = _buildEmail(row);
    MailApp.sendEmail({ to: email, subject: built.subject, htmlBody: built.html });
    sh.getRange(CFG.DATA_START + i, C.ALERT_SENT)
      .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"));
    sent++;
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(`Sent: ${sent}  |  Skipped: ${skipped}`, "Bulk Email Done", 5);
}


// ──────────────────────────────────────────────────────────────
//  EMAIL TEMPLATE BUILDER
// ──────────────────────────────────────────────────────────────
function _buildEmail(row) {
  const jobId    = row[C.JOB_ID     - 1] || "—";
  const custName = row[C.CUST_NAME  - 1] || "Valued Customer";
  const make     = row[C.MAKE       - 1] || "";
  const model    = row[C.MODEL      - 1] || "";
  const plate    = row[C.PLATE      - 1] || "";
  const status   = row[C.STATUS     - 1] || "";
  const workDone = row[C.WORK_DONE  - 1] || "—";
  const parts    = row[C.PARTS_QAR  - 1] || 0;
  const labour   = row[C.LABOUR_QAR - 1] || 0;
  const disc     = row[C.DISCOUNT   - 1] || 0;
  const vat      = row[C.VAT_QAR    - 1] || 0;
  const total    = row[C.TOTAL_QAR  - 1] || 0;
  const tech     = row[C.TECHNICIAN - 1] || "—";

  const statusLabel = String(status).replace(/^[^\s]+ /,"");
  const subject = `[${jobId}] Vehicle Update — ${statusLabel} | Falcon Garage`;

  const html = `
<!DOCTYPE html><html><body style="margin:0;padding:0;background:#0F172A;font-family:Arial,sans-serif;">
<table width="600" align="center" style="background:#1E293B;border-radius:10px;overflow:hidden;margin:20px auto;">
  <tr><td style="background:#111827;padding:24px;text-align:center;">
    <h1 style="color:#F59E0B;margin:0;font-size:22px;">🦅 Falcon Garage</h1>
    <p style="color:#9CA3AF;margin:4px 0;font-size:12px;">ورشة فالكون | Doha, Qatar</p>
  </td></tr>
  <tr><td style="padding:24px;">
    <p style="color:#E2E8F0;font-size:15px;">Dear <strong>${custName}</strong>,</p>
    <p style="color:#9CA3AF;font-size:13px;">Here is an update for your vehicle currently at our workshop.</p>
    <table width="100%" style="border-collapse:collapse;margin:16px 0;">
      <tr style="background:#0D47A1;"><td colspan="2" style="padding:8px 12px;color:#fff;font-weight:bold;font-size:12px;">🔧 JOB DETAILS</td></tr>
      <tr style="background:#1A2332;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;width:40%;">Job ID</td><td style="padding:8px 12px;color:#F59E0B;font-weight:bold;font-size:12px;">${jobId}</td></tr>
      <tr style="background:#111827;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Vehicle</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${make} ${model} — ${plate}</td></tr>
      <tr style="background:#1A2332;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Status</td><td style="padding:8px 12px;color:#4ADE80;font-weight:bold;font-size:12px;">${status}</td></tr>
      <tr style="background:#111827;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Work Done</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${workDone}</td></tr>
      <tr style="background:#1A2332;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Technician</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${tech}</td></tr>
    </table>
    <table width="100%" style="border-collapse:collapse;margin:16px 0;">
      <tr style="background:#065F46;"><td colspan="2" style="padding:8px 12px;color:#F59E0B;font-weight:bold;font-size:12px;">💰 FINANCIAL SUMMARY (QAR)</td></tr>
      <tr style="background:#1A2332;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;width:40%;">Parts</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${_fmtQAR(parts)}</td></tr>
      <tr style="background:#111827;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Labour</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${_fmtQAR(labour)}</td></tr>
      <tr style="background:#1A2332;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">Discount</td><td style="padding:8px 12px;color:#FCA5A5;font-size:12px;">- ${_fmtQAR(disc)}</td></tr>
      <tr style="background:#111827;"><td style="padding:8px 12px;color:#9CA3AF;font-size:12px;">VAT (5%)</td><td style="padding:8px 12px;color:#E2E8F0;font-size:12px;">${_fmtQAR(vat)}</td></tr>
      <tr style="background:#065F46;"><td style="padding:10px 12px;color:#F59E0B;font-weight:bold;font-size:13px;">TOTAL</td><td style="padding:10px 12px;color:#F59E0B;font-weight:bold;font-size:15px;">${_fmtQAR(total)}</td></tr>
    </table>
    <p style="color:#9CA3AF;font-size:11px;margin-top:20px;">For enquiries: <a href="tel:+97444862018" style="color:#F59E0B;">+974 4486 2018</a> | <a href="mailto:${CFG.GARAGE_EMAIL}" style="color:#F59E0B;">${CFG.GARAGE_EMAIL}</a></p>
  </td></tr>
  <tr><td style="background:#0F172A;padding:16px;text-align:center;">
    <p style="color:#6B7280;font-size:10px;margin:0;">${CFG.GARAGE_NAME} | ${CFG.GARAGE_ADDR}</p>
    <p style="color:#6B7280;font-size:10px;margin:4px 0;">${CFG.GARAGE_WEB}</p>
  </td></tr>
</table></body></html>`;

  return { subject, html };
}


// ──────────────────────────────────────────────────────────────
//  MONTHLY PDF
// ──────────────────────────────────────────────────────────────
function generateMonthlyPDF() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "📄 Monthly PDF",
    "Enter month & year (e.g. May 2026)  — leave blank for current month:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  _runMonthlyPDF(input);
}

function _runMonthlyPDF(monthInput) {
  const now    = new Date();
  let   target = now;
  if (monthInput) {
    const parts = monthInput.split(/\s+/);
    if (parts.length === 2) {
      const m = ["january","february","march","april","may","june","july","august","september","october","november","december"]
        .indexOf(parts[0].toLowerCase());
      if (m !== -1) target = new Date(parseInt(parts[1]), m, 1);
    }
  }
  const month = target.getMonth();
  const year  = target.getFullYear();

  const sh    = _getSheet();
  const last  = sh.getLastRow();
  if (last < CFG.DATA_START) { SpreadsheetApp.getUi().alert("No job data found."); return; }

  const data = sh.getRange(CFG.DATA_START, 1, last - CFG.DATA_START + 1, CFG.TOTAL_COLS).getValues();
  const filtered = data.filter(row => {
    const di = row[C.DATE_IN - 1];
    if (!di) return false;
    const parts = String(di).split("/");
    if (parts.length !== 3) return false;
    const d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    return d.getMonth() === month && d.getFullYear() === year;
  });

  if (!filtered.length) {
    SpreadsheetApp.getUi().alert(`No jobs found for ${_monthName(month)} ${year}.`);
    return;
  }

  // Build stats
  const totalRevenue = filtered.reduce((s, r) => s + (parseFloat(r[C.TOTAL_QAR - 1]) || 0), 0);
  const totalVAT     = filtered.reduce((s, r) => s + (parseFloat(r[C.VAT_QAR   - 1]) || 0), 0);
  const statusBreak  = {};
  filtered.forEach(r => {
    const k = String(r[C.STATUS - 1]).replace(/^[^\s]+ /,"");
    statusBreak[k] = (statusBreak[k] || 0) + 1;
  });

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const tmpName= `PDF_TEMP_${Date.now()}`;
  const tmp    = ss.insertSheet(tmpName);
  try {
    _buildPDFSheet(tmp, filtered, month, year, totalRevenue, totalVAT, statusBreak);
    const blob = _exportSheetAsPDF(ss, tmp);
    const folder = _getOrCreateDriveFolder("Falcon Garage / Monthly Reports");
    const file = folder.createFile(blob.setName(`Falcon_Garage_${_monthName(month)}_${year}.pdf`));
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ PDF saved: ${file.getName()} | Jobs: ${filtered.length} | Revenue: QAR ${_fmtQAR(totalRevenue)}`,
      "PDF Generated", 8
    );
  } finally {
    ss.deleteSheet(tmp);
  }
}

function _buildPDFSheet(sh, rows, month, year, revenue, vat, statusBreak) {
  let r = 1;

  // Title
  sh.getRange(r, 1, 1, 9).merge()
    .setValue(`🦅 FALCON GARAGE — Monthly Job Report — ${_monthName(month)} ${year}`)
    .setBackground("#111827").setFontColor("#F59E0B")
    .setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sh.setRowHeight(r, 40); r++;

  sh.getRange(r, 1, 1, 9).merge()
    .setValue(`${CFG.GARAGE_ADDR}  |  ${CFG.GARAGE_TEL}  |  ${CFG.GARAGE_WEB}`)
    .setBackground("#1F2937").setFontColor("#9CA3AF")
    .setFontSize(9).setHorizontalAlignment("center");
  sh.setRowHeight(r, 22); r += 2;

  // Stats row
  [
    ["Total Jobs", rows.length, "#0D47A1"],
    ["Total Revenue (QAR)", _fmtQAR(revenue), "#065F46"],
    ["Total VAT (QAR)", _fmtQAR(vat), "#374151"],
  ].forEach(([lbl, val, bg], i) => {
    const col = i * 3 + 1;
    sh.getRange(r,   col, 1, 3).merge().setValue(lbl).setBackground(bg).setFontColor("#9CA3AF").setFontSize(10).setHorizontalAlignment("center");
    sh.getRange(r+1, col, 1, 3).merge().setValue(val).setBackground(bg).setFontColor("#F59E0B").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  });
  r += 3;

  // Status breakdown
  const sbLine = Object.entries(statusBreak).map(([k,v]) => `${k}: ${v}`).join("   |   ");
  sh.getRange(r, 1, 1, 9).merge().setValue(`Status Breakdown: ${sbLine}`)
    .setBackground("#1A2332").setFontColor("#D1D5DB").setFontSize(9).setHorizontalAlignment("center");
  r += 2;

  // Table headers
  const hdrs = ["Job ID","Date In","Customer","Plate","Make / Model","Status","Work Done","Total (QAR)","Tech"];
  hdrs.forEach((h, i) => {
    sh.getRange(r, i + 1).setValue(h)
      .setBackground("#0F172A").setFontColor("#E2E8F0")
      .setFontSize(8).setFontWeight("bold").setHorizontalAlignment("center");
  });
  r++;

  // Data rows
  rows.forEach((row, idx) => {
    const bg = idx % 2 === 0 ? "#1A2332" : "#111827";
    const cells = [
      row[C.JOB_ID    - 1],
      row[C.DATE_IN   - 1],
      row[C.CUST_NAME - 1],
      row[C.PLATE     - 1],
      `${row[C.MAKE-1]} ${row[C.MODEL-1]}`,
      String(row[C.STATUS - 1]).replace(/^[^\s]+ /,""),
      row[C.WORK_DONE - 1],
      _fmtQAR(parseFloat(row[C.TOTAL_QAR - 1]) || 0),
      row[C.TECHNICIAN- 1],
    ];
    cells.forEach((v, i) => {
      sh.getRange(r, i + 1).setValue(v)
        .setBackground(bg).setFontColor("#D1D5DB").setFontSize(8).setWrap(true);
    });
    r++;
  });

  // Column widths
  [120,80,130,80,110,90,200,90,90].forEach((w, i) => sh.setColumnWidth(i + 1, w));
}

function _exportSheetAsPDF(ss, targetSheet) {
  const ssId  = ss.getId();
  const shId  = targetSheet.getSheetId();
  const url   = `https://docs.google.com/spreadsheets/d/${ssId}/export`
    + `?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false`
    + `&printtitle=false&pagenumbers=false&gridlines=false`
    + `&fzr=false&gid=${shId}`;
  const token = ScriptApp.getOAuthToken();
  const resp  = UrlFetchApp.fetch(url, { headers: { Authorization: `Bearer ${token}` }, muteHttpExceptions: true });
  return resp.getBlob();
}

function _getOrCreateDriveFolder(path) {
  const parts = path.split("/").map(s => s.trim());
  let folder  = DriveApp.getRootFolder();
  parts.forEach(name => {
    const it = folder.getFoldersByName(name);
    folder = it.hasNext() ? it.next() : folder.createFolder(name);
  });
  return folder;
}


// ──────────────────────────────────────────────────────────────
//  UTILITIES
// ──────────────────────────────────────────────────────────────
function _getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${CFG.SHEET_NAME}" not found. Run Setup first.`);
  return sh;
}

function _isValidEmail(e) {
  return typeof e === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim());
}

function _monthName(m) {
  return ["January","February","March","April","May","June","July","August","September","October","November","December"][m];
}

function _colLetter(n) {
  let s = "";
  while (n > 0) { s = String.fromCharCode(64 + (n - 1) % 26 + 1) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function _fmtQAR(v) {
  return "QAR " + Number(v).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
