/**
 * ================================================================
 *  FALCON GARAGE — JOB UPDATE TRACKER
 *  Google Sheets + Apps Script System
 *  Version 2.0  — Email Alerts + Monthly PDF Summary
 * ================================================================
 *
 *  QUICK SETUP (one-time, ~60 seconds):
 *  ─────────────────────────────────────
 *  1. Open a NEW blank Google Sheet
 *  2. Click  Extensions → Apps Script
 *  3. Delete ALL existing code in the editor
 *  4. Paste THIS entire file
 *  5. Press Ctrl+S  (or Cmd+S on Mac) to save
 *  6. Close the Apps Script tab
 *  7. Refresh your Google Sheet  (F5 / Cmd+R)
 *  8. A new menu  "🔧 Falcon Garage"  appears in the menu bar
 *  9. Click it → "⚙️ Setup Sheet (First Time Only)"
 * 10. When Google asks for permission, click "Allow"
 *     (Email and Drive access are needed for v2 features)
 * 11. Your tracker is live!
 *
 *  DAILY USE:
 *  ─────────────────────────────────────
 *  • Menu → "📋 New Job Card"  to auto-generate a job card number
 *  • Fill columns C–L + email in Column U
 *  • Column M dropdown → select current status
 *  • Column R → click  "📱 Send Update"  to open WhatsApp
 *  • Menu → 📧 Email Alerts  to send email updates to customers
 *  • Menu → 📄 Monthly PDF Summary  to generate monthly reports
 *  • Each row is a PERMANENT record — never delete rows
 *
 *  NEW IN v2.0:
 *  ─────────────────────────────────────
 *  • Column U  — Customer Email address
 *  • Email single customer update (selected row)
 *  • Bulk email all active jobs
 *  • Bulk email Ready-for-Collection alerts
 *  • Monthly PDF job summary (saves to Drive + optional email)
 *
 * ================================================================
 */


// ================================================================
//  CONFIGURATION
// ================================================================

const CFG = {
  GARAGE_NAME: "Falcon Garage",
  SHEET_NAME:  "JOB UPDATE TRACKER",
  HEADER_ROW:  3,
  DATA_START:  4,
  MAX_ROWS:    500,
  JOB_PREFIX:  "FG",
  TOTAL_COLS:  21,    // A–U  (v2 adds Customer Email in col U)
};


// ================================================================
//  COLUMN MAP
// ================================================================

const C = {
  JOB_CARD:    1,   // A
  DATE_OPENED: 2,   // B
  CUST_NAME:   3,   // C
  MOBILE:      4,   // D
  VEHICLE:     5,   // E
  PLATE:       6,   // F
  VIN:         7,   // G
  MILEAGE:     8,   // H
  COMPLAINT:   9,   // I
  DIAGNOSIS:   10,  // J
  WORK_DONE:   11,  // K
  PARTS:       12,  // L
  STATUS:      13,  // M
  READY:       14,  // N
  TECHNICIAN:  15,  // O
  ADVISOR:     16,  // P
  UPDATE_DATE: 17,  // Q
  WHATSAPP:    18,  // R
  ALERT_SENT:  19,  // S
  REMARKS:     20,  // T
  EMAIL:       21,  // U  ← v2
};

const HEADERS = [
  "Job Card No",
  "Date Opened",
  "Customer Name",
  "Mobile Number",
  "Vehicle Make & Model",
  "Plate Number",
  "VIN / Chassis No",
  "Mileage (KM)",
  "Customer Complaint",
  "Diagnosis Details",
  "Work Done / Repairs Performed",
  "Parts Changed / Replaced",
  "Current Status",
  "Ready for Collection",
  "Technician Name",
  "Service Advisor",
  "Last Update Date",
  "WhatsApp Update Link",
  "Customer Alert Sent",
  "Remarks / Notes",
  "Customer Email",   // ← v2
];

const STATUS_LIST = [
  "Vehicle Received",
  "Diagnosis Started",
  "Waiting Customer Approval",
  "Parts Ordered",
  "Repair In Progress",
  "Additional Work Required",
  "Work Finished",
  "Ready for Collection",
  "Delivered",
];

const YES_NO = ["YES", "NO"];

const P = {
  TITLE_BG:  "#0D2137",
  HEADER_BG: "#1A3A5C",
  BAND_BG:   "#2E86AB",
  WHITE:     "#FFFFFF",
  ROW_A:     "#F0F7FF",
  ROW_B:     "#FFFFFF",
  BORDER:    "#BDC3C7",
  ACCENT:    "#1A6FAE",
  LINK:      "#1D6A39",
  RED_BG:    "#FADBD8",  RED_FG:  "#7B241C",
  ORG_BG:    "#FAE5D3",  ORG_FG:  "#784212",
  BLU_BG:    "#D6EAF8",  BLU_FG:  "#1A5276",
  GRN_BG:    "#D5F5E3",  GRN_FG:  "#1D6A39",
  GRY_BG:    "#EAECEE",  GRY_FG:  "#5D6D7E",
};


// ================================================================
//  MENU
// ================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const emailSubMenu = ui.createMenu("📧 Email Alerts")
    .addItem("📧 Send Email Update (Selected Row)",   "emailCustomerUpdate")
    .addItem("📧 Email All Active Jobs",              "emailAllActiveJobs")
    .addItem("📧 Email Ready-for-Collection Alerts",  "emailReadyAlerts");

  ui.createMenu("🔧 Falcon Garage")
    .addItem("⚙️ Setup Sheet (First Time Only)",      "setupSheet")
    .addSeparator()
    .addItem("📋 New Job Card",                        "createNewJobCard")
    .addItem("📅 Stamp Update Date on Selected Row",   "stampUpdateDate")
    .addSeparator()
    .addItem("📱 Open WhatsApp for Selected Row",      "openWhatsAppForActiveRow")
    .addSeparator()
    .addSubMenu(emailSubMenu)
    .addSeparator()
    .addItem("📄 Monthly PDF Job Summary",             "generateMonthlyPDF")
    .addSeparator()
    .addItem("🔍 Search Jobs",                         "searchJobs")
    .addItem("🔄 Clear All Filters",                   "clearFilters")
    .addSeparator()
    .addItem("🎨 Refresh Row Colours",                 "refreshRowColors")
    .addItem("🔗 Rebuild WhatsApp Links",              "rebuildWhatsAppFormulas")
    .addToUi();
}


// ================================================================
//  MAIN SETUP
// ================================================================

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let ws = ss.getSheetByName(CFG.SHEET_NAME);
  if (!ws) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getName() === "Sheet1") {
      ws = sheets[0];
      ws.setName(CFG.SHEET_NAME);
    } else {
      ws = ss.insertSheet(CFG.SHEET_NAME, 0);
    }
  }

  ws.clear();
  ws.clearConditionalFormatRules();
  ws.setTabColor(P.HEADER_BG);

  _buildTitle(ws);
  _buildHeaders(ws);
  _applyColumnWidths(ws);
  _setupValidation(ws);
  _setupConditionalFormatting(ws);
  _freezeAndFilter(ws);
  _styleDataBand(ws, CFG.DATA_START, CFG.MAX_ROWS);
  _addSampleData(ws);
  _buildWaFormulas(ws);

  SpreadsheetApp.getUi().alert(
    "✅  Falcon Garage Job Tracker v2.0 — Ready!\n\n" +
    "WHAT'S NEW IN v2.0:\n" +
    "──────────────────────────────────────\n" +
    "• Column U  — Customer Email address\n" +
    "• Menu → 📧 Email Alerts  (3 options)\n" +
    "• Menu → 📄 Monthly PDF Summary\n\n" +
    "HOW TO USE:\n" +
    "──────────────────────────────────────\n" +
    "• Menu → 📋 New Job Card  (auto-numbers the job)\n" +
    "• Fill columns C to T, add customer email in Column U\n" +
    "• Column M  — choose status from dropdown\n" +
    "• Column R  — click  📱 Send Update  to open WhatsApp\n" +
    "• Menu → 📧 Email Alerts  to send email updates\n" +
    "• Menu → 📄 Monthly PDF  to generate monthly reports\n\n" +
    "⚠  IMPORTANT:\n" +
    "Each row = one permanent job record.\n" +
    "Never delete rows — just add new ones below."
  );
}


// ================================================================
//  TITLE & HEADERS
// ================================================================

function _buildTitle(ws) {
  const tc = CFG.TOTAL_COLS;

  ws.setRowHeight(1, 52);
  ws.getRange(1, 1, 1, tc)
    .merge()
    .setValue("🔧   FALCON GARAGE  |  JOB UPDATE TRACKER")
    .setBackground(P.TITLE_BG)
    .setFontColor(P.WHITE)
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  ws.setRowHeight(2, 26);
  ws.getRange(2, 1, 1, tc)
    .merge()
    .setValue(
      "Manage Job Cards  ·  Track Repairs  ·  WhatsApp & Email Updates  ·  " +
      "Monthly PDF Reports  ·  Permanent History"
    )
    .setBackground(P.HEADER_BG)
    .setFontColor("#AACCEE")
    .setFontSize(10)
    .setFontStyle("italic")
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function _buildHeaders(ws) {
  ws.setRowHeight(CFG.HEADER_ROW, 42);
  ws.getRange(CFG.HEADER_ROW, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setBackground(P.BAND_BG)
    .setFontColor(P.WHITE)
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontFamily("Arial")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(null, null, true, null, null, null,
               P.TITLE_BG, SpreadsheetApp.BorderStyle.MEDIUM);
}


// ================================================================
//  COLUMN WIDTHS
// ================================================================

function _applyColumnWidths(ws) {
  [120, 110, 160, 145, 200, 120, 175, 105,
   250, 250, 285, 255, 195, 145, 150, 150,
   145, 158, 140, 250, 210   // 21st = Customer Email
  ].forEach((w, i) => ws.setColumnWidth(i + 1, w));
}


// ================================================================
//  DATA VALIDATION (DROPDOWNS)
// ================================================================

function _setupValidation(ws) {
  const n = CFG.MAX_ROWS;

  ws.getRange(CFG.DATA_START, C.STATUS, n, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(STATUS_LIST, true)
        .setAllowInvalid(false)
        .setHelpText("Select the current repair status")
        .build()
    );

  ws.getRange(CFG.DATA_START, C.READY, n, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(YES_NO, true)
        .setAllowInvalid(false)
        .setHelpText("Is the vehicle ready for customer collection?")
        .build()
    );

  ws.getRange(CFG.DATA_START, C.ALERT_SENT, n, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(YES_NO, true)
        .setAllowInvalid(false)
        .setHelpText("Has the customer been notified?")
        .build()
    );
}


// ================================================================
//  CONDITIONAL FORMATTING  (status-based row colours)
// ================================================================

function _setupConditionalFormatting(ws) {
  const endRow     = CFG.DATA_START + CFG.MAX_ROWS - 1;
  const lastColLtr = _colLetter(CFG.TOTAL_COLS);
  const range      = ws.getRange(`A${CFG.DATA_START}:${lastColLtr}${endRow}`);

  const rules = [
    { status: "Waiting Customer Approval", bg: P.RED_BG, fg: P.RED_FG },
    { status: "Parts Ordered",             bg: P.ORG_BG, fg: P.ORG_FG },
    { status: "Additional Work Required",  bg: P.ORG_BG, fg: P.ORG_FG },
    { status: "Repair In Progress",        bg: P.BLU_BG, fg: P.BLU_FG },
    { status: "Work Finished",             bg: P.GRN_BG, fg: P.GRN_FG },
    { status: "Ready for Collection",      bg: P.GRN_BG, fg: P.GRN_FG },
    { status: "Delivered",                 bg: P.GRY_BG, fg: P.GRY_FG },
  ].map(fmt =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$M${CFG.DATA_START}="${fmt.status}"`)
      .setBackground(fmt.bg)
      .setFontColor(fmt.fg)
      .setRanges([range])
      .build()
  );

  ws.setConditionalFormatRules(rules);
}


// ================================================================
//  FREEZE PANES & AUTO-FILTER
// ================================================================

function _freezeAndFilter(ws) {
  ws.setFrozenRows(CFG.HEADER_ROW);
  ws.setFrozenColumns(1);
  ws.getRange(CFG.HEADER_ROW, 1, 1, CFG.TOTAL_COLS).createFilter();
}


// ================================================================
//  ROW BAND STYLING
// ================================================================

function _styleDataBand(ws, startRow, numRows) {
  const tc = CFG.TOTAL_COLS;
  const full = ws.getRange(startRow, 1, numRows, tc);
  full.setFontFamily("Arial")
      .setFontSize(10)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true,
                 P.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  for (let i = 0; i < numRows; i++) {
    ws.getRange(startRow + i, 1, 1, tc)
      .setBackground(i % 2 === 0 ? P.ROW_A : P.ROW_B);
    ws.setRowHeight(startRow + i, 22);
  }

  ws.getRange(startRow, C.JOB_CARD, numRows, 1)
    .setFontWeight("bold").setFontColor(P.ACCENT);

  [C.DATE_OPENED, C.MOBILE, C.MILEAGE, C.STATUS,
   C.READY, C.UPDATE_DATE, C.PLATE, C.ALERT_SENT].forEach(col =>
    ws.getRange(startRow, col, numRows, 1).setHorizontalAlignment("center")
  );

  ws.getRange(startRow, C.DATE_OPENED, numRows, 1).setNumberFormat("DD-MMM-YYYY");
  ws.getRange(startRow, C.UPDATE_DATE, numRows, 1).setNumberFormat("DD-MMM-YYYY HH:mm");
  ws.getRange(startRow, C.MILEAGE,     numRows, 1).setNumberFormat("#,##0");
  ws.getRange(startRow, C.STATUS,      numRows, 1).setFontWeight("bold");
  ws.getRange(startRow, C.READY,       numRows, 1).setFontWeight("bold");

  [C.COMPLAINT, C.DIAGNOSIS, C.WORK_DONE, C.PARTS, C.REMARKS].forEach(col =>
    ws.getRange(startRow, col, numRows, 1).setWrap(true)
  );

  ws.getRange(startRow, C.EMAIL, numRows, 1)
    .setFontColor("#1A6FAE")
    .setHorizontalAlignment("center");
}


// ================================================================
//  WHATSAPP HYPERLINK FORMULA
// ================================================================

function _waFormula(row) {
  const g = CFG.GARAGE_NAME;
  return (
    `=IF(A${row}="","",` +
    `HYPERLINK(` +
      `"https://wa.me/"` +
      `&REGEXREPLACE(D${row},"[^0-9]","")` +
      `&"?text="` +
      `&ENCODEURL(` +
        `"Dear "&C${row}&","` +
        `&CHAR(10)&CHAR(10)` +
        `&"Vehicle Update – ${g}"` +
        `&CHAR(10)&CHAR(10)` +
        `&"Job Card: "&A${row}` +
        `&CHAR(10)` +
        `&"Vehicle: "&E${row}&" | "&F${row}` +
        `&CHAR(10)&CHAR(10)` +
        `&"Work Completed:"` +
        `&CHAR(10)&IF(K${row}="","N/A",K${row})` +
        `&CHAR(10)&CHAR(10)` +
        `&"Parts Changed:"` +
        `&CHAR(10)&IF(L${row}="","N/A",L${row})` +
        `&CHAR(10)&CHAR(10)` +
        `&"Current Status:"` +
        `&CHAR(10)&IF(M${row}="","Not Set",M${row})` +
        `&CHAR(10)&CHAR(10)` +
        `&"Ready for Collection:"` +
        `&CHAR(10)&IF(N${row}="","Pending",N${row})` +
        `&CHAR(10)&CHAR(10)` +
        `&"Thank you,"` +
        `&CHAR(10)&"${g}"` +
      `)` +
    `,"📱 Send Update"))`
  );
}

function _buildWaFormulas(ws) {
  const endRow = CFG.DATA_START + CFG.MAX_ROWS - 1;
  for (let row = CFG.DATA_START; row <= endRow; row++) {
    ws.getRange(row, C.WHATSAPP).setFormula(_waFormula(row));
  }
  ws.getRange(CFG.DATA_START, C.WHATSAPP, CFG.MAX_ROWS, 1)
    .setFontColor(P.LINK)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}


// ================================================================
//  SAMPLE DATA  (3 demo job cards)
// ================================================================

function _addSampleData(ws) {
  if (ws.getRange(CFG.DATA_START, C.JOB_CARD).getValue()) return;

  const now = new Date();
  const d1  = new Date(now); d1.setDate(d1.getDate() - 2);
  const d2  = new Date(now); d2.setDate(d2.getDate() - 1);

  const rows = [
    [
      "FG-2024-001", d1, "Ahmed Al Rashidi", "+971501234567",
      "Toyota Land Cruiser 2020", "DXB-A-12345", "JTMHX02J504012345", 85000,
      "Engine overheating. AC not cooling.",
      "Coolant leak at upper hose. AC compressor low pressure.",
      "Replaced upper radiator hose. Re-gassed AC system. Full road test passed.",
      "Upper radiator hose × 1, AC refrigerant R134a 800 g",
      "Work Finished", "YES", "Mohammed Hassan", "Khalid Mansoor",
      now, "", "YES", "Customer confirmed collection tomorrow morning.",
      "ahmed.rashidi@email.com",
    ],
    [
      "FG-2024-002", d2, "Sara Al Mansoori", "+971502345678",
      "Nissan Patrol 2019", "AUH-B-67890", "JN8AZ2KR5BT012345", 120000,
      "Gearbox slipping on 2nd gear. Whining noise at speed.",
      "Gearbox oil burnt black. Solenoid pack faulty. Torque converter suspect.",
      "", "",
      "Waiting Customer Approval", "NO", "Faisal Al Zaabi", "Khalid Mansoor",
      d2, "", "NO", "Gearbox overhaul quote: AED 4,200. Awaiting customer go-ahead.",
      "sara.mansoori@email.com",
    ],
    [
      "FG-2024-003", now, "James Wilson", "+971503456789",
      "BMW X5 xDrive40i 2021", "SHJ-C-11111", "5UXKR6C56F0K12345", 45000,
      "Check engine light on. Steering vibration above 80 km/h.",
      "Fault P0138: O2 sensor bank 1. Front brake discs warped.",
      "Replacing front brake discs and pads. Fitting new O2 sensor bank 1.",
      "Front brake discs × 2, front brake pads set, O2 sensor B1S2",
      "Parts Ordered", "NO", "Mohammed Hassan", "David Chen",
      now, "", "YES", "Parts ETA 2 business days. Customer informed via WhatsApp.",
      "james.wilson@email.com",
    ],
  ];

  ws.getRange(CFG.DATA_START, 1, rows.length, CFG.TOTAL_COLS).setValues(rows);
}


// ================================================================
//  onEdit TRIGGER  — auto-timestamp, auto-set Ready, rebuild WA
// ================================================================

function onEdit(e) {
  if (!e) return;
  const ws  = e.range.getSheet();
  if (ws.getName() !== CFG.SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < CFG.DATA_START) return;

  const stampTriggers = [
    C.JOB_CARD, C.CUST_NAME, C.MOBILE, C.VEHICLE, C.PLATE, C.VIN,
    C.MILEAGE, C.COMPLAINT, C.DIAGNOSIS, C.WORK_DONE, C.PARTS,
    C.STATUS, C.READY, C.TECHNICIAN, C.ADVISOR, C.REMARKS, C.EMAIL,
  ];

  if (stampTriggers.includes(col)) {
    ws.getRange(row, C.UPDATE_DATE)
      .setValue(new Date())
      .setNumberFormat("DD-MMM-YYYY HH:mm");
  }

  if (col === C.JOB_CARD) {
    const dateCell = ws.getRange(row, C.DATE_OPENED);
    if (!dateCell.getValue()) {
      dateCell.setValue(new Date()).setNumberFormat("DD-MMM-YYYY");
    }
  }

  if (col === C.STATUS) {
    const status = ws.getRange(row, C.STATUS).getValue();
    if (status === "Ready for Collection" || status === "Delivered") {
      if (!ws.getRange(row, C.READY).getValue()) {
        ws.getRange(row, C.READY).setValue("YES");
      }
    }
  }

  const waTriggers = [
    C.JOB_CARD, C.CUST_NAME, C.MOBILE, C.VEHICLE, C.PLATE,
    C.WORK_DONE, C.PARTS, C.STATUS, C.READY,
  ];
  if (waTriggers.includes(col)) {
    ws.getRange(row, C.WHATSAPP)
      .setFormula(_waFormula(row))
      .setFontColor(P.LINK)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }
}


// ================================================================
//  NEW JOB CARD
// ================================================================

function createNewJobCard() {
  const ws = _getSheet(); if (!ws) return;

  const nextRow   = Math.max(ws.getLastRow() + 1, CFG.DATA_START);
  const jobCardNo = _nextJobCardNo(ws);

  ws.getRange(nextRow, C.JOB_CARD)
    .setValue(jobCardNo).setFontWeight("bold").setFontColor(P.ACCENT);
  ws.getRange(nextRow, C.DATE_OPENED)
    .setValue(new Date()).setNumberFormat("DD-MMM-YYYY");
  ws.getRange(nextRow, C.STATUS).setValue("Vehicle Received");
  ws.getRange(nextRow, C.READY).setValue("NO");
  ws.getRange(nextRow, C.ALERT_SENT).setValue("NO");
  ws.getRange(nextRow, C.UPDATE_DATE)
    .setValue(new Date()).setNumberFormat("DD-MMM-YYYY HH:mm");
  ws.getRange(nextRow, C.WHATSAPP)
    .setFormula(_waFormula(nextRow))
    .setFontColor(P.LINK).setFontWeight("bold").setHorizontalAlignment("center");

  _styleDataBand(ws, nextRow, 1);
  SpreadsheetApp.setActiveSheet(ws);
  ws.setActiveRange(ws.getRange(nextRow, C.CUST_NAME));

  SpreadsheetApp.getUi().alert(
    `✅  New Job Card Created\n\n` +
    `Job Card No : ${jobCardNo}\n` +
    `Row         : ${nextRow}\n\n` +
    `Fill in customer & vehicle details.\n` +
    `Add customer email in Column U for email alerts.`
  );
}

function _nextJobCardNo(ws) {
  const year   = new Date().getFullYear();
  const prefix = `${CFG.JOB_PREFIX}-${year}-`;
  const last   = ws.getLastRow();
  if (last < CFG.DATA_START) return `${prefix}001`;

  const existing = ws
    .getRange(CFG.DATA_START, C.JOB_CARD, last - CFG.DATA_START + 1, 1)
    .getValues().flat()
    .filter(v => String(v).startsWith(prefix))
    .map(v => parseInt(String(v).replace(prefix, ""), 10) || 0);

  const max = existing.length ? Math.max(...existing) : 0;
  return `${prefix}${String(max + 1).padStart(3, "0")}`;
}


// ================================================================
//  OPEN WHATSAPP  — modal dialog with clickable button
// ================================================================

function openWhatsAppForActiveRow() {
  const ws = _getSheet(); if (!ws) return;

  const row = ws.getActiveRange().getRow();
  if (row < CFG.DATA_START) {
    SpreadsheetApp.getUi().alert(
      "⚠️  Please click on a job data row first (row 4 or below)."
    );
    return;
  }

  const v        = ws.getRange(row, 1, 1, CFG.TOTAL_COLS).getValues()[0];
  const jobCard  = v[C.JOB_CARD - 1]  || "N/A";
  const custName = v[C.CUST_NAME - 1] || "Customer";
  const mobile   = String(v[C.MOBILE - 1]).replace(/[^0-9]/g, "");
  const vehicle  = v[C.VEHICLE - 1]   || "N/A";
  const plate    = v[C.PLATE - 1]     || "N/A";
  const workDone = v[C.WORK_DONE - 1] || "In Progress";
  const parts    = v[C.PARTS - 1]     || "N/A";
  const status   = v[C.STATUS - 1]    || "Not Set";
  const ready    = v[C.READY - 1]     || "Pending";

  if (!mobile) {
    SpreadsheetApp.getUi().alert(
      "⚠️  No mobile number found in Column D for this row."
    );
    return;
  }

  const message =
    `Dear ${custName},\n\n` +
    `Vehicle Update – ${CFG.GARAGE_NAME}\n\n` +
    `Job Card: ${jobCard}\n` +
    `Vehicle: ${vehicle} | ${plate}\n\n` +
    `Work Completed:\n${workDone}\n\n` +
    `Parts Changed:\n${parts}\n\n` +
    `Current Status:\n${status}\n\n` +
    `Ready for Collection:\n${ready}\n\n` +
    `Thank you,\n${CFG.GARAGE_NAME}`;

  const waUrl   = `https://wa.me/${mobile}?text=${encodeURIComponent(message)}`;
  const safeMsg = message
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body  { font-family:Arial,sans-serif; padding:18px; margin:0; color:#222; }
      h3    { color:#1A3A5C; margin:0 0 12px; font-size:16px; }
      .info { background:#EBF5FB; border-left:4px solid #2980B9;
              padding:8px 12px; border-radius:4px; font-size:13px; margin-bottom:14px; }
      .btn  { display:block; width:100%; box-sizing:border-box;
              padding:13px; background:#25D366; color:#fff; font-size:16px;
              font-weight:bold; text-align:center; text-decoration:none;
              border-radius:8px; margin-bottom:10px; }
      .btn:hover { background:#1ebe57; }
      summary { cursor:pointer; color:#555; font-size:12px; margin-top:8px; }
      pre   { background:#f5f5f5; padding:10px; border-radius:4px;
              font-size:11px; white-space:pre-wrap; word-break:break-word; margin-top:6px; }
    </style></head><body>
      <h3>📱 WhatsApp Update — ${jobCard}</h3>
      <div class="info">
        <b>Customer:</b> ${custName} &nbsp;|&nbsp; <b>Mobile:</b> +${mobile}
      </div>
      <a class="btn" href="${waUrl}" target="_blank">📱 Open WhatsApp Now</a>
      <p style="font-size:11px;color:#888;margin:4px 0 0;">
        WhatsApp Web (or desktop app) will open with the message pre-filled.
      </p>
      <details>
        <summary>▸ Preview message</summary>
        <pre>${safeMsg}</pre>
      </details>
    </body></html>
  `).setWidth(400).setHeight(310);

  SpreadsheetApp.getUi().showModalDialog(html, "Send WhatsApp Update");

  ws.getRange(row, C.UPDATE_DATE)
    .setValue(new Date())
    .setNumberFormat("DD-MMM-YYYY HH:mm");
}


// ================================================================
//  STAMP UPDATE DATE
// ================================================================

function stampUpdateDate() {
  const ws = _getSheet(); if (!ws) return;
  const row = ws.getActiveRange().getRow();
  if (row < CFG.DATA_START) {
    SpreadsheetApp.getUi().alert("Please select a data row first."); return;
  }
  const now = new Date();
  ws.getRange(row, C.UPDATE_DATE).setValue(now).setNumberFormat("DD-MMM-YYYY HH:mm");
  SpreadsheetApp.getUi().alert(`✅  Date stamped: ${now.toLocaleString()}`);
}


// ================================================================
//  SEARCH JOBS
// ================================================================

function searchJobs() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "🔍 Search Jobs",
    "Type a Job Card No, Customer Name, Plate Number, or Vehicle:",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const query = res.getResponseText().trim().toLowerCase();
  if (!query) return;

  const ws   = _getSheet(); if (!ws) return;
  const last = ws.getLastRow();
  if (last < CFG.DATA_START) { ui.alert("No job data found."); return; }

  const searchCols = [C.JOB_CARD, C.CUST_NAME, C.VEHICLE, C.PLATE];
  const hits = [];

  for (let row = CFG.DATA_START; row <= last; row++) {
    const data = ws.getRange(row, 1, 1, CFG.TOTAL_COLS).getValues()[0];
    if (searchCols.some(col => String(data[col - 1]).toLowerCase().includes(query))) {
      hits.push({
        row,
        jobCard:  data[C.JOB_CARD - 1],
        customer: data[C.CUST_NAME - 1],
        vehicle:  data[C.VEHICLE - 1],
        plate:    data[C.PLATE - 1],
        status:   data[C.STATUS - 1],
      });
    }
  }

  if (!hits.length) { ui.alert(`No records found for: "${query}"`); return; }

  let msg = `Found ${hits.length} result(s) for "${query}":\n\n`;
  hits.forEach((h, i) => {
    msg += `${i + 1}.  Row ${h.row}  |  ${h.jobCard}  |  ${h.customer}\n`;
    msg += `    Vehicle: ${h.vehicle}  (${h.plate})\n`;
    msg += `    Status:  ${h.status}\n\n`;
  });

  if (hits.length === 1) {
    ws.setActiveRange(ws.getRange(hits[0].row, 1));
    ui.alert(msg + "(Navigated to record)");
  } else {
    ui.alert(msg + "Tip: Use the column filter arrows to narrow down further.");
  }
}


// ================================================================
//  CLEAR FILTERS
// ================================================================

function clearFilters() {
  const ws = _getSheet(); if (!ws) return;
  const f  = ws.getFilter();
  if (f) {
    f.remove();
    ws.getRange(CFG.HEADER_ROW, 1, 1, CFG.TOTAL_COLS).createFilter();
  }
  SpreadsheetApp.getUi().alert("✅  All filters cleared.");
}


// ================================================================
//  REFRESH ROW COLOURS
// ================================================================

function refreshRowColors() {
  const ws = _getSheet(); if (!ws) return;
  const n  = Math.max(ws.getLastRow() - CFG.DATA_START + 1, CFG.MAX_ROWS);
  _styleDataBand(ws, CFG.DATA_START, n);
  ws.clearConditionalFormatRules();
  _setupConditionalFormatting(ws);
  SpreadsheetApp.getUi().alert("✅  Row colours and conditional formatting refreshed.");
}


// ================================================================
//  REBUILD WHATSAPP FORMULAS
// ================================================================

function rebuildWhatsAppFormulas() {
  const ws = _getSheet(); if (!ws) return;
  _buildWaFormulas(ws);
  SpreadsheetApp.getUi().alert("✅  WhatsApp links rebuilt for all rows.");
}


// ================================================================
// ──────────────────────────────────────────────────────────────
//  ▼▼▼  v2.0: EMAIL ALERTS  ▼▼▼
// ──────────────────────────────────────────────────────────────
// ================================================================


// ================================================================
//  EMAIL UPDATE — selected row
// ================================================================

function emailCustomerUpdate() {
  const ws = _getSheet(); if (!ws) return;
  const ui = SpreadsheetApp.getUi();

  const row = ws.getActiveRange().getRow();
  if (row < CFG.DATA_START) {
    ui.alert("⚠️  Please click on a job data row first (row 4 or below).");
    return;
  }

  const v        = ws.getRange(row, 1, 1, CFG.TOTAL_COLS).getValues()[0];
  const email    = String(v[C.EMAIL - 1]).trim();
  const jobCard  = v[C.JOB_CARD - 1]  || "N/A";
  const custName = v[C.CUST_NAME - 1] || "Customer";

  if (!_isValidEmail(email)) {
    ui.alert(
      `⚠️  No valid email for this job.\n\n` +
      `Job Card : ${jobCard}\n` +
      `Customer : ${custName}\n\n` +
      `Please add the customer email in Column U first.`
    );
    return;
  }

  const confirm = ui.alert(
    "📧 Confirm Email Send",
    `Send vehicle update email to:\n\n` +
    `Customer : ${custName}\n` +
    `Email    : ${email}\n` +
    `Job Card : ${jobCard}\n\n` +
    `Proceed?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    const { subject, body, htmlBody } = _buildEmail(v);
    MailApp.sendEmail({ to: email, subject, body, htmlBody });

    ws.getRange(row, C.ALERT_SENT).setValue("YES");
    ws.getRange(row, C.UPDATE_DATE)
      .setValue(new Date()).setNumberFormat("DD-MMM-YYYY HH:mm");

    ui.alert(
      `✅  Email sent!\n\n` +
      `To     : ${email}\n` +
      `Job    : ${jobCard}\n` +
      `Status : ${v[C.STATUS - 1]}`
    );
  } catch (err) {
    ui.alert(`❌  Failed to send email.\n\nError: ${err.message}`);
  }
}


// ================================================================
//  EMAIL ALL ACTIVE JOBS  — every non-delivered, non-empty job
// ================================================================

function emailAllActiveJobs() {
  _bulkEmail(
    row => row[C.JOB_CARD - 1] &&
           row[C.STATUS - 1] !== "Delivered" &&
           row[C.STATUS - 1] !== "",
    "All Active Jobs"
  );
}


// ================================================================
//  EMAIL READY-FOR-COLLECTION ALERTS
// ================================================================

function emailReadyAlerts() {
  _bulkEmail(
    row => row[C.JOB_CARD - 1] &&
           row[C.READY - 1] === "YES" &&
           row[C.STATUS - 1] !== "Delivered",
    "Ready for Collection"
  );
}


// ================================================================
//  BULK EMAIL HELPER
// ================================================================

function _bulkEmail(filterFn, label) {
  const ws = _getSheet(); if (!ws) return;
  const ui = SpreadsheetApp.getUi();

  const last = ws.getLastRow();
  if (last < CFG.DATA_START) { ui.alert("No job data found."); return; }

  const allData  = ws
    .getRange(CFG.DATA_START, 1, last - CFG.DATA_START + 1, CFG.TOTAL_COLS)
    .getValues();

  const eligible = allData
    .map((row, i) => ({ row, rowNum: CFG.DATA_START + i }))
    .filter(({ row }) => filterFn(row) && _isValidEmail(row[C.EMAIL - 1]));

  if (!eligible.length) {
    ui.alert(
      `⚠️  No eligible jobs found for "${label}".\n\n` +
      `Make sure:\n` +
      `• Customer email is filled in Column U\n` +
      `• Jobs match the filter criteria`
    );
    return;
  }

  let preview =
    `Found ${eligible.length} job(s) to email for "${label}":\n\n`;
  eligible.slice(0, 10).forEach(({ row }, i) => {
    preview +=
      `${i + 1}. ${row[C.JOB_CARD - 1]}  |  ` +
      `${row[C.CUST_NAME - 1]}  |  ${row[C.STATUS - 1]}\n` +
      `   → ${row[C.EMAIL - 1]}\n`;
  });
  if (eligible.length > 10) {
    preview += `   ...and ${eligible.length - 10} more\n`;
  }
  preview += `\nSend all ${eligible.length} email(s)?`;

  if (ui.alert("📧 Bulk Email", preview, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let sent = 0, failed = 0, failedList = "";

  eligible.forEach(({ row, rowNum }) => {
    try {
      const { subject, body, htmlBody } = _buildEmail(row);
      MailApp.sendEmail({ to: row[C.EMAIL - 1], subject, body, htmlBody });
      ws.getRange(rowNum, C.ALERT_SENT).setValue("YES");
      ws.getRange(rowNum, C.UPDATE_DATE)
        .setValue(new Date()).setNumberFormat("DD-MMM-YYYY HH:mm");
      sent++;
    } catch (err) {
      failed++;
      failedList += `\n  • ${row[C.JOB_CARD - 1]}  (${row[C.EMAIL - 1]})`;
    }
  });

  let result = `📧 Bulk Email Complete\n\n✅ Sent    : ${sent}\n`;
  if (failed) result += `❌ Failed  : ${failed}${failedList}`;
  ui.alert(result);
}


// ================================================================
//  EMAIL BUILDER  — returns { subject, body, htmlBody }
// ================================================================

function _buildEmail(v) {
  const jobCard  = v[C.JOB_CARD - 1]  || "N/A";
  const custName = v[C.CUST_NAME - 1] || "Valued Customer";
  const vehicle  = v[C.VEHICLE - 1]   || "N/A";
  const plate    = v[C.PLATE - 1]     || "N/A";
  const workDone = v[C.WORK_DONE - 1] || "In progress";
  const parts    = v[C.PARTS - 1]     || "N/A";
  const status   = v[C.STATUS - 1]    || "Not set";
  const ready    = v[C.READY - 1]     || "Pending";
  const remarks  = v[C.REMARKS - 1]   || "";
  const garage   = CFG.GARAGE_NAME;

  const subject = `Vehicle Update – ${jobCard} | ${garage}`;

  const body =
    `Dear ${custName},\n\n` +
    `Vehicle Update – ${garage}\n\n` +
    `Job Card : ${jobCard}\n` +
    `Vehicle  : ${vehicle}  |  Plate: ${plate}\n\n` +
    `Work Completed:\n${workDone}\n\n` +
    `Parts Changed:\n${parts}\n\n` +
    `Current Status       : ${status}\n` +
    `Ready for Collection : ${ready}\n\n` +
    (remarks ? `Notes:\n${remarks}\n\n` : "") +
    `Thank you for choosing ${garage}.\n` +
    `For queries, please reply to this email or call us directly.\n\n` +
    `Best regards,\n${garage} Service Team`;

  const statusColour = ({
    "Vehicle Received":         "#1A5276",
    "Diagnosis Started":        "#1A5276",
    "Waiting Customer Approval":"#922B21",
    "Parts Ordered":            "#784212",
    "Repair In Progress":       "#1A5276",
    "Additional Work Required": "#784212",
    "Work Finished":            "#1D6A39",
    "Ready for Collection":     "#1D6A39",
    "Delivered":                "#5D6D7E",
  })[status] || "#333";

  const readyBadge = ready === "YES"
    ? `<span style="background:#D5F5E3;color:#1D6A39;padding:3px 12px;` +
      `border-radius:12px;font-weight:bold;">&#10003; YES — Ready for Collection</span>`
    : `<span style="background:#FAE5D3;color:#784212;padding:3px 12px;` +
      `border-radius:12px;font-weight:bold;">&#8987; Not Yet Ready</span>`;

  const htmlBody = `
<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:24px 0;">
<tr><td align="center">
<table width="620" cellpadding="0" cellspacing="0"
       style="background:#fff;border-radius:8px;overflow:hidden;
              box-shadow:0 2px 8px rgba(0,0,0,.12);">

  <tr>
    <td style="background:#0D2137;padding:28px 32px;text-align:center;">
      <p style="margin:0;color:#fff;font-size:22px;font-weight:bold;letter-spacing:1px;">
        &#128296; ${garage}
      </p>
      <p style="margin:6px 0 0;color:#AACCEE;font-size:13px;">Vehicle Service Update</p>
    </td>
  </tr>

  <tr>
    <td style="padding:28px 32px 0;">
      <p style="margin:0;font-size:15px;color:#333;">
        Dear <strong>${custName}</strong>,
      </p>
      <p style="margin:10px 0 0;font-size:14px;color:#555;">
        Here is the latest update on your vehicle currently in our care.
      </p>
    </td>
  </tr>

  <tr>
    <td style="padding:20px 32px 0;">
      <table width="100%" cellpadding="10" cellspacing="0"
             style="background:#F0F7FF;border-radius:6px;border:1px solid #D0E8FA;">
        <tr>
          <td width="50%" style="font-size:13px;color:#555;">
            <strong style="color:#1A3A5C;">Job Card No</strong><br>${jobCard}
          </td>
          <td width="50%" style="font-size:13px;color:#555;">
            <strong style="color:#1A3A5C;">Vehicle</strong><br>${vehicle}
          </td>
        </tr>
        <tr>
          <td style="font-size:13px;color:#555;">
            <strong style="color:#1A3A5C;">Plate Number</strong><br>${plate}
          </td>
          <td style="font-size:13px;color:#555;">
            <strong style="color:#1A3A5C;">Current Status</strong><br>
            <span style="color:${statusColour};font-weight:bold;">${status}</span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr>
    <td style="padding:20px 32px 0;">
      <p style="margin:0 0 6px;font-size:13px;font-weight:bold;color:#1A3A5C;
                border-bottom:2px solid #2E86AB;padding-bottom:4px;">
        &#128296; Work Completed
      </p>
      <p style="margin:8px 0 0;font-size:13px;color:#333;line-height:1.7;">
        ${workDone.replace(/\n/g, "<br>")}
      </p>
    </td>
  </tr>

  <tr>
    <td style="padding:16px 32px 0;">
      <p style="margin:0 0 6px;font-size:13px;font-weight:bold;color:#1A3A5C;
                border-bottom:2px solid #2E86AB;padding-bottom:4px;">
        &#128297; Parts Changed / Replaced
      </p>
      <p style="margin:8px 0 0;font-size:13px;color:#333;line-height:1.7;">
        ${parts.replace(/\n/g, "<br>")}
      </p>
    </td>
  </tr>

  ${remarks ? `
  <tr>
    <td style="padding:16px 32px 0;">
      <p style="margin:0 0 6px;font-size:13px;font-weight:bold;color:#1A3A5C;
                border-bottom:2px solid #2E86AB;padding-bottom:4px;">
        &#128221; Additional Notes
      </p>
      <p style="margin:8px 0 0;font-size:13px;color:#333;line-height:1.7;">
        ${remarks.replace(/\n/g, "<br>")}
      </p>
    </td>
  </tr>` : ""}

  <tr>
    <td style="padding:20px 32px;">
      <table width="100%" cellpadding="18" cellspacing="0"
             style="background:#F8F8F8;border-radius:6px;border:1px solid #DDD;
                    text-align:center;">
        <tr>
          <td>
            <p style="margin:0 0 10px;font-size:11px;font-weight:bold;color:#777;
                      text-transform:uppercase;letter-spacing:1px;">
              Collection Status
            </p>
            ${readyBadge}
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr>
    <td style="background:#0D2137;padding:20px 32px;text-align:center;">
      <p style="margin:0;color:#AACCEE;font-size:12px;line-height:1.6;">
        Thank you for choosing
        <strong style="color:#fff;">${garage}</strong>.<br>
        For queries, please reply to this email or call us directly.
      </p>
    </td>
  </tr>

</table>
</td></tr></table>
</body></html>`;

  return { subject, body, htmlBody };
}


// ================================================================
// ──────────────────────────────────────────────────────────────
//  ▼▼▼  v2.0: MONTHLY PDF SUMMARY  ▼▼▼
// ──────────────────────────────────────────────────────────────
// ================================================================


// ================================================================
//  GENERATE MONTHLY PDF  — menu entry point
// ================================================================

function generateMonthlyPDF() {
  const ui    = SpreadsheetApp.getUi();
  const today = new Date();

  // Step 1 — month
  const mRes = ui.prompt(
    "📄 Monthly PDF Summary  (Step 1 of 3)",
    "Enter the MONTH number (1–12):\n\n" +
    "  1=Jan  2=Feb  3=Mar  4=Apr  5=May  6=Jun\n" +
    "  7=Jul  8=Aug  9=Sep  10=Oct  11=Nov  12=Dec\n\n" +
    `Default: ${today.getMonth() + 1}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (mRes.getSelectedButton() !== ui.Button.OK) return;
  const month = parseInt(mRes.getResponseText().trim()) || today.getMonth() + 1;
  if (month < 1 || month > 12) { ui.alert("⚠️  Invalid month. Enter 1–12."); return; }

  // Step 2 — year
  const yRes = ui.prompt(
    "📄 Monthly PDF Summary  (Step 2 of 3)",
    `Enter the YEAR:\n\nDefault: ${today.getFullYear()}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (yRes.getSelectedButton() !== ui.Button.OK) return;
  const year = parseInt(yRes.getResponseText().trim()) || today.getFullYear();

  // Step 3 — optional email recipient
  const eRes = ui.prompt(
    "📄 Monthly PDF Summary  (Step 3 of 3)",
    "Email the PDF to (optional):\n\n" +
    "Enter an email address to receive the report.\n" +
    "Leave blank to save to Google Drive only.",
    ui.ButtonSet.OK_CANCEL
  );
  if (eRes.getSelectedButton() !== ui.Button.OK) return;
  const destEmail = eRes.getResponseText().trim();

  ui.alert(
    "⏳  Generating PDF — please wait...\n\n" +
    `Month : ${_monthName(month)} ${year}\n` +
    `Email : ${destEmail || "(Drive only)"}\n\n` +
    "This may take 15–30 seconds. Click OK to start."
  );

  _runMonthlyPDF(month, year, destEmail);
}


// ================================================================
//  PDF CORE LOGIC
// ================================================================

function _runMonthlyPDF(month, year, destEmail) {
  const ui = SpreadsheetApp.getUi();
  const ws = _getSheet(); if (!ws) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const last = ws.getLastRow();
  if (last < CFG.DATA_START) { ui.alert("No job data found."); return; }

  // Filter jobs whose Date Opened falls in month/year
  const allRows = ws
    .getRange(CFG.DATA_START, 1, last - CFG.DATA_START + 1, CFG.TOTAL_COLS)
    .getValues();

  const jobs = allRows.filter(row => {
    if (!row[C.JOB_CARD - 1]) return false;
    const d = new Date(row[C.DATE_OPENED - 1]);
    return !isNaN(d) && d.getMonth() + 1 === month && d.getFullYear() === year;
  });

  if (!jobs.length) {
    ui.alert(
      `⚠️  No jobs found for ${_monthName(month)} ${year}.\n\n` +
      `Ensure jobs have a "Date Opened" (Column B) in that month.`
    );
    return;
  }

  // Build temp summary sheet
  const tmpName = `_PDF_TMP_${Date.now()}`;
  let tmpWs = ss.insertSheet(tmpName);

  try {
    _buildPDFSheet(tmpWs, jobs, month, year);
    SpreadsheetApp.flush();

    const filename =
      `Falcon_Garage_Summary_${_monthName(month)}_${year}.pdf`;
    const pdfBlob = _exportSheetAsPDF(ss.getId(), tmpWs.getSheetId(), filename);

    // Save to Drive
    const folder   = _getOrCreateDriveFolder("Falcon Garage Reports");
    const driveFile = folder.createFile(pdfBlob);

    // Email if requested
    if (destEmail && _isValidEmail(destEmail)) {
      MailApp.sendEmail({
        to:          destEmail,
        subject:     `${CFG.GARAGE_NAME} — Monthly Summary: ${_monthName(month)} ${year}`,
        body:
          `Please find attached the monthly job summary for ` +
          `${_monthName(month)} ${year}.\n\n` +
          `Total Jobs : ${jobs.length}\n` +
          `Generated  : ${new Date().toLocaleString()}\n\n` +
          `Google Drive: ${driveFile.getUrl()}\n\n` +
          `— ${CFG.GARAGE_NAME} Job Tracker`,
        attachments: [pdfBlob],
      });
    }

    ui.alert(
      `✅  PDF Summary Generated!\n\n` +
      `Month       : ${_monthName(month)} ${year}\n` +
      `Total Jobs  : ${jobs.length}\n` +
      `Saved to    : Google Drive → "Falcon Garage Reports"\n` +
      (destEmail && _isValidEmail(destEmail)
        ? `Emailed to  : ${destEmail}\n` : "") +
      `\nFile: ${filename}`
    );
  } finally {
    ss.deleteSheet(tmpWs);
  }
}


// ================================================================
//  BUILD THE PDF SUMMARY SHEET
// ================================================================

function _buildPDFSheet(ws, jobs, month, year) {
  const garage = CFG.GARAGE_NAME;
  const tz     = Session.getScriptTimeZone();
  const genStr = Utilities.formatDate(new Date(), tz, "dd MMM yyyy, HH:mm");

  // Stats
  const total     = jobs.length;
  const delivered = jobs.filter(r => r[C.STATUS - 1] === "Delivered").length;
  const readyNow  = jobs.filter(
    r => r[C.READY - 1] === "YES" && r[C.STATUS - 1] !== "Delivered"
  ).length;
  const active    = total - delivered - readyNow;

  const statusCount = {};
  STATUS_LIST.forEach(s => { statusCount[s] = 0; });
  jobs.forEach(r => {
    const s = r[C.STATUS - 1];
    if (s in statusCount) statusCount[s]++;
  });

  // ── Column widths for 10-column table ──────────────────────────
  [115, 95, 150, 175, 105, 205, 175, 175, 65, 130]
    .forEach((w, i) => ws.setColumnWidth(i + 1, w));

  // ── Row 1: Title ───────────────────────────────────────────────
  ws.setRowHeight(1, 50);
  ws.getRange(1, 1, 1, 10).merge()
    .setValue(`${garage}  —  Monthly Job Summary`)
    .setBackground(P.TITLE_BG)
    .setFontColor(P.WHITE)
    .setFontSize(18).setFontWeight("bold").setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Row 2: Subtitle ────────────────────────────────────────────
  ws.setRowHeight(2, 26);
  ws.getRange(2, 1, 1, 10).merge()
    .setValue(`${_monthName(month)} ${year}   |   Generated: ${genStr}`)
    .setBackground(P.HEADER_BG)
    .setFontColor("#AACCEE")
    .setFontSize(10).setFontStyle("italic").setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Row 3: Spacer ──────────────────────────────────────────────
  ws.setRowHeight(3, 8);

  // ── Rows 4–5: Stat boxes (4 boxes × 2 cols = 8; cols 9–10 blank) ──
  const statBoxes = [
    { label: "Total Jobs",           value: total,     bg: P.BLU_BG, fg: P.BLU_FG, col: 1 },
    { label: "Active / In Progress", value: active,    bg: P.ORG_BG, fg: P.ORG_FG, col: 3 },
    { label: "Ready for Collection", value: readyNow,  bg: P.GRN_BG, fg: P.GRN_FG, col: 5 },
    { label: "Delivered / Closed",   value: delivered, bg: P.GRY_BG, fg: P.GRY_FG, col: 7 },
  ];
  [4, 5].forEach(r => ws.setRowHeight(r, r === 4 ? 38 : 26));

  statBoxes.forEach(({ label, value, bg, fg, col }) => {
    ws.getRange(4, col, 1, 2).merge()
      .setValue(value)
      .setBackground(bg).setFontColor(fg)
      .setFontSize(24).setFontWeight("bold").setFontFamily("Arial")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(5, col, 1, 2).merge()
      .setValue(label)
      .setBackground(bg).setFontColor(fg)
      .setFontSize(9).setFontFamily("Arial")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });

  // ── Row 6: Spacer ──────────────────────────────────────────────
  ws.setRowHeight(6, 8);

  // ── Row 7: Status breakdown label ─────────────────────────────
  ws.setRowHeight(7, 24);
  ws.getRange(7, 1, 1, 10).merge()
    .setValue("STATUS BREAKDOWN")
    .setBackground(P.HEADER_BG).setFontColor(P.WHITE)
    .setFontSize(10).setFontWeight("bold").setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Rows 8–9: Status counts (5 per row) ───────────────────────
  const sKeys  = Object.keys(statusCount);
  const rowA   = sKeys.slice(0, 5);
  const rowB   = sKeys.slice(5);
  [rowA, rowB].forEach((group, gi) => {
    ws.setRowHeight(8 + gi, 24);
    group.forEach((status, j) => {
      ws.getRange(8 + gi, j * 2 + 1, 1, 2).merge()
        .setValue(`${status}: ${statusCount[status]}`)
        .setBackground("#F5F5F5").setFontColor("#333333")
        .setFontSize(9).setFontFamily("Arial")
        .setHorizontalAlignment("center").setVerticalAlignment("middle")
        .setBorder(true, true, true, true, null, null,
                   "#DDDDDD", SpreadsheetApp.BorderStyle.SOLID);
    });
  });

  // ── Row 10: Spacer ─────────────────────────────────────────────
  ws.setRowHeight(10, 8);

  // ── Row 11: Table header ───────────────────────────────────────
  const tblHdr = [
    "Job Card No", "Date Opened", "Customer Name",
    "Vehicle", "Plate No", "Work Done (Summary)",
    "Parts Changed", "Status", "Ready", "Technician",
  ];
  ws.setRowHeight(11, 30);
  ws.getRange(11, 1, 1, 10)
    .setValues([tblHdr])
    .setBackground(P.BAND_BG).setFontColor(P.WHITE)
    .setFontSize(9).setFontWeight("bold").setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setWrap(true);

  // ── Rows 12+: Job data ─────────────────────────────────────────
  const truncate = (s, n) =>
    String(s).length > n ? String(s).slice(0, n) + "…" : String(s);

  const tblData = jobs.map(r => [
    r[C.JOB_CARD - 1],
    r[C.DATE_OPENED - 1],
    r[C.CUST_NAME - 1],
    r[C.VEHICLE - 1],
    r[C.PLATE - 1],
    truncate(r[C.WORK_DONE - 1], 120),
    truncate(r[C.PARTS - 1],     100),
    r[C.STATUS - 1],
    r[C.READY - 1],
    r[C.TECHNICIAN - 1],
  ]);

  const DATA_ROW = 12;
  ws.getRange(DATA_ROW, 1, tblData.length, 10).setValues(tblData);

  tblData.forEach((_, i) => {
    const r  = DATA_ROW + i;
    const bg = i % 2 === 0 ? "#F0F7FF" : "#FFFFFF";
    ws.setRowHeight(r, 20);
    ws.getRange(r, 1, 1, 10)
      .setBackground(bg).setFontFamily("Arial").setFontSize(9)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true,
                 "#DDDDDD", SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r, 2).setNumberFormat("DD-MMM-YYYY");
    ws.getRange(r, 1).setFontWeight("bold").setFontColor(P.ACCENT);
    ws.getRange(r, 8).setFontWeight("bold");
    ws.getRange(r, 9).setHorizontalAlignment("center");
  });

  // Conditional colours on Status column (col 8) in the table
  if (tblData.length > 0) {
    const statusRange = ws.getRange(DATA_ROW, 8, tblData.length, 1);
    ws.setConditionalFormatRules([
      { status: "Waiting Customer Approval", bg: P.RED_BG },
      { status: "Parts Ordered",             bg: P.ORG_BG },
      { status: "Additional Work Required",  bg: P.ORG_BG },
      { status: "Repair In Progress",        bg: P.BLU_BG },
      { status: "Work Finished",             bg: P.GRN_BG },
      { status: "Ready for Collection",      bg: P.GRN_BG },
      { status: "Delivered",                 bg: P.GRY_BG },
    ].map(({ status, bg }) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(bg)
        .setRanges([statusRange])
        .build()
    ));
  }

  // ── Footer ─────────────────────────────────────────────────────
  const footerRow = DATA_ROW + tblData.length + 1;
  ws.setRowHeight(footerRow, 22);
  ws.getRange(footerRow, 1, 1, 10).merge()
    .setValue(
      `${garage}  |  Confidential — Internal Use Only  |  ` +
      `${_monthName(month)} ${year} Summary`
    )
    .setBackground(P.TITLE_BG).setFontColor("#AACCEE")
    .setFontSize(9).setFontStyle("italic").setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
}


// ================================================================
//  EXPORT SHEET AS PDF  — returns a Blob
// ================================================================

function _exportSheetAsPDF(spreadsheetId, sheetGid, filename) {
  const token = ScriptApp.getOAuthToken();
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
    `?exportFormat=pdf&format=pdf` +
    `&size=A4&portrait=false` +
    `&fitw=true&sheetnames=false&printtitle=false` +
    `&pagenumbers=true&gridlines=false&fzr=false` +
    `&gid=${sheetGid}`;

  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error(
      `PDF export failed (HTTP ${resp.getResponseCode()}). ` +
      `Check that the script has Drive access.`
    );
  }

  return resp.getBlob().setName(filename);
}


// ================================================================
//  GOOGLE DRIVE FOLDER HELPER
// ================================================================

function _getOrCreateDriveFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}


// ================================================================
//  UTILITY HELPERS
// ================================================================

function _getSheet() {
  const ws = SpreadsheetApp.getActiveSpreadsheet()
               .getSheetByName(CFG.SHEET_NAME);
  if (!ws) {
    SpreadsheetApp.getUi().alert(
      "Sheet not found.\n" +
      "Run  🔧 Falcon Garage → ⚙️ Setup Sheet (First Time Only)  first."
    );
  }
  return ws;
}

function _isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}

function _monthName(m) {
  return [
    "", "January", "February", "March", "April",
    "May", "June", "July", "August",
    "September", "October", "November", "December",
  ][m];
}

function _colLetter(col) {
  let result = "";
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}
