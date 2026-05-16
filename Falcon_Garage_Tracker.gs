/**
 * ================================================================
 *  FALCON GARAGE — JOB UPDATE TRACKER
 *  Google Sheets + Apps Script System
 *  Version 1.0
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
 * 11. Your tracker is live!
 *
 *  DAILY USE:
 *  ─────────────────────────────────────
 *  • Menu → "📋 New Job Card"  to auto-generate a job card number
 *  • Fill in columns C–L (customer, vehicle, work details)
 *  • Column M dropdown → select current status
 *  • Column R → click  "📱 Send Update"  to open WhatsApp
 *  • Each row is a PERMANENT record — never delete rows
 *
 * ================================================================
 */


// ================================================================
//  CONFIGURATION  (edit these to customise)
// ================================================================

const CFG = {
  GARAGE_NAME:    "Falcon Garage",
  SHEET_NAME:     "JOB UPDATE TRACKER",
  HEADER_ROW:     3,         // Row number of the column headers
  DATA_START:     4,         // First data row
  MAX_ROWS:       500,       // Pre-formatted rows
  JOB_PREFIX:     "FG",      // Job card prefix  → FG-2024-001
};


// ================================================================
//  COLUMN MAP  (do not change unless you add/remove columns)
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

// Palette
const P = {
  TITLE_BG:   "#0D2137",
  HEADER_BG:  "#1A3A5C",
  BAND_BG:    "#2E86AB",
  WHITE:      "#FFFFFF",
  ROW_A:      "#F0F7FF",
  ROW_B:      "#FFFFFF",
  BORDER:     "#BDC3C7",
  ACCENT:     "#1A6FAE",
  LINK:       "#1D6A39",
  // Status row colours
  RED_BG:     "#FADBD8",  RED_FG:     "#7B241C",
  ORG_BG:     "#FAE5D3",  ORG_FG:     "#784212",
  BLU_BG:     "#D6EAF8",  BLU_FG:     "#1A5276",
  GRN_BG:     "#D5F5E3",  GRN_FG:     "#1D6A39",
  GRY_BG:     "#EAECEE",  GRY_FG:     "#5D6D7E",
};


// ================================================================
//  MENU
// ================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔧 Falcon Garage")
    .addItem("⚙️ Setup Sheet (First Time Only)",      "setupSheet")
    .addSeparator()
    .addItem("📋 New Job Card",                        "createNewJobCard")
    .addItem("📅 Stamp Update Date on Selected Row",   "stampUpdateDate")
    .addSeparator()
    .addItem("📱 Open WhatsApp for Selected Row",      "openWhatsAppForActiveRow")
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

  // Get or create the sheet
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
    "✅  Falcon Garage Job Tracker — Ready!\n\n" +
    "HOW TO USE:\n" +
    "──────────────────────────────\n" +
    "• Menu → 📋 New Job Card  (auto-numbers the job)\n" +
    "• Fill columns C to T for the job details\n" +
    "• Column M  — choose status from dropdown\n" +
    "• Column R  — click  📱 Send Update  to open WhatsApp\n" +
    "• Column Q  updates automatically when you edit a row\n\n" +
    "⚠  IMPORTANT:\n" +
    "Each row = one permanent job record.\n" +
    "Never delete rows — just add new ones below."
  );
}


// ================================================================
//  TITLE & HEADERS
// ================================================================

function _buildTitle(ws) {
  // Row 1 — garage title
  ws.setRowHeight(1, 52);
  const t = ws.getRange(1, 1, 1, 20);
  t.merge()
   .setValue("🔧   FALCON GARAGE  |  JOB UPDATE TRACKER")
   .setBackground(P.TITLE_BG)
   .setFontColor(P.WHITE)
   .setFontSize(20)
   .setFontWeight("bold")
   .setFontFamily("Arial")
   .setHorizontalAlignment("center")
   .setVerticalAlignment("middle");

  // Row 2 — subtitle band
  ws.setRowHeight(2, 26);
  const s = ws.getRange(2, 1, 1, 20);
  s.merge()
   .setValue("Manage Job Cards  ·  Track Repairs  ·  Send WhatsApp Updates  ·  Keep Permanent History")
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
  const r = ws.getRange(CFG.HEADER_ROW, 1, 1, HEADERS.length);
  r.setValues([HEADERS])
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
  [120,110,160,145,200,120,175,105,
   250,250,285,255,195,145,150,150,
   145,158,140,250
  ].forEach((w, i) => ws.setColumnWidth(i + 1, w));
}


// ================================================================
//  DATA VALIDATION (DROPDOWNS)
// ================================================================

function _setupValidation(ws) {
  const n = CFG.MAX_ROWS;

  // Column M — Current Status
  ws.getRange(CFG.DATA_START, C.STATUS, n, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(STATUS_LIST, true)
        .setAllowInvalid(false)
        .setHelpText("Select the current repair status")
        .build()
    );

  // Column N — Ready for Collection
  ws.getRange(CFG.DATA_START, C.READY, n, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(YES_NO, true)
        .setAllowInvalid(false)
        .setHelpText("Is the vehicle ready for customer collection?")
        .build()
    );

  // Column S — Customer Alert Sent
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
  const endRow = CFG.DATA_START + CFG.MAX_ROWS - 1;
  const range  = ws.getRange(`A${CFG.DATA_START}:T${endRow}`);

  const rules = [
    // Red  — needs customer action
    { status: "Waiting Customer Approval", bg: P.RED_BG, fg: P.RED_FG },
    // Orange — waiting on parts / extra work
    { status: "Parts Ordered",             bg: P.ORG_BG, fg: P.ORG_FG },
    { status: "Additional Work Required",  bg: P.ORG_BG, fg: P.ORG_FG },
    // Blue  — actively being worked
    { status: "Repair In Progress",        bg: P.BLU_BG, fg: P.BLU_FG },
    // Green — ready or done
    { status: "Work Finished",             bg: P.GRN_BG, fg: P.GRN_FG },
    { status: "Ready for Collection",      bg: P.GRN_BG, fg: P.GRN_FG },
    // Gray  — closed
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
  ws.setFrozenRows(CFG.HEADER_ROW);          // Freeze title + headers
  ws.setFrozenColumns(1);                    // Freeze Job Card No column
  ws.getRange(CFG.HEADER_ROW, 1, 1, 20).createFilter();
}


// ================================================================
//  ROW BAND STYLING
// ================================================================

function _styleDataBand(ws, startRow, numRows) {
  const endRow = startRow + numRows - 1;

  // Bulk-set fonts and alignment for the whole band at once
  const full = ws.getRange(startRow, 1, numRows, 20);
  full.setFontFamily("Arial")
      .setFontSize(10)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true,
                 P.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Alternating row backgrounds (conditional formatting overrides when status set)
  for (let i = 0; i < numRows; i++) {
    ws.getRange(startRow + i, 1, 1, 20)
      .setBackground(i % 2 === 0 ? P.ROW_A : P.ROW_B);
    ws.setRowHeight(startRow + i, 22);
  }

  // Column-specific tweaks (applied in batch ranges for speed)
  ws.getRange(startRow, C.JOB_CARD, numRows, 1)
    .setFontWeight("bold").setFontColor(P.ACCENT);

  [C.DATE_OPENED, C.MOBILE, C.MILEAGE, C.STATUS,
   C.READY, C.UPDATE_DATE, C.PLATE, C.ALERT_SENT].forEach(col => {
    ws.getRange(startRow, col, numRows, 1).setHorizontalAlignment("center");
  });

  ws.getRange(startRow, C.DATE_OPENED, numRows, 1).setNumberFormat("DD-MMM-YYYY");
  ws.getRange(startRow, C.UPDATE_DATE, numRows, 1).setNumberFormat("DD-MMM-YYYY HH:mm");
  ws.getRange(startRow, C.MILEAGE,     numRows, 1).setNumberFormat("#,##0");
  ws.getRange(startRow, C.STATUS,      numRows, 1).setFontWeight("bold");
  ws.getRange(startRow, C.READY,       numRows, 1).setFontWeight("bold");

  // Wrap long-text columns
  [C.COMPLAINT, C.DIAGNOSIS, C.WORK_DONE, C.PARTS, C.REMARKS].forEach(col => {
    ws.getRange(startRow, col, numRows, 1).setWrap(true);
  });
}


// ================================================================
//  WHATSAPP HYPERLINK FORMULA
//  Uses Google Sheets' native ENCODEURL() for perfect encoding.
//  REGEXREPLACE strips all non-digits from the phone number.
// ================================================================

function _waFormula(row) {
  const g = CFG.GARAGE_NAME;
  // Build the message as a concatenated string inside the formula.
  // CHAR(10) = newline, ENCODEURL() handles all special characters.
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
  // Batch-set one row at a time (setFormula is per-cell)
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
  // Skip if the sheet already has job data
  if (ws.getRange(CFG.DATA_START, C.JOB_CARD).getValue()) return;

  const now  = new Date();
  const d1   = new Date(now); d1.setDate(d1.getDate() - 2);
  const d2   = new Date(now); d2.setDate(d2.getDate() - 1);

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
    ],
    [
      "FG-2024-002", d2, "Sara Al Mansoori", "+971502345678",
      "Nissan Patrol 2019", "AUH-B-67890", "JN8AZ2KR5BT012345", 120000,
      "Gearbox slipping on 2nd gear. Whining noise at speed.",
      "Gearbox oil burnt black. Solenoid pack faulty. Torque converter suspect.",
      "", "",
      "Waiting Customer Approval", "NO", "Faisal Al Zaabi", "Khalid Mansoor",
      d2, "", "NO", "Gearbox overhaul quote: AED 4,200. Awaiting customer go-ahead.",
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
    ],
  ];

  ws.getRange(CFG.DATA_START, 1, rows.length, 20).setValues(rows);
}


// ================================================================
//  onedit TRIGGER  — auto-timestamp, auto-set Ready, rebuild WA
// ================================================================

function onEdit(e) {
  if (!e) return;
  const ws  = e.range.getSheet();
  if (ws.getName() !== CFG.SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < CFG.DATA_START) return;

  // Columns that should trigger a "Last Update Date" stamp
  const stampTriggers = [
    C.JOB_CARD, C.CUST_NAME, C.MOBILE, C.VEHICLE, C.PLATE, C.VIN,
    C.MILEAGE, C.COMPLAINT, C.DIAGNOSIS, C.WORK_DONE, C.PARTS,
    C.STATUS, C.READY, C.TECHNICIAN, C.ADVISOR, C.REMARKS,
  ];

  if (stampTriggers.includes(col)) {
    ws.getRange(row, C.UPDATE_DATE)
      .setValue(new Date())
      .setNumberFormat("DD-MMM-YYYY HH:mm");
  }

  // Auto-fill Date Opened when a brand-new job card number is typed
  if (col === C.JOB_CARD) {
    const dateCell = ws.getRange(row, C.DATE_OPENED);
    if (!dateCell.getValue()) {
      dateCell.setValue(new Date()).setNumberFormat("DD-MMM-YYYY");
    }
  }

  // Auto-set Ready for Collection when status becomes green-phase
  if (col === C.STATUS) {
    const status = ws.getRange(row, C.STATUS).getValue();
    if (status === "Ready for Collection" || status === "Delivered") {
      if (!ws.getRange(row, C.READY).getValue()) {
        ws.getRange(row, C.READY).setValue("YES");
      }
    }
  }

  // Rebuild the WhatsApp link when key data columns change
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
//  NEW JOB CARD  — auto-generates sequential job card number
// ================================================================

function createNewJobCard() {
  const ws = _getSheet(); if (!ws) return;

  const nextRow  = Math.max(ws.getLastRow() + 1, CFG.DATA_START);
  const jobCardNo = _nextJobCardNo(ws);

  ws.getRange(nextRow, C.JOB_CARD).setValue(jobCardNo).setFontWeight("bold").setFontColor(P.ACCENT);
  ws.getRange(nextRow, C.DATE_OPENED).setValue(new Date()).setNumberFormat("DD-MMM-YYYY");
  ws.getRange(nextRow, C.STATUS).setValue("Vehicle Received");
  ws.getRange(nextRow, C.READY).setValue("NO");
  ws.getRange(nextRow, C.ALERT_SENT).setValue("NO");
  ws.getRange(nextRow, C.UPDATE_DATE).setValue(new Date()).setNumberFormat("DD-MMM-YYYY HH:mm");
  ws.getRange(nextRow, C.WHATSAPP)
    .setFormula(_waFormula(nextRow))
    .setFontColor(P.LINK).setFontWeight("bold").setHorizontalAlignment("center");

  // Style the new row
  _styleDataBand(ws, nextRow, 1);

  SpreadsheetApp.setActiveSheet(ws);
  ws.setActiveRange(ws.getRange(nextRow, C.CUST_NAME));

  SpreadsheetApp.getUi().alert(
    `✅  New Job Card Created\n\n` +
    `Job Card No : ${jobCardNo}\n` +
    `Row         : ${nextRow}\n\n` +
    `Please fill in the customer and vehicle details.`
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
    SpreadsheetApp.getUi().alert("⚠️  Please click on a job data row first (row 4 or below).");
    return;
  }

  const v = ws.getRange(row, 1, 1, 20).getValues()[0];
  const jobCard  = v[C.JOB_CARD - 1]    || "N/A";
  const custName = v[C.CUST_NAME - 1]   || "Customer";
  const mobile   = String(v[C.MOBILE - 1]).replace(/[^0-9]/g, "");
  const vehicle  = v[C.VEHICLE - 1]     || "N/A";
  const plate    = v[C.PLATE - 1]       || "N/A";
  const workDone = v[C.WORK_DONE - 1]   || "In Progress";
  const parts    = v[C.PARTS - 1]       || "N/A";
  const status   = v[C.STATUS - 1]      || "Not Set";
  const ready    = v[C.READY - 1]       || "Pending";

  if (!mobile) {
    SpreadsheetApp.getUi().alert("⚠️  No mobile number found in Column D for this row.");
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

  const waUrl = `https://wa.me/${mobile}?text=${encodeURIComponent(message)}`;

  // Sanitise for HTML display
  const safeMsg = message
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
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
      </style>
    </head>
    <body>
      <h3>📱 WhatsApp Update — ${jobCard}</h3>
      <div class="info">
        <b>Customer:</b> ${custName} &nbsp;|&nbsp;
        <b>Mobile:</b> +${mobile}
      </div>
      <a class="btn" href="${waUrl}" target="_blank">📱 Open WhatsApp Now</a>
      <p style="font-size:11px;color:#888;margin:4px 0 0;">
        WhatsApp Web (or desktop app) will open with the message pre-filled.
      </p>
      <details>
        <summary>▸ Preview message</summary>
        <pre>${safeMsg}</pre>
      </details>
    </body>
    </html>
  `).setWidth(400).setHeight(310);

  SpreadsheetApp.getUi().showModalDialog(html, "Send WhatsApp Update");

  // Auto-stamp the update date
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
    const data = ws.getRange(row, 1, 1, 20).getValues()[0];
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
    ws.getRange(CFG.HEADER_ROW, 1, 1, 20).createFilter();
  }
  SpreadsheetApp.getUi().alert("✅  All filters cleared. Every record is now visible.");
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
//  UTILITY
// ================================================================

function _getSheet() {
  const ws = SpreadsheetApp.getActiveSpreadsheet()
               .getSheetByName(CFG.SHEET_NAME);
  if (!ws) {
    SpreadsheetApp.getUi().alert(
      "Sheet not found.\nPlease run  🔧 Falcon Garage → ⚙️ Setup Sheet (First Time Only)  first."
    );
  }
  return ws;
}
