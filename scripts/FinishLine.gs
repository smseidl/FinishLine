/**
 * ============================================================
 *  FINISH LINE
 *  Track & Field Meet Management System
 *  Built for Middle School Track
 *
 *  Version: see VERSION constant below
 * ============================================================
 *
 *  DATA_ENTRY COLUMN REFERENCE:
 *    A: Meet #
 *    B: Gender
 *    C: Event
 *    D: Athlete Name
 *    E: Relay Team ID (relays only — "1", "2", etc.)
 *    F: Result/Mark (final time, distance, or relay team time)
 *    G: Splits/Attempts (comma-separated)
 *         - Relay:      individual leg split (one row per athlete)
 *         - 800M Run:   "lap1, lap2"
 *         - 1600M Run:  "lap1, lap2, lap3, lap4"
 *         - Shot Put / Discus / Long Jump: "att1, att2, att3"
 *    H: (reserved / notes)
 *    I: Place
 *
 *  EVENT ORDER (matches meet schedule):
 *    Track: 100M Hurdles, 100M Dash, 1600M Run, 800M Relay,
 *           400M Dash, 400M Relay, 800M Run, 200M Dash, 1600M Relay
 *    Field: High Jump, Discus, Shot Put, Long Jump
 * ============================================================
 */

// ── VERSION ──────────────────────────────────────────────────
// Update this one value only when bumping the version.
const VERSION = "v2.11";

// ── EVENT LISTS ──────────────────────────────────────────────

const FULL_EVT = [
  "100M Hurdles", "100 M Dash", "1600 M Run", "800 M Relay",
  "400 M Dash", "400 M Relay", "800 M Run", "200 M Dash", "1600 M Relay",
  "High Jump", "Discus", "Shot Put", "Long Jump",
  "65M Hurdles", "Hurdle Shuttle", "Sprint Relay", "Special Relay", "Distance Relay"
];

const PRINT_EVT = [
  "100M Hurdles", "100 M Dash", "1600 M Run", "800 M Relay",
  "400 M Dash", "400 M Relay", "800 M Run", "200 M Dash", "1600 M Relay",
  "High Jump", "Discus", "Shot Put", "Long Jump"
];

// Events that use comma-separated splits in column G
const SPLIT_EVTS   = ["800 M Run", "1600 M Run"];
const ATTEMPT_EVTS = ["Shot Put", "Discus", "Long Jump"];
const RELAY_EVTS   = ["400 M Relay", "800 M Relay", "1600 M Relay",
                      "Hurdle Shuttle", "Sprint Relay", "Special Relay", "Distance Relay"];

// ── MENU ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏁 FINISH LINE')
    .addItem('1. Build / Rebuild Entire System', 'fullInitialize')
    .addItem('2. Generate Printable Lineup',      'generateLineupReport')
    .addItem('3. Generate Printable Event Forms', 'generateEventFormReport')
    .addSeparator()
    .addItem('4. Check PR Setup (debug)',          'checkPRSetup')
    .addItem('5. Check Meet Roster (debug)',         'checkMeetRoster')
    .addToUi();
}

// ── CHECKBOX BUTTON HANDLER ───────────────────────────────────
// The Home tab has styled checkbox cells that act as buttons.
// This function must be registered as an INSTALLABLE trigger to work
// reliably (toast feedback, flush, getUi). See the Setup Checklist on
// the Home tab for instructions. Function name: onEditInstallable.
//
// Simple onEdit() is intentionally left absent — the installable trigger
// replaces it entirely, avoiding the double-fire problem.

function onEditInstallable(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Home') return;
  if (e.range.getColumn() !== 1) return;
  if (e.range.getValue() !== true) return;

  const row = e.range.getRow();
  e.range.setValue(false); // reset immediately so it looks like a button
  if (row !== 6 && row !== 7) return;

  const home       = e.range.getSheet();
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const meetNum    = home.getRange("B3").getValue();
  const gender     = home.getRange("B4").getValue();
  const statusCell = home.getRange("A8:B8");

  // Clear any previous status from the status cell
  statusCell.merge().setValue("").setBackground(null).setFontColor("#000000").setFontWeight("normal");

  if (!meetNum || !gender) {
    statusCell
      .setValue("⚠️  Select a Meet # and Gender above first.")
      .setBackground("#ea9999").setFontColor("#990000").setFontWeight("bold");
    return;
  }

  const label = row === 6 ? "Lineup" : "Event Forms";

  // Show "generating" state — flush() pushes this to the browser immediately
  // when running as an installable trigger. Expect 2–5 sec startup delay before
  // anything appears; that's GAS spinning up, not a bug.
  statusCell
    .setValue("⏳  Generating " + label + "…  (5–15 sec startup delay is normal)")
    .setBackground("#fff2cc").setFontColor("#7f6000").setFontWeight("bold");
  SpreadsheetApp.flush();

  if (row === 6) generateLineupReport();
  if (row === 7) generateEventFormReport();

  // Persistent green "done" — stays visible so there's no question it finished
  statusCell
    .setValue("✅  " + label + " ready!  (generated " + new Date().toLocaleTimeString() + ")")
    .setBackground("#b6d7a8").setFontColor("#274e13").setFontWeight("bold");
}

// ── 1. SYSTEM BUILDER ─────────────────────────────────────────
//
// ⚠️  DATA SAFETY CONTRACT — do not violate this in future edits:
//   • Data tabs (Schedule, Data_Entry, Roster, School_Records):
//       Only row 1 headers are written. Row 2+ data is NEVER cleared.
//   • Generated tabs (Lineup_View, Event_Form_Printable):
//       sh.clear() is called — these are fully regenerated, no user data.
//   • Home tab: fully rebuilt by fullInitialize. No user data stored here.

function fullInitialize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── SCHEDULE ──
  const sched = getOrCreateSheet(ss, 'Schedule');
  sched.getRange(1, 1, 1, 8)
    .setValues([["Meet #", "Date", "Time", "Type", "Location", "Meet Name", "Boys Standing", "Girls Standing"]])
    .setBackground("#444444").setFontColor("white").setFontWeight("bold");
  sched.setFrozenRows(1);
  sched.setColumnWidth(1, 70);
  sched.setColumnWidth(2, 100);
  sched.setColumnWidth(3, 80);
  sched.setColumnWidth(4, 100);
  sched.setColumnWidth(5, 180);
  sched.setColumnWidth(6, 180);
  sched.setColumnWidth(7, 120);
  sched.setColumnWidth(8, 120);
  // Type column — reference only, not used by report logic
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'Dual Meet', 'Invitational', 'Championship', 'Time Trials', 'Relays', '8th Grade Pentathlon', 'Scrimmage'])
    .setAllowInvalid(true).build();
  sched.getRange("D2:D50").setDataValidation(typeRule);

  // ── DATA ENTRY ──
  const entry = getOrCreateSheet(ss, 'Data_Entry');
  entry.getRange(1, 1, 1, 9)
    .setValues([["Meet #", "Gender", "Event", "Athlete Name", "Relay Team ID", "Result/Mark", "Splits/Attempts", "Notes", "Place"]])
    .setBackground("#000000").setFontColor("white").setFontWeight("bold");
  entry.setFrozenRows(1);

  const meetRule   = SpreadsheetApp.newDataValidation().requireValueInRange(sched.getRange("A2:A50")).build();
  const genderRule = SpreadsheetApp.newDataValidation().requireValueInList(['Girls', 'Boys']).build();
  const eventRule  = SpreadsheetApp.newDataValidation().requireValueInList(FULL_EVT).build();

  entry.getRange("A2:A2000").setDataValidation(meetRule);
  entry.getRange("B2:B2000").setDataValidation(genderRule);
  entry.getRange("C2:C2000").setDataValidation(eventRule);
  if (entry.getFilter()) entry.getFilter().remove();
  entry.getRange("A1:I2000").createFilter();

  entry.setColumnWidth(1, 70);
  entry.setColumnWidth(2, 70);
  entry.setColumnWidth(3, 150);
  entry.setColumnWidth(4, 180);
  entry.setColumnWidth(5, 110);
  entry.setColumnWidth(6, 110);
  entry.setColumnWidth(7, 220);
  entry.setColumnWidth(8, 160);
  entry.setColumnWidth(9, 70);

  // ── ROSTER ──
  // Col A: Athlete Name (full legal name, for records/roster)
  // Col B: Display Name (the name used in Data_Entry — first name, nickname, etc.)
  //        This is the PRIMARY lookup key for PR matching.
  const roster = getOrCreateSheet(ss, 'Roster');
  const rosterHeaders = [["Athlete Name", "Display Name", "Gender", "Grade", "School", "Email", "Parent Cell"].concat(FULL_EVT)];
  roster.getRange(1, 1, 1, rosterHeaders[0].length)
    .setValues(rosterHeaders)
    .setBackground("#0b5394").setFontColor("white").setFontWeight("bold");
  roster.setColumnWidth(2, 140);
  roster.setFrozenRows(1);
  roster.setFrozenColumns(1);

  // ── SCHOOL RECORDS ──
  // Columns: Gender, Event, Athlete, Record, Year, Notes
  // Only Record (col D, index 3) is used by report logic — Year/Notes are history/reference only.
  const records = getOrCreateSheet(ss, 'School_Records');
  records.getRange(1, 1, 1, 6)
    .setValues([["Gender", "Event", "Athlete", "Record", "Year", "Notes"]])
    .setBackground("#bf9000").setFontColor("white").setFontWeight("bold");
  records.setColumnWidth(5, 80);
  records.setColumnWidth(6, 240);

  // ── HOME TAB ──
  const meetValRule   = SpreadsheetApp.newDataValidation().requireValueInRange(sched.getRange("A2:A50")).build();
  const genderValRule = SpreadsheetApp.newDataValidation().requireValueInList(['Girls', 'Boys']).build();

  const home = getOrCreateSheet(ss, 'Home');
  home.clear();
  home.setColumnWidth(1, 72);
  home.setColumnWidth(2, 480);
  // Header
  home.getRange("A1:B1").merge()
    .setValue("🏁 FINISH LINE " + VERSION)
    .setBackground("#000000").setFontColor("white")
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  home.setRowHeight(1, 48);
  // Selection dropdowns
  home.getRange("A3").setValue("Meet #:").setFontWeight("bold");
  home.getRange("B3").setBackground("#fff2cc").setDataValidation(meetValRule).setHorizontalAlignment("left");
  home.getRange("A4").setValue("Gender:").setFontWeight("bold");
  home.getRange("B4").setBackground("#fff2cc").setDataValidation(genderValRule).setHorizontalAlignment("left");
  // Button rows — checkbox in col A, label in col B
  const btnStyle = (rng, label, bg) => {
    rng.getSheet().getRange(rng.getRow(), 1).insertCheckboxes();
    rng.merge().setValue(label)
      .setBackground(bg).setFontColor("white").setFontWeight("bold")
      .setFontSize(11).setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setWrap(false);
    rng.getSheet().setRowHeight(rng.getRow(), 36);
  };
  btnStyle(home.getRange("B6"), "  ▶  Generate Printable Lineup",      "#38761d");
  btnStyle(home.getRange("B7"), "  ▶  Generate Printable Event Forms", "#1c4587");
  // Row 5: delay hint — static text, always visible near the buttons
  home.getRange("A5:B5").merge()
    .setValue("ℹ️  After clicking, wait 5–15 sec for status to appear below.")
    .setFontSize(8).setFontStyle("italic").setFontColor("#666666");
  home.setRowHeight(5, 16);
  // Row 8: status cell — written to by onEditInstallable to show progress/errors
  home.getRange("A8:B8").merge()
    .setValue("")
    .setBackground(null).setFontWeight("normal");
  home.setRowHeight(8, 28);
  home.getRange("A9:B9").merge()
    .setValue("── Future Features ───────────────────────────")
    .setFontStyle("italic").setFontColor("#aaaaaa").setFontSize(9);

  // ── SETUP CHECKLIST ──
  // Visible reminder for first-time setup and when copying to a new season.
  home.setRowHeight(11, 28);
  home.getRange("A11:B11").merge()
    .setValue("📋  SETUP CHECKLIST")
    .setBackground("#444444").setFontColor("white")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setPaddingTop ? null : null; // padding not available in GAS — just style

  const checks = [
    ["☐", "Add meets to the Schedule tab (Meet #, Date, Meet Name)"],
    ["☐", "Add athletes to the Roster tab (Athlete Name + Display Name)"],
    ["☐", "Add school records to School_Records tab (optional)"],
    ["☐", "Enter results in Data_Entry after each meet"],
    ["⚠️", "TRIGGER SETUP (once per spreadsheet copy):"],
    ["",   "  Extensions → Apps Script → Triggers (clock icon)"],
    ["",   "  + Add Trigger: onEditInstallable | Spreadsheet | On edit"],
    ["",   "  This enables live status updates on the Home tab buttons."],
    ["ℹ️", "NORMAL: Expect a 5–15 sec delay after clicking a button."],
    ["",   "  Google must spin up a script environment before running."],
    ["",   "  The yellow ⏳ status above will appear once it starts."],
  ];
  checks.forEach(([icon, text], idx) => {
    const r = 12 + idx;
    home.getRange(r, 1).setValue(icon).setHorizontalAlignment("center").setFontSize(9);
    home.getRange(r, 2).setValue(text).setFontSize(9)
      .setFontColor(icon === "⚠️" ? "#990000" : "#333333")
      .setFontWeight(icon === "⚠️" ? "bold" : "normal");
    home.setRowHeight(r, 18);
  });

  // ── TAB COLORS ──
  // Home: black. Data tabs: vivid cyan. Output tabs: vivid orange.
  // Using high-saturation colors because Google Sheets tab chips are tiny.
  ss.getSheetByName('Home')             && ss.getSheetByName('Home').setTabColor('#000000');
  ss.getSheetByName('Schedule')         && ss.getSheetByName('Schedule').setTabColor('#00bcd4');
  ss.getSheetByName('Data_Entry')       && ss.getSheetByName('Data_Entry').setTabColor('#00bcd4');
  ss.getSheetByName('Roster')           && ss.getSheetByName('Roster').setTabColor('#00bcd4');
  ss.getSheetByName('School_Records')   && ss.getSheetByName('School_Records').setTabColor('#00bcd4');
  ss.getSheetByName('Lineup_View')      && ss.getSheetByName('Lineup_View').setTabColor('#ff6d00');
  ss.getSheetByName('Event_Form_Printable') && ss.getSheetByName('Event_Form_Printable').setTabColor('#ff6d00');

  // ── PRINTABLE TABS ── (pure output — selection is on Home tab)
  ['Lineup_View', 'Event_Form_Printable'].forEach(name => {
    const sh = getOrCreateSheet(ss, name);
    sh.clear();
    sh.getRange(1, 1, sh.getMaxRows(), 9).clearNote();
  });

  ui.alert('✅ FINISH LINE ' + VERSION + ' — System built successfully!\n\nNext steps:\n1. Add meets to the Schedule tab\n2. Add athletes to the Roster tab\n3. Enter results in Data_Entry\n4. Use the menu to generate printable reports');
}

// ── 2. LINEUP REPORT ──────────────────────────────────────────

function generateLineupReport() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Lineup_View');
  const home    = ss.getSheetByName('Home');
  const meetNum = home.getRange("B3").getValue();
  const gender  = home.getRange("B4").getValue();

  if (!meetNum || !gender) {
    SpreadsheetApp.getUi().alert("Please select a Meet # and Gender on the Home tab first.");
    return;
  }

  const schedData  = ss.getSheetByName("Schedule").getDataRange().getValues();
  const entryData  = ss.getSheetByName("Data_Entry").getDataRange().getValues();
  const rosterData = ss.getSheetByName("Roster").getDataRange().getValues();
  const meetRow    = schedData.find(r => r[0] == meetNum);
  const meetName   = (meetRow?.[5] || "MEET").toUpperCase();

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 9).clearNote();
  applyReportLayout(sheet);

  // Title row
  sheet.getRange("A1:E1").merge()
    .setValue(meetName + " — LINEUP (" + gender + ")")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#000000").setFontColor("white");
  sheet.setRowHeight(1, 36);

  let curL = 3, curR = 3;

  // Layout: first 5 track events → left col, next 4 track events → right col,
  // field events sync to bottom of both cols then alternate left/right.
  const TRACK_COUNT = 9; // PRINT_EVT indices 0–8 are track
  const HALF_TRACK  = 5; // first half in left column

  PRINT_EVT.forEach((ev, i) => {
    // At the transition to field events, sync both columns to the same row
    // so field events start aligned. (Page breaks must be set manually in
    // File → Print → page break settings — GAS has no API for this.)
    if (i === TRACK_COUNT) {
      const syncRow = Math.max(curL, curR);
      curL = syncRow;
      curR = syncRow;
    }
    const isLeft = (i < HALF_TRACK) ? true
                 : (i < TRACK_COUNT) ? false
                 : ((i - TRACK_COUNT) % 2 === 0);
    const col    = isLeft ? 1 : 4;
    let   row    = isLeft ? curL : curR;

    const aths = entryData.filter(r => r[0] == meetNum && r[1] == gender && r[2] == ev);
    const startRow = row;

    // Event header
    sheet.getRange(row, col, 1, 2)
      .setValues([[ev.toUpperCase(), "PR"]])
      .setBackground("#000000").setFontColor("white").setFontWeight("bold")
      .setBorder(true, true, true, true, true, null);
    row++;

    if (isRelayEvent(ev)) {
      const teams = [...new Set(aths.map(a => a[4] || "1"))];
      teams.forEach((tId, tIdx) => {
        // Thin spacer row between teams (not before the first)
        if (tIdx > 0) {
          sheet.getRange(row, col, 1, 2)
            .setValues([["", ""]])
            .setBackground("#dddddd")
            .setBorder(true, true, true, true, true, null);
          row++;
        }
        // Team header — dark to distinguish from athlete rows
        sheet.getRange(row, col, 1, 2)
          .setValues([["\u25b8 TEAM " + tId, "LEG SPLIT"]])
          .setBackground("#444444").setFontColor("white").setFontWeight("bold")
          .setBorder(true, true, true, true, true, null);
        row++;
        aths.filter(a => a[4] == tId).forEach((m, idx) => {
          sheet.getRange(row, col, 1, 2)
            .setValues([["  " + (idx + 1) + ". " + m[3], m[6] || ""]])
            .setBorder(true, true, true, true, true, null);
          row++;
        });
      });
    } else {
      aths.forEach(a => {
        const pr = findPR(rosterData, a[3], ev);
        const prDisplay = (pr && pr !== "-") ? pr : "—";
        sheet.getRange(row, col, 1, 2)
          .setValues([[a[3], prDisplay]])
          .setBorder(true, true, true, true, true, null);
        if (pr && pr !== "-") {
          sheet.getRange(row, col + 1).setFontColor("#1155cc");
        } else {
          // Highlight athletes with no prior PR — useful pre-meet signal
          sheet.getRange(row, col + 1).setBackground("#cfe2f3").setNote("⭐ No prior PR on record");
        }
        row++;
      });
    }

    // Pad to minimum 6 data rows
    const minRow = startRow + 6;
    while (row < minRow) {
      sheet.getRange(row, col, 1, 2)
        .setValues([["", ""]])
        .setBorder(true, true, true, true, true, null);
      row++;
    }

    if (isLeft) curL = row + 1; else curR = row + 1;
  });
}

// ── 3. EVENT FORM REPORT ──────────────────────────────────────

function generateEventFormReport() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Event_Form_Printable');
  const home    = ss.getSheetByName('Home');
  const meetNum = home.getRange("B3").getValue();
  const gender  = home.getRange("B4").getValue();

  if (!meetNum || !gender) {
    SpreadsheetApp.getUi().alert("Please select a Meet # and Gender on the Home tab first.");
    return;
  }

  const schedData   = ss.getSheetByName("Schedule").getDataRange().getValues();
  const entryData   = ss.getSheetByName("Data_Entry").getDataRange().getValues();
  const recordsData = ss.getSheetByName("School_Records").getDataRange().getValues();
  const rosterData  = ss.getSheetByName("Roster").getDataRange().getValues();

  const meetRow  = schedData.find(r => r[0] == meetNum);
  if (!meetRow) {
    try {
      SpreadsheetApp.getUi().alert("Meet #" + meetNum + " not found in Schedule.");
    } catch(err) {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Home')
        .getRange("A8:B8").merge()
        .setValue("❌  Meet #" + meetNum + " not found in Schedule.")
        .setBackground("#ea9999").setFontColor("#990000").setFontWeight("bold");
    }
    return;
  }

  const meetName    = (meetRow[5] || "MEET").toUpperCase();
  const standing    = (gender === "Boys") ? meetRow[6] : meetRow[7];
  const standingStr = standing ? " | STANDING: " + standing : "";

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 9).clearNote();
  applyReportLayout(sheet);

  // Title row
  sheet.getRange("A1:E1").merge()
    .setValue(meetName + standingStr + " (" + gender + ")")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#1c4587").setFontColor("white");
  sheet.setRowHeight(1, 36);

  let curL = 3, curR = 3;

  // Layout: first 5 track events → left col, next 4 track events → right col,
  // field events sync to bottom of both cols then alternate left/right.
  const TRACK_COUNT = 9; // PRINT_EVT indices 0–8 are track
  const HALF_TRACK  = 5; // first half in left column
  PRINT_EVT.forEach((ev, i) => {
    // At the transition to field events, sync both columns to the same row
    // so field events start aligned. (Page breaks must be set manually in
    // File → Print → page break settings — GAS has no API for this.)
    if (i === TRACK_COUNT) {
      const syncRow = Math.max(curL, curR);
      curL = syncRow;
      curR = syncRow;
    }
    const isLeft = (i < HALF_TRACK) ? true
                 : (i < TRACK_COUNT) ? false
                 : ((i - TRACK_COUNT) % 2 === 0);
    const col    = isLeft ? 1 : 4;
    let   row    = isLeft ? curL : curR;

    const startRow  = row;
    const schoolRec = formatCellValue(recordsData.find(r => r[0] == gender && r[1] == ev)?.[3]) || null;
    const aths      = entryData.filter(r => r[0] == meetNum && r[1] == gender && r[2] == ev);

    // Event header
    sheet.getRange(row, col, 1, 2)
      .setValues([[ev.toUpperCase(), "Rec: " + (schoolRec || "")]])
      .setBackground("#444444").setFontColor("white").setFontWeight("bold")
      .setBorder(true, true, true, true, true, null);
    // Highlight the Rec: cell in amber when a school record exists
    if (schoolRec) {
      sheet.getRange(row, col + 1)
        .setBackground("#bf9000").setFontColor("white");
    }
    row++;

    if (isRelayEvent(ev)) {
      row = renderRelayBlock(sheet, aths, row, col, schoolRec, ev);
    } else if (isSplitEvent(ev)) {
      row = renderSplitBlock(sheet, aths, row, col, rosterData, schoolRec, ev);
    } else if (isAttemptEvent(ev)) {
      row = renderAttemptBlock(sheet, aths, row, col, rosterData, schoolRec, ev);
    } else {
      row = renderStandardBlock(sheet, aths, row, col, rosterData, schoolRec, ev);
    }

    // Pad to minimum 6 data rows
    const minRow = startRow + 6;
    while (row < minRow) {
      sheet.getRange(row, col, 1, 2)
        .setValues([["", ""]])
        .setBorder(true, true, true, true, true, null);
      row++;
    }

    if (isLeft) curL = row + 1; else curR = row + 1;
  });
}

// ── RENDER HELPERS ────────────────────────────────────────────

/** Standard track event (100M, 200M, 400M, hurdles) */
function renderStandardBlock(sheet, aths, row, col, rosterData, schoolRec, ev) {
  aths.forEach(a => {
    const pr    = findPR(rosterData, a[3], ev);
    const res   = formatCellValue(a[5]);
    const place = a[8];
    const plTag = (place !== "" && place !== null && place !== undefined) ? "  [" + place + "]" : "";

    const nameCell   = sheet.getRange(row, col);
    const resultCell = sheet.getRange(row, col + 1);
    const rowRange   = sheet.getRange(row, col, 1, 2);
    const nameText   = a[3] + "  (PR: " + (pr || "—") + ")";
    if (pr && pr !== "-") {
      nameCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(nameText)
          .setTextStyle(a[3].length + 2, nameText.length,
            SpreadsheetApp.newTextStyle().setForegroundColor("#1155cc").build())
          .build()
      );
    } else {
      nameCell.setValue(nameText);
    }
    resultCell.setValue(res + plTag);
    rowRange.setBorder(true, true, true, true, true, null);

    // Highlight result cell and name cell
    if (res && !isNoMark(res)) {
      if (schoolRec && isBetter(res, schoolRec, ev)) {
        rowRange.setBackground("#ffe599");
        resultCell.setNote("🏆 School Record!");
      } else if (!pr || pr === "-" || isBetter(res, pr, ev)) {
        rowRange.setBackground("#b6d7a8");
        resultCell.setNote((!pr || pr === "-") ? "⭐ First Time!" : "🎉 Personal Record!");
      }
    }
    row++;
  });
  return row;
}

/** Relay event — grouped by team ID */
function renderRelayBlock(sheet, aths, row, col, schoolRec, ev) {
  const teams = [...new Set(aths.map(a => a[4] || "1"))];
  teams.forEach(tId => {
    const members = aths.filter(a => a[4] == tId);
    members.forEach((m, idx) => {
      sheet.getRange(row, col, 1, 2)
        .setValues([[(idx + 1) + ". " + m[3], formatCellValue(m[6]) || ""]])
        .setBorder(true, true, true, true, true, null);
      row++;
    });
    // Team total row
    const lead    = members.find(m => m[5]) || members[0];
    const res     = lead ? formatCellValue(lead[5]) : "";
    const place   = lead ? lead[8] : "";
    const plTag   = (place !== "" && place !== null && place !== undefined) ? "  [" + place + "]" : "";
    const totalRange = sheet.getRange(row, col, 1, 2);
    totalRange.setValues([["TEAM " + tId + " TOTAL", res + plTag]])
      .setBackground("#eeeeee").setFontWeight("bold")
      .setBorder(true, true, true, true, true, null);
    if (res && !isNoMark(res) && schoolRec && isBetter(res, schoolRec, ev)) {
      totalRange.setBackground("#ffe599");
      sheet.getRange(row, col + 1).setNote("🏆 School Record!");
    }
    row++;
  });
  return row;
}

/** 800M / 1600M Run — one row per athlete, splits on sub-rows */
function renderSplitBlock(sheet, aths, row, col, rosterData, schoolRec, ev) {
  const lapCount = ev.includes("1600") ? 4 : 2;
  const lapLabels = ev.includes("1600")
    ? ["Lap 1", "Lap 2", "Lap 3", "Lap 4"]
    : ["Lap 1", "Lap 2"];

  aths.forEach(a => {
    const pr    = findPR(rosterData, a[3], ev);
    const res   = formatCellValue(a[5]);
    const place = a[8];
    const plTag = (place !== "" && place !== null && place !== undefined) ? "  [" + place + "]" : "";
    const splits = a[6] ? a[6].toString().split(",").map(s => s.trim()) : [];

    // Athlete main row
    const nameCell   = sheet.getRange(row, col);
    const resultCell = sheet.getRange(row, col + 1);
    const rowRange   = sheet.getRange(row, col, 1, 2);
    const nameText   = a[3] + "  (PR: " + (pr || "—") + ")";
    if (pr && pr !== "-") {
      nameCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(nameText)
          .setTextStyle(a[3].length + 2, nameText.length,
            SpreadsheetApp.newTextStyle().setForegroundColor("#1155cc").build())
          .build()
      );
    } else {
      nameCell.setValue(nameText);
    }
    resultCell.setValue(res + plTag);
    rowRange.setBorder(true, true, true, true, true, null).setFontWeight("bold");

    if (res && !isNoMark(res)) {
      if (schoolRec && isBetter(res, schoolRec, ev)) {
        rowRange.setBackground("#ffe599");
        resultCell.setNote("🏆 School Record!");
      } else if (!pr || pr === "-" || isBetter(res, pr, ev)) {
        rowRange.setBackground("#b6d7a8");
        resultCell.setNote((!pr || pr === "-") ? "⭐ First Time!" : "🎉 Personal Record!");
      }
    }
    row++;

    // Split sub-rows
    for (let l = 0; l < lapCount; l++) {
      sheet.getRange(row, col, 1, 2)
        .setValues([["   ↳ " + lapLabels[l], splits[l] || ""]])
        .setFontSize(8).setFontStyle("italic").setFontColor("#555555")
        .setBorder(true, true, true, true, true, null);
      row++;
    }
  });
  return row;
}

/** Shot Put / Discus / Long Jump — 3 attempts */
function renderAttemptBlock(sheet, aths, row, col, rosterData, schoolRec, ev) {
  aths.forEach(a => {
    const pr       = findPR(rosterData, a[3], ev);
    const res      = formatCellValue(a[5]);
    const place    = a[8];
    const plTag    = (place !== "" && place !== null && place !== undefined) ? "  [" + place + "]" : "";
    const attempts = a[6] ? a[6].toString().split(",").map(s => s.trim()) : [];

    // Athlete main row
    const nameCell   = sheet.getRange(row, col);
    const resultCell = sheet.getRange(row, col + 1);
    const rowRange   = sheet.getRange(row, col, 1, 2);
    const nameText   = a[3] + "  (PR: " + (pr || "—") + ")";
    if (pr && pr !== "-") {
      nameCell.setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(nameText)
          .setTextStyle(a[3].length + 2, nameText.length,
            SpreadsheetApp.newTextStyle().setForegroundColor("#1155cc").build())
          .build()
      );
    } else {
      nameCell.setValue(nameText);
    }
    resultCell.setValue(res + plTag);
    rowRange.setBorder(true, true, true, true, true, null).setFontWeight("bold");

    if (res && !isNoMark(res)) {
      if (schoolRec && isBetter(res, schoolRec, ev)) {
        rowRange.setBackground("#ffe599");
        resultCell.setNote("🏆 School Record!");
      } else if (!pr || pr === "-" || isBetter(res, pr, ev)) {
        rowRange.setBackground("#b6d7a8");
        resultCell.setNote((!pr || pr === "-") ? "⭐ First Time!" : "🎉 Personal Record!");
      }
    }
    row++;

    // Attempt sub-rows
    for (let att = 0; att < 3; att++) {
      sheet.getRange(row, col, 1, 2)
        .setValues([["   ↳ Attempt " + (att + 1), attempts[att] || ""]])
        .setFontSize(8).setFontStyle("italic").setFontColor("#555555")
        .setBorder(true, true, true, true, true, null);
      row++;
    }
  });
  return row;
}

// ── UTILITY FUNCTIONS ─────────────────────────────────────────

/**
 * Returns true if a result value is a no-mark / non-participation indicator.
 * These should never trigger PR or school-record highlights.
 * Covers: - / DNR / DNS / DQ / NH (no height) / NM (no mark) / ND / Scratch / X
 */
function isNoMark(val) {
  if (!val) return false;
  return /^(-|dnr|dns|dq|nh|nm|nd|scratch|x)$/i.test(val.toString().trim());
}

function isRelayEvent(ev) {
  return RELAY_EVTS.some(r => ev === r) || ev.includes("Relay");
}

function isSplitEvent(ev) {
  return SPLIT_EVTS.some(s => ev === s);
}

function isAttemptEvent(ev) {
  return ATTEMPT_EVTS.some(s => ev === s);
}

function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function applyReportLayout(sh) {
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 20);
  sh.setColumnWidth(4, 260);
  sh.setColumnWidth(5, 110);
  // Format result columns as plain text so setValue("18.00") isn't
  // auto-converted to the number 18 by Sheets.
  sh.getRange("B:B").setNumberFormat("@");
  sh.getRange("E:E").setNumberFormat("@");
}

/**
 * Scans Roster and Data_Entry to identify PR lookup problems:
 *   - Athletes with no Display Name (col B) filled in
 *   - Names in Data_Entry that don’t match any Roster Display Name or Athlete Name
 * Run from the 🏁 FINISH LINE menu → "Check PR Setup".
 */
function checkPRSetup() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const rosterData  = ss.getSheetByName('Roster').getDataRange().getValues();
  const entryData   = ss.getSheetByName('Data_Entry').getDataRange().getValues();

  const issues = [];

  // 1. Roster athletes with blank Display Name
  rosterData.slice(1).forEach(r => {
    if (r[0] && !r[1]) {
      issues.push('⚠️ Roster: "' + r[0] + '" has no Display Name — fallback is exact full-name match');
    }
  });

  // 2. Data_Entry names that don’t resolve to any Roster row
  const allEntryNames = [...new Set(
    entryData.slice(1)
      .filter(r => r[3])
      .map(r => r[3].toString().trim())
  )];

  const rosterDisplayNames = rosterData.slice(1)
    .map(r => (r[1] || '').toString().trim().toLowerCase());
  const rosterFullNames = rosterData.slice(1)
    .map(r => (r[0] || '').toString().trim().toLowerCase());

  allEntryNames.forEach(name => {
    const lower = name.toLowerCase();
    const matched = rosterDisplayNames.includes(lower) || rosterFullNames.includes(lower);
    if (!matched) {
      issues.push('❌ Data_Entry: "' + name + '" not found in Roster Display Name or Athlete Name');
    }
  });

  if (issues.length === 0) {
    SpreadsheetApp.getUi().alert('✅ PR Setup looks good! All Data_Entry names match a Roster entry.');
  } else {
    SpreadsheetApp.getUi().alert('PR Setup Issues Found:\n\n' + issues.join('\n'));
  }
}

/**
 * Checks all athlete names entered in Data_Entry for a specific meet
 * against the Roster, reporting any names that don't match.
 *
 * Uses Home B3 (Meet #) and B4 (Gender) as defaults; prompts if blank.
 * Run from the 🏁 FINISH LINE menu → "Check Meet Roster".
 */
function checkMeetRoster() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const ui         = SpreadsheetApp.getUi();
  const home       = ss.getSheetByName('Home');
  const entryData  = ss.getSheetByName('Data_Entry').getDataRange().getValues();
  const rosterData = ss.getSheetByName('Roster').getDataRange().getValues();

  // Use Home B3/B4 as defaults; fall back to prompting
  let meetNum = home ? home.getRange('B3').getValue() : '';
  let gender  = home ? home.getRange('B4').getValue() : '';

  if (!meetNum) {
    const resp = ui.prompt('Check Meet Roster', 'Enter Meet # to check:', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    meetNum = resp.getResponseText().trim();
    if (!meetNum) return;
  }

  // Build lookup sets from Roster
  const rosterDisplayNames = new Set(
    rosterData.slice(1).map(r => (r[1] || '').toString().trim().toLowerCase()).filter(Boolean)
  );
  const rosterFullNames = new Set(
    rosterData.slice(1).map(r => (r[0] || '').toString().trim().toLowerCase()).filter(Boolean)
  );

  // Filter Data_Entry rows for this meet (optionally by gender)
  const meetRows = entryData.slice(1).filter(r => {
    if (r[0] == meetNum && r[3]) {
      return gender ? r[1] == gender : true;
    }
    return false;
  });

  if (meetRows.length === 0) {
    ui.alert('No entries found for Meet #' + meetNum + (gender ? ' / ' + gender : '') + '.');
    return;
  }

  const issues = [];
  const seen   = new Set();

  meetRows.forEach(r => {
    const raw  = r[3].toString().trim();
    const key  = raw.toLowerCase();
    const ev   = r[2] || '?';
    if (seen.has(key)) return; // only report each name once
    seen.add(key);
    const matched = rosterDisplayNames.has(key) || rosterFullNames.has(key);
    if (!matched) {
      issues.push('❌ "' + raw + '"  (first seen in: ' + ev + ')');
    }
  });

  const scope = 'Meet #' + meetNum + (gender ? ' / ' + gender : ' / all genders');
  if (issues.length === 0) {
    ui.alert('✅ All ' + seen.size + ' athlete names in ' + scope + ' match the Roster.');
  } else {
    ui.alert(
      '⚠️ ' + issues.length + ' unmatched name(s) in ' + scope + 
      ' (' + seen.size + ' unique athletes checked):\n\n' +
      issues.join('\n') +
      '\n\nFix: update the name in Data_Entry to exactly match the ' +
      'Display Name (or Athlete Name) in the Roster tab.'
    );
  }
}

/**
 * Safely convert a cell value from getValues() to a display string.
 * Handles the Google Sheets / Apps Script issue where time-formatted cells
 * (e.g. "1:02.45") are returned as JavaScript Date objects anchored at
 * 1899-12-30. Without this, setValue() re-formats them as dates (12/30/1899).
 */
function formatCellValue(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date) {
    const h  = val.getHours();
    const m  = val.getMinutes();
    const s  = val.getSeconds();
    const ms = val.getMilliseconds();
    const secStr = String(s).padStart(2, "0");
    const msStr  = ms > 0 ? "." + String(ms).padStart(3, "0").replace(/0+$/, "") : "";
    if (h > 0) return h + ":" + String(m).padStart(2, "0") + ":" + secStr + msStr;
    return m + ":" + secStr + msStr;
  }
  // Numeric cells (e.g. 13.20, 18.00) lose trailing zeros via toString().
  // Always render to 2 decimal places so 18 → "18.00" and 13.2 → "13.20".
  if (typeof val === "number") return val.toFixed(2);
  return val.toString().trim();
}

/**
 * Compare two performance values.
 * Returns true if val is better than bench for the given event.
 * Field events (Jump/Put/Discus): bigger = better
 * Track events: smaller = better
 */
function isBetter(val, bench, event) {
  if (!val || !bench) return false;
  const parse = (s) => {
    if (s instanceof Date) s = formatCellValue(s); // guard against raw Date objects
    if (typeof s !== 'string') return parseFloat(s) || 0;
    s = s.toString().trim();
    // Feet'Inches" format  e.g. 18'3"
    if (s.includes("'")) {
      const parts = s.split("'");
      const feet  = parseFloat(parts[0]) || 0;
      const inches = parts[1] ? parseFloat(parts[1].replace('"', '')) || 0 : 0;
      return (feet * 12) + inches;
    }
    // MM:SS or MM:SS.ss format
    if (s.includes(':')) {
      const parts = s.split(':');
      return (parseFloat(parts[0]) * 60) + parseFloat(parts[1]);
    }
    return parseFloat(s.replace(/[^\d.]/g, '')) || 0;
  };
  const v = parse(val), b = parse(bench);
  const isField = event.includes("Jump") || event.includes("Put") || event.includes("Discus") || event.includes("High Jump");
  return isField ? (v > b) : (v < b);
}

/**
 * Look up an athlete's PR for a given event from the Roster tab.
 *
 * Matching priority:
 *   1. Display Name (col B) — this is the intended primary key.
 *      Enter the same short name here that the coach uses in Data_Entry.
 *   2. Athlete Name (col A) exact match — fallback if Display Name is blank.
 *
 * ⚠️  No fuzzy/first-name guessing. If a PR is missing, check that the
 *      Display Name in Roster exactly matches what’s in Data_Entry col D.
 */
function findPR(rData, athleteName, eventName) {
  if (!rData || rData.length < 2) return "";
  const headers = rData[0];
  const evIdx   = headers.indexOf(eventName);
  if (evIdx < 0) return "";

  const name = athleteName.toString().trim().toLowerCase();

  // 1. Match on Display Name (col B, index 1)
  let athRow = rData.slice(1).find(r =>
    r[1] && r[1].toString().trim().toLowerCase() === name
  );

  // 2. Fallback: exact match on Athlete Name (col A, index 0)
  if (!athRow) {
    athRow = rData.slice(1).find(r =>
      r[0] && r[0].toString().trim().toLowerCase() === name
    );
  }

  return athRow ? formatCellValue(athRow[evIdx]) : "";
}
