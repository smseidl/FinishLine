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
const VERSION = "v2.34";

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

// Page layout for Event Forms — adjust based on your printer/margin settings
// Typical: 45-50 rows per page. Increase if using narrow margins, decrease for wide margins.
const PAGE_HEIGHT = 45;

// ── MENU ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏁 FINISH LINE')
    .addItem('1. Build / Rebuild Entire System', 'fullInitialize')
    .addItem('2. Generate Printable Lineup',      'generateLineupReport')
    .addItem('3. Generate Printable Event Forms', 'generateEventFormReport')
    .addSeparator()
    .addItem('4. Generate Top Marks (YTD)',       'generateTopMarks')
    .addItem('5. Generate All Athlete Recaps',    'generateAllAthleteRecaps')
    .addSeparator()
    .addItem('6. Check PR Setup (debug)',          'checkPRSetup')
    .addItem('7. Check Meet Roster (debug)',         'checkMeetRoster')
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
  if (row !== 7 && row !== 8 && row !== 9 && row !== 10) return;

  const home       = e.range.getSheet();
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const meetNum    = home.getRange("B3").getValue();
  const gender     = home.getRange("B4").getValue();
  const statusCell = home.getRange("A11:B11");

  // Clear any previous status from the status cell
  statusCell.merge().setValue("").setBackground(null).setFontColor("#000000").setFontWeight("normal");

  // Row 10 (athlete recaps) doesn't require meet/gender selection
  if (row !== 10 && (!meetNum || !gender)) {
    statusCell
      .setValue("⚠️  Select a Meet # and Gender above first.")
      .setBackground("#ea9999").setFontColor("#990000").setFontWeight("bold");
    return;
  }

  const label = row === 7 ? "Lineup" : row === 8 ? "Event Forms" : row === 9 ? "Update PRs" : "Athlete Recaps";

  // Show "generating" state — flush() pushes this to the browser immediately
  // when running as an installable trigger. Expect 2–5 sec startup delay before
  // anything appears; that's GAS spinning up, not a bug.
  statusCell
    .setValue("⏳  Generating " + label + "…  (5–15 sec startup delay is normal)")
    .setBackground("#fff2cc").setFontColor("#7f6000").setFontWeight("bold");
  SpreadsheetApp.flush();

  if (row === 7) generateLineupReport();
  if (row === 8) generateEventFormReport();
  if (row === 9) findAndUpdatePRs();
  if (row === 10) generateAllAthleteRecaps();

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
  try {
    sched.getRange(1, 1, 1, 9)
      .setValues([["Meet #", "Date", "Time", "Type", "Location", "Address", "Meet Name", "Boys Standing", "Girls Standing"]])
      .setBackground("#444444").setFontColor("white").setFontWeight("bold");
  } catch(e) {
    if (e.message.includes('typed columns')) {
      ui.alert('⚠️ Schedule has Typed Columns enabled.\n\nFix:\n1. Select any cell in Schedule\n2. Data menu → Remove column type\n3. Repeat for all columns with types\n4. Re-run "Build / Rebuild Entire System"');
    } else {
      ui.alert('⚠️ Error updating Schedule headers:\n\n' + e.message + '\n\nTry manually removing any Tables or special formatting from the Schedule tab, then re-run "Build / Rebuild Entire System".');
    }
    return;
  }
  sched.setFrozenRows(1);
  sched.setColumnWidth(1, 70);
  sched.setColumnWidth(2, 100);
  sched.setColumnWidth(3, 80);
  sched.setColumnWidth(4, 100);
  sched.setColumnWidth(5, 180);
  sched.setColumnWidth(6, 200);
  sched.setColumnWidth(7, 180);
  sched.setColumnWidth(8, 120);
  sched.setColumnWidth(9, 120);
  // Type column — reference only, not used by report logic
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'Regular', 'Dual Meet', 'Invitational', 'Championship', 'Time Trials', 'Relays', '8th Grade Pentathlon', 'Scrimmage'])
    .setAllowInvalid(true).build();
  sched.getRange("D2:D50").setDataValidation(typeRule);

  // ── DATA ENTRY ──
  const entry = getOrCreateSheet(ss, 'Data_Entry');
  const entryHeaders = ["Meet #", "Gender", "Event", "Athlete Name", "Relay Team ID", "Result/Mark", "Splits/Attempts", "Notes", "Place"];
  
  // Remove filter FIRST — filters can interfere with setValues()
  if (entry.getFilter()) entry.getFilter().remove();
  
  try {
    const currEntryHeaders = entry.getRange(1, 1, 1, 9).getValues()[0];
    const headersMatch = entryHeaders.every((h, i) => currEntryHeaders[i] === h);
    if (!headersMatch) {
      entry.getRange(1, 1, 1, 9).setValues([entryHeaders]);
    }
    entry.getRange(1, 1, 1, 9)
      .setBackground("#000000").setFontColor("white").setFontWeight("bold");
  } catch(e) {
    if (e.message.includes('typed columns')) {
      ui.alert('⚠️ Data_Entry has Typed Columns enabled.\n\nFix:\n1. Select any cell in Data_Entry\n2. Data menu → Remove column type\n3. Repeat for all columns with types\n4. Re-run "Build / Rebuild Entire System"');
    } else {
      ui.alert('⚠️ Error updating Data_Entry headers:\n\n' + e.message + '\n\nTry manually removing any Tables or special formatting from the Data_Entry tab, then re-run "Build / Rebuild Entire System".');
    }
    return;
  }
  entry.setFrozenRows(1);

  const meetRule   = SpreadsheetApp.newDataValidation().requireValueInRange(sched.getRange("A2:A50")).build();
  const genderRule = SpreadsheetApp.newDataValidation().requireValueInList(['Girls', 'Boys']).build();
  const eventRule  = SpreadsheetApp.newDataValidation().requireValueInList(FULL_EVT).build();

  entry.getRange("A2:A2000").setDataValidation(meetRule);
  entry.getRange("B2:B2000").setDataValidation(genderRule);
  entry.getRange("C2:C2000").setDataValidation(eventRule);
  // Re-create filter after all other setup
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
  const rosterHeaders = [["Athlete Name", "Display Name", "Gender", "Grade", "School", "Email", "Parent Cell"].concat(FULL_EVT).concat(["High Jump Start Dist", "Long Jump Start Dist"])];
  try {
    const currRosterHeaders = roster.getRange(1, 1, 1, rosterHeaders[0].length).getValues()[0];
    const headersMatch = rosterHeaders[0].every((h, i) => currRosterHeaders[i] === h);
    if (!headersMatch) {
      roster.getRange(1, 1, 1, rosterHeaders[0].length).setValues(rosterHeaders);
    }
    roster.getRange(1, 1, 1, rosterHeaders[0].length)
      .setBackground("#0b5394").setFontColor("white").setFontWeight("bold");
  } catch(e) {
    if (e.message.includes('typed columns')) {
      ui.alert('⚠️ Roster has Typed Columns enabled.\n\nFix:\n1. Select any cell in Roster\n2. Data menu → Remove column type\n3. Repeat for all columns with types\n4. Re-run "Build / Rebuild Entire System"');
    } else {
      ui.alert('⚠️ Error updating Roster headers:\n\n' + e.message + '\n\nTry manually removing any Tables or special formatting from the Roster tab, then re-run "Build / Rebuild Entire System".');
    }
    return;
  }
  roster.setColumnWidth(2, 140);
  roster.setFrozenRows(1);
  roster.setFrozenColumns(1);

  // ── SCHOOL RECORDS ──
  // Columns: Gender, Event, Athlete, Record, Year, Notes
  // Only Record (col D, index 3) is used by report logic — Year/Notes are history/reference only.
  const records = getOrCreateSheet(ss, 'School_Records');
  const recordHeaders = ["Gender", "Event", "Athlete", "Record", "Year", "Notes"];
  try {
    const currRecordHeaders = records.getRange(1, 1, 1, 6).getValues()[0];
    const headersMatch = recordHeaders.every((h, i) => currRecordHeaders[i] === h);
    if (!headersMatch) {
      records.getRange(1, 1, 1, 6).setValues([recordHeaders]);
    }
    records.getRange(1, 1, 1, 6)
      .setBackground("#bf9000").setFontColor("white").setFontWeight("bold");
  } catch(e) {
    if (e.message.includes('typed columns')) {
      ui.alert('⚠️ School_Records has Typed Columns enabled.\n\nFix:\n1. Select any cell in School_Records\n2. Data menu → Remove column type\n3. Repeat for all columns with types\n4. Re-run "Build / Rebuild Entire System"');
    } else {
      ui.alert('⚠️ Error updating School_Records headers:\n\n' + e.message + '\n\nTry manually removing any Tables or special formatting from the School_Records tab, then re-run "Build / Rebuild Entire System".');
    }
    return;
  }
  records.setColumnWidth(5, 80);
  records.setColumnWidth(6, 240);

  // ── HISTORICAL PRS ──
  // Storage for past-year PRs — one row per athlete per year.
  // Same event columns as Roster for easy copy/paste at end of season.
  const histPR = getOrCreateSheet(ss, 'Historical_PRs');
  const histHeaders = [["Athlete Name", "Display Name", "Year"].concat(FULL_EVT)];
  try {
    const currHistHeaders = histPR.getRange(1, 1, 1,histHeaders[0].length).getValues()[0];
    const headersMatch = histHeaders[0].every((h, i) => currHistHeaders[i] === h);
    if (!headersMatch) {
      histPR.getRange(1, 1, 1, histHeaders[0].length).setValues(histHeaders);
    }
    histPR.getRange(1, 1, 1, histHeaders[0].length)
      .setBackground("#6d9eeb").setFontColor("white").setFontWeight("bold");
  } catch(e) {
    if (e.message.includes('typed columns')) {
      ui.alert('⚠️ Historical_PRs has Typed Columns enabled.\n\nFix:\n1. Select any cell in Historical_PRs\n2. Data menu → Remove column type\n3. Repeat for all columns with types\n4. Re-run "Build / Rebuild Entire System"');
    } else {
      ui.alert('⚠️ Error updating Historical_PRs headers:\n\n' + e.message);
    }
    return;
  }
  histPR.setFrozenRows(1);
  histPR.setFrozenColumns(3);
  histPR.setColumnWidth(1, 180);
  histPR.setColumnWidth(2, 140);
  histPR.setColumnWidth(3, 70);

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
  home.getRange("A5").setValue("Top N:").setFontWeight("bold");
  home.getRange("B5").setValue(7).setBackground("#fff2cc").setHorizontalAlignment("left")
    .setNote("Number of top athletes to show in Top Marks report (e.g., 3, 5, 7, 10)");
  // Button rows — checkbox in col A, label in col B
  const btnStyle = (rng, label, bg) => {
    rng.getSheet().getRange(rng.getRow(), 1).insertCheckboxes();
    rng.merge().setValue(label)
      .setBackground(bg).setFontColor("white").setFontWeight("bold")
      .setFontSize(11).setHorizontalAlignment("left").setVerticalAlignment("middle")
      .setWrap(false);
    rng.getSheet().setRowHeight(rng.getRow(), 36);
  };
  btnStyle(home.getRange("B7"), "  ▶  Generate Printable Lineup",      "#38761d");
  btnStyle(home.getRange("B8"), "  ▶  Generate Printable Event Forms", "#1c4587");
  btnStyle(home.getRange("B9"), "  ▶  Update PRs from This Meet",      "#bf9000");
  btnStyle(home.getRange("B10"), "  ▶  Generate All Athlete Recaps",    "#674ea7");
  // Row 6: delay hint — static text, always visible near the buttons
  home.getRange("A6:B6").clearDataValidations().merge()
    .setValue("ℹ️  After clicking, wait 5–15 sec for status to appear below.")
    .setFontSize(8).setFontStyle("italic").setFontColor("#666666");
  home.setRowHeight(6, 16);
  // Row 11: status cell — written to by onEditInstallable to show progress/errors
  home.getRange("A11:B11").merge()
    .setValue("")
    .setBackground(null).setFontWeight("normal");
  home.setRowHeight(11, 28);
  home.getRange("A12:B12").merge()
    .setValue("── Future Features ───────────────────────────")
    .setFontStyle("italic").setFontColor("#aaaaaa").setFontSize(9);

  // ── SETUP CHECKLIST ──
  // Visible reminder for first-time setup and when copying to a new season.
  home.setRowHeight(14, 28);
  home.getRange("A14:B14").merge()
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
    const r = 15 + idx;
    home.getRange(r, 1).setValue(icon).setHorizontalAlignment("center").setFontSize(9);
    home.getRange(r, 2).setValue(text).setFontSize(9)
      .setFontColor(icon === "⚠️" ? "#990000" : "#333333")
      .setFontWeight(icon === "⚠️" ? "bold" : "normal");
    home.setRowHeight(r, 18);
  });

  // ── TAB COLORS ──
  // Home: black. Data tabs: vivid cyan. Output tabs: vivid orange. Experimental: purple.
  // Using high-saturation colors because Google Sheets tab chips are tiny.
  ss.getSheetByName('Home')             && ss.getSheetByName('Home').setTabColor('#000000');
  ss.getSheetByName('Schedule')         && ss.getSheetByName('Schedule').setTabColor('#00bcd4');
  ss.getSheetByName('Data_Entry')       && ss.getSheetByName('Data_Entry').setTabColor('#00bcd4');
  ss.getSheetByName('Roster')           && ss.getSheetByName('Roster').setTabColor('#00bcd4');
  ss.getSheetByName('School_Records')   && ss.getSheetByName('School_Records').setTabColor('#00bcd4');
  ss.getSheetByName('Historical_PRs')   && ss.getSheetByName('Historical_PRs').setTabColor('#00bcd4');
  ss.getSheetByName('Lineup_View')      && ss.getSheetByName('Lineup_View').setTabColor('#ff6d00');
  ss.getSheetByName('Event_Form_Printable') && ss.getSheetByName('Event_Form_Printable').setTabColor('#ff6d00');
  ss.getSheetByName('Top_Marks')        && ss.getSheetByName('Top_Marks').setTabColor('#674ea7');
  ss.getSheetByName('Athlete_Recaps')   && ss.getSheetByName('Athlete_Recaps').setTabColor('#674ea7');

  // ── PRINTABLE TABS ── (pure output — selection is on Home tab)
  ['Lineup_View', 'Event_Form_Printable', 'Top_Marks', 'Athlete_Recaps'].forEach(name => {
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
  const meetName   = (meetRow?.[6] || "MEET").toUpperCase();

  // Validate that all athlete names exist in Roster
  const unmatchedNames = validateRosterNames(entryData, rosterData, meetNum, gender);
  if (unmatchedNames.length > 0) {
    SpreadsheetApp.getUi().alert(
      '❌ Athletes Not Found in Roster',
      'The following athletes from Data_Entry were not found in the Roster:\n\n' +
      unmatchedNames.join('\n') + '\n\n' +
      'Please add them to the Roster tab first (using the exact Display Name), then regenerate.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Check for athletes with >4 events
  const overLimit = checkAthleteEventCount(entryData, meetNum, gender);
  if (overLimit.length > 0) {
    const warnings = overLimit.map(a => 
      '⚠️ ' + a.name + ': ' + a.count + ' events (' + a.events.join(', ') + ')'
    );
    const response = SpreadsheetApp.getUi().alert(
      '⚠️ Athletes Over 4-Event Limit',
      'The following athletes are registered for more than 4 events:\n\n' +
      warnings.join('\n') + 
      '\n\nMost meets limit athletes to 4 events. Continue anyway?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 9).clearNote();
  // Reset all row heights to default
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);
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
      curL = syncRow + 2; // add 2 blank rows before field events
      curR = syncRow + 2;
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
        
        // Get starting distance for jump events
        let startDist = null;
        if (ev === "High Jump" || ev === "Long Jump") {
          const startDistColName = ev + " Start Dist";
          const rosterHeaders = rosterData[0];
          const startDistIdx = rosterHeaders.indexOf(startDistColName);
          
          if (startDistIdx >= 0) {
            const nameL = a[3].toString().trim().toLowerCase();
            const athleteRow = rosterData.slice(1).find(r => {
              const rosterDisplay = (r[1] || '').toString().trim().toLowerCase();
              const rosterFull = (r[0] || '').toString().trim().toLowerCase();
              return rosterDisplay === nameL || rosterFull === nameL;
            });
            
            if (athleteRow && athleteRow[startDistIdx]) {
              const dist = formatCellValue(athleteRow[startDistIdx]);
              if (dist && dist !== "-") {
                startDist = dist;
              }
            }
          }
        }
        
        const nameCell = sheet.getRange(row, col);
        if (startDist) {
          // Use rich text to format name normally and starting distance smaller/italic
          const athleteName = a[3];
          const startDistText = " (start: " + startDist + ")";
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(athleteName + startDistText)
            .setTextStyle(0, athleteName.length, SpreadsheetApp.newTextStyle().build())
            .setTextStyle(athleteName.length, athleteName.length + startDistText.length, 
              SpreadsheetApp.newTextStyle()
                .setFontSize(9)
                .setItalic(true)
                .build())
            .build();
          nameCell.setRichTextValue(richText);
        } else {
          nameCell.setValue(a[3]);
        }
        
        sheet.getRange(row, col, 1, 2)
          .setBorder(true, true, true, true, true, null);
        sheet.getRange(row, col + 1).setValue(prDisplay);
        
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

  // ── BY-ATHLETE VIEW ──────────────────────────────────────────
  // Start after the by-event section. Sync to ensure we start on a new page.
  const byAthleteStart = Math.max(curL, curR) + 3; // add spacing before new section
  
  // Title row for By-Athlete view
  sheet.getRange(byAthleteStart, 1, 1, 5).merge()
    .setValue(meetName + " — LINEUP BY ATHLETE (" + gender + ")")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#000000").setFontColor("white");
  sheet.setRowHeight(byAthleteStart, 36);

  // Collect all athletes and their events from entryData
  const athleteEvents = {}; // Map: athleteName -> [eventStrings]
  
  entryData.filter(r => r[0] == meetNum && r[1] == gender).forEach(r => {
    const name = r[3];
    const event = r[2];
    const teamId = r[4];
    
    if (!name || !event) return;
    
    if (!athleteEvents[name]) {
      athleteEvents[name] = [];
    }
    
    // For relay events, determine leg position
    let eventText = event;
    if (isRelayEvent(event) && teamId) {
      const teamMembers = entryData.filter(
        m => m[0] == meetNum && m[1] == gender && m[2] == event && m[4] == teamId
      );
      const legNum = teamMembers.findIndex(m => m[3] === name) + 1;
      eventText += " (Team " + teamId + ", Leg " + legNum + ")";
    }
    
    athleteEvents[name].push(eventText);
  });

  // Sort athletes alphabetically
  const sortedAthletes = Object.keys(athleteEvents).sort();
  
  // Batch write all athlete data for performance
  if (sortedAthletes.length > 0) {
    const athDataStart = byAthleteStart + 2;
    
    // Build all values at once
    const athValues = sortedAthletes.map(athName => {
      const events = athleteEvents[athName];
      const eventsList = events.join(", ");
      return [athName, eventsList, '', '', ''];
    });
    
    // Write all values in one call
    sheet.getRange(athDataStart, 1, sortedAthletes.length, 5).setValues(athValues);
    
    // Apply formatting to ranges (much faster than row-by-row)
    sheet.getRange(athDataStart, 1, sortedAthletes.length, 1)
      .setFontWeight("bold")
      .setFontSize(13);
    
    // Merge columns 2-5 for event lists and apply formatting
    for (let i = 0; i < sortedAthletes.length; i++) {
      sheet.getRange(athDataStart + i, 2, 1, 4).merge();
    }
    
    sheet.getRange(athDataStart, 2, sortedAthletes.length, 4)
      .setFontSize(13)
      .setWrap(true);
    
    sheet.getRange(athDataStart, 1, sortedAthletes.length, 5)
      .setBorder(true, true, true, true, false, null);
  }
  
  const athRow = byAthleteStart + 2 + sortedAthletes.length;

  // ── CONFERENCE LINEUP ─────────────────────────────────────────
  // Event-by-event lineup for conference submission using short name format
  // Positioned at end since it's never printed — only used for digital submission
  const confLineupStart = athRow + 3;
  
  // Title row
  sheet.getRange(confLineupStart, 1, 1, 5).merge()
    .setValue(meetName + " — CONFERENCE LINEUP (" + gender + ")")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#1c4587").setFontColor("white");
  sheet.setRowHeight(confLineupStart, 36);
  
  let confRow = confLineupStart + 2;
  
  // Loop through each event and list athletes in conference format
  PRINT_EVT.forEach(ev => {
    const eventAthletes = entryData.filter(r => r[0] == meetNum && r[1] == gender && r[2] == ev);
    
    if (eventAthletes.length === 0) return; // skip events with no entries
    
    // Event header
    sheet.getRange(confRow, 1)
      .setValue(ev.toUpperCase())
      .setFontWeight("bold").setFontSize(11)
      .setBackground("#e0e0e0");
    confRow++;
    
    // Collect athlete names and format for conference
    const conferenceNames = [];
    
    if (isRelayEvent(ev)) {
      // For relays, group by team ID and list all members
      const teams = [...new Set(eventAthletes.map(a => a[4] || "1"))];
      teams.forEach(tId => {
        const members = eventAthletes.filter(a => a[4] == tId);
        members.forEach(m => {
          const displayName = m[3].toString().trim();
          const nameL = displayName.toLowerCase();
          const rosterRow = rosterData.slice(1).find(r => {
            const rosterDisplay = (r[1] || '').toString().trim().toLowerCase();
            const rosterFull = (r[0] || '').toString().trim().toLowerCase();
            return rosterDisplay === nameL || rosterFull === nameL;
          });
          
          if (rosterRow) {
            const displayNameUse = (rosterRow[1] || displayName).toString().trim();
            const fullName = (rosterRow[0] || '').toString().trim();
            const nameParts = fullName.split(/\s+/);
            const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';
            conferenceNames.push({
              displayName: displayNameUse,
              lastName: lastName,
              teamId: tId
            });
          }
        });
      });
      
      // Format relay names with team grouping
      const teamGroups = {};
      conferenceNames.forEach(cn => {
        if (!teamGroups[cn.teamId]) teamGroups[cn.teamId] = [];
        teamGroups[cn.teamId].push(cn);
      });
      
      Object.keys(teamGroups).sort().forEach(tId => {
        const teamMembers = teamGroups[tId];
        const formattedNames = formatConferenceNamesFromParts(teamMembers);
        
        if (formattedNames.length > 0) {
          sheet.getRange(confRow, 1)
            .setValue("Team " + tId + ": " + formattedNames.join(", "))
            .setFontSize(10);
          confRow++;
        }
      });
      
    } else {
      // For individual events, list each athlete
      eventAthletes.forEach(a => {
        const displayName = a[3].toString().trim();
        const nameL = displayName.toLowerCase();
        const rosterRow = rosterData.slice(1).find(r => {
          const rosterDisplay = (r[1] || '').toString().trim().toLowerCase();
          const rosterFull = (r[0] || '').toString().trim().toLowerCase();
          return rosterDisplay === nameL || rosterFull === nameL;
        });
        
        if (rosterRow) {
          const displayNameUse = (rosterRow[1] || displayName).toString().trim();
          const fullName = (rosterRow[0] || '').toString().trim();
          const nameParts = fullName.split(/\s+/);
          const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';
          conferenceNames.push({
            displayName: displayNameUse,
            lastName: lastName
          });
        } else {
          // Fallback: use display name from Data_Entry
          const parts = displayName.split(/\s+/);
          conferenceNames.push({
            displayName: parts.slice(0, -1).join(' ') || parts[0],
            lastName: parts[parts.length - 1] || ''
          });
        }
      });
      
      // Format and output individual names
      const formattedNames = formatConferenceNamesFromParts(conferenceNames);
      if (formattedNames.length > 0) {
        formattedNames.forEach(name => {
          sheet.getRange(confRow, 1)
            .setValue(name)
            .setFontSize(10);
          confRow++;
        });
      }
    }
    
    confRow++; // blank line between events
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

  // Validate that all athlete names exist in Roster
  const unmatchedNames = validateRosterNames(entryData, rosterData, meetNum, gender);
  if (unmatchedNames.length > 0) {
    SpreadsheetApp.getUi().alert(
      '❌ Athletes Not Found in Roster',
      'The following athletes from Data_Entry were not found in the Roster:\n\n' +
      unmatchedNames.join('\n') + '\n\n' +
      'Please add them to the Roster tab first (using the exact Display Name), then regenerate.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Check for athletes with >4 events
  const overLimit = checkAthleteEventCount(entryData, meetNum, gender);
  if (overLimit.length > 0) {
    const warnings = overLimit.map(a => 
      '⚠️ ' + a.name + ': ' + a.count + ' events (' + a.events.join(', ') + ')'
    );
    const response = SpreadsheetApp.getUi().alert(
      '⚠️ Athletes Over 4-Event Limit',
      'The following athletes are registered for more than 4 events:\n\n' +
      warnings.join('\n') + 
      '\n\nMost meets limit athletes to 4 events. Continue anyway?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }

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

  const meetName    = (meetRow[6] || "MEET").toUpperCase();
  const standing    = (gender === "Boys") ? meetRow[7] : meetRow[8];
  const standingStr = standing ? " | STANDING: " + standing : "";

  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 9).clearNote();
  // Reset all row heights to default
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);
  applyReportLayout(sheet);

  // Title row
  sheet.getRange("A1:E1").merge()
    .setValue(meetName + standingStr + " (" + gender + ")")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#1c4587").setFontColor("white");
  sheet.setRowHeight(1, 36);

  // Team scoring summary section (row 2)
  sheet.getRange("A2").setValue("TEAM PLACE:").setFontWeight("bold").setFontSize(10);
  sheet.getRange("B2").setBackground("#fff2cc");
  sheet.getRange("C2").setValue("  ");
  sheet.getRange("D2").setValue("TEAM POINTS:").setFontWeight("bold").setFontSize(10);
  sheet.getRange("E2").setBackground("#fff2cc");
  sheet.setRowHeight(2, 24);

  let curL = 4, curR = 4;

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
      curL = syncRow + 2; // add 2 blank rows before field events
      curR = syncRow + 2;
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

    // Add blank rows for at-meet additions (write-ins)
    // 4 lines for relays, 2 lines for all other events
    const additionCount = isRelayEvent(ev) ? 4 : 2;
    
    sheet.getRange(row, col, 1, 2)
      .setValues([["─ At-Meet Additions ─", ""]])
      .setBackground("#f3f3f3").setFontStyle("italic").setFontSize(8).setFontColor("#666666")
      .setBorder(true, true, true, true, true, null);
    row++;
    
    // Dynamic number of blank write-in rows
    for (let i = 0; i < additionCount; i++) {
      sheet.getRange(row, col, 1, 2)
        .setValues([["", ""]])
        .setBackground("#ffffff")
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
 * Format athlete names for conference roster using Display Name + Last Name.
 * - Input: array of {displayName, lastName} objects
 * - Default format: "DisplayName L" (display name + last initial)
 * - If multiple athletes have the same "DisplayName L", expand all to full last name
 * - Returns array sorted alphabetically by display name
 * - Handles cases like "MJ Smith" → "MJ S" or "Mary Jane Smith" (goes by MJ) → "MJ S"
 */
function formatConferenceNamesFromParts(nameParts) {
  const parsed = nameParts.map(np => {
    const lastInitial = np.lastName.charAt(0).toUpperCase();
    const shortForm = np.displayName + ' ' + lastInitial;
    
    return {
      displayName: np.displayName,
      lastName: np.lastName,
      lastInitial: lastInitial,
      shortForm: shortForm
    };
  });
  
  // Find duplicates (same shortForm)
  const shortFormCounts = {};
  parsed.forEach(p => { 
    shortFormCounts[p.shortForm] = (shortFormCounts[p.shortForm] || 0) + 1;
  });
  
  // Format: use full last name if duplicate, otherwise use short form
  const formatted = parsed.map(p => {
    if (shortFormCounts[p.shortForm] > 1) {
      return p.displayName + ' ' + p.lastName;
    }
    return p.shortForm;
  });
  
  // Sort alphabetically by display name
  return formatted.sort((a, b) => {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  });
}

/**
 * Format athlete names for conference roster.
 * - Default format: "First L" (first name + last initial)
 * - If multiple athletes have the same "First L", expand all to full last name
 * - Returns array sorted alphabetically by first name
 * - Handles multi-word first names (e.g., "Mary Jane Smith" → "Mary Jane S")
 */
function formatConferenceNames(athleteNames) {
  const parsed = athleteNames.map(name => {
    const parts = name.trim().split(/\s+/);
    
    // Last word is the last name, everything else is the first name
    const lastName = parts.length > 1 ? parts[parts.length - 1] : '';
    const firstName = parts.length > 1 ? parts.slice(0, -1).join(' ') : parts[0] || '';
    const lastInitial = lastName.charAt(0).toUpperCase();
    
    return {
      original: name,
      firstName: firstName,
      lastName: lastName,
      lastInitial: lastInitial,
      shortForm: firstName + ' ' + lastInitial
    };
  });
  
  // Find duplicates (same shortForm)
  const shortFormCounts = {};
  parsed.forEach(p => {
    shortFormCounts[p.shortForm] = (shortFormCounts[p.shortForm] || 0) + 1;
  });
  
  // Format: use full last name if duplicate, otherwise use short form
  const formatted = parsed.map(p => {
    if (shortFormCounts[p.shortForm] > 1) {
      return p.firstName + ' ' + p.lastName;
    }
    return p.shortForm;
  });
  
  // Sort alphabetically by first name
  return formatted.sort((a, b) => {
    const aFirst = a.split(' ')[0].toLowerCase();
    const bFirst = b.split(' ')[0].toLowerCase();
    return aFirst.localeCompare(bFirst);
  });
}

/**
 * Returns true if a result value is a no-mark / non-participation indicator.
 * These should never trigger PR or school-record highlights.
 * Covers: - / DNR / DNS / DQ / NH (no height) / NM (no mark) / ND / Scratch / X / NA / CANCELED / CANCELLED
 */
function isNoMark(val) {
  if (!val) return false;
  return /^(-|dnr|dns|dq|nh|nm|nd|scratch|x|na|cancell?ed)$/i.test(val.toString().trim());
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
 * Validate that all athlete names in Data_Entry exist in the Roster.
 * Similar to checkMeetRoster(), but with additional filtering
 * Returns array of unmatched name strings.
 */
function validateRosterNames(entryData, rosterData, meetNum, gender) {
  const rosterDisplayNames = new Set(
    rosterData.slice(1).map(r => (r[1] || '').toString().trim().toLowerCase()).filter(Boolean)
  );
  const rosterFullNames = new Set(
    rosterData.slice(1).map(r => (r[0] || '').toString().trim().toLowerCase()).filter(Boolean)
  );

  const meetRows = entryData.slice(1).filter(r => {
    if (r[0] == meetNum && r[1] == gender && r[3]) return true;
    return false;
  });

  const unmatched = [];
  const seen = new Set();

  meetRows.forEach(r => {
    const raw = r[3].toString().trim();
    const key = raw.toLowerCase();
    const ev = r[2] || '?';
    if (seen.has(key)) return; // only report each name once
    seen.add(key);
    const matched = rosterDisplayNames.has(key) || rosterFullNames.has(key);
    if (!matched) {
      unmatched.push('❌ "' + raw + '"  (in: ' + ev + ')');
    }
  });

  return unmatched;
}

/**
 * Check if any athletes are registered for more than 4 events in a meet.
 * Returns array of {name, count, events} for athletes over the limit.
 * Note: Relays count as 1 event (not per leg).
 */
function checkAthleteEventCount(entryData, meetNum, gender) {
  const athleteEvents = {}; // Map: athleteName -> Set of events
  
  entryData.filter(r => r[0] == meetNum && r[1] == gender && r[3]).forEach(r => {
    const name = r[3].toString().trim();
    const event = r[2];
    
    if (!athleteEvents[name]) {
      athleteEvents[name] = new Set();
    }
    // For relays, only count the event once (not each leg)
    athleteEvents[name].add(event);
  });
  
  // Find athletes with >4 events
  const overLimit = [];
  Object.keys(athleteEvents).forEach(name => {
    const count = athleteEvents[name].size;
    if (count > 4) {
      overLimit.push({
        name: name,
        count: count,
        events: Array.from(athleteEvents[name])
      });
    }
  });
  
  return overLimit;
}

/**
 * Checks all athlete names entered in Data_Entry for a specific meet
 * against the Roster, reporting any names that don't match.
 *
 * Similar to validateRosterNames(), but less filtering.
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

// ── 4. UPDATE PRS FROM MEET ──────────────────────────────────

/**
 * Scans results from a selected meet and updates Roster PRs where athletes
 * have improved. Shows a confirmation dialog before making changes.
 * Uses Home B3 (Meet #) and B4 (Gender).
 * 
 * For individual events: uses Result/Mark (col F)
 * For relay events: uses individual leg split (col G)
 */
function findAndUpdatePRs() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const ui         = SpreadsheetApp.getUi();
  const home       = ss.getSheetByName('Home');
  const meetNum    = home.getRange("B3").getValue();
  const gender     = home.getRange("B4").getValue();
  const statusCell = home.getRange("A10:B10");

  if (!meetNum || !gender) {
    statusCell.setValue("⚠️ Select Meet # and Gender first").setBackground("#f4cccc").setFontColor("#990000");
    SpreadsheetApp.flush();
    ui.alert('❌ Missing Selection', 'Please select Meet # and Gender on the Home tab.', ui.ButtonSet.OK);
    return;
  }

  const entrySheet  = ss.getSheetByName('Data_Entry');
  const rosterSheet = ss.getSheetByName('Roster');
  const entryData   = entrySheet.getDataRange().getValues();
  const rosterData  = rosterSheet.getDataRange().getValues();
  const rosterHeaders = rosterData[0];

  // Filter to this meet/gender, exclude no-marks
  const results = entryData.slice(1).filter(r => {
    if (r[0] != meetNum || r[1] != gender) return false;
    // For relays, we need a split time in column G
    if (isRelayEvent(r[2])) {
      return r[6] && !isNoMark(r[6]); // use split time for relay legs
    }
    // For individual events, use result from column F
    if (isNoMark(r[5])) return false;
    return true;
  });

  if (results.length === 0) {
    statusCell.setValue("ℹ️ No valid results found").setBackground("#fff2cc").setFontColor("#7f6000");
    ui.alert('ℹ️ No Results', 'No valid results found for Meet #' + meetNum + ' / ' + gender + '.', ui.ButtonSet.OK);
    return;
  }

  // Build list of PRs to update
  const updates = []; // [{name, event, oldPR, newPR, rowIdx, colIdx}]
  const errors  = [];

  results.forEach(r => {
    const name   = r[3];
    const event  = r[2];
    // For relay events, use split time (col G); for individual events, use result (col F)
    const result = isRelayEvent(event) ? r[6] : r[5];
    const nameL  = name.toString().trim().toLowerCase();

    // Find athlete in Roster
    let rosterRowIdx = -1;
    for (let i = 1; i < rosterData.length; i++) {
      const displayName = (rosterData[i][1] || '').toString().trim().toLowerCase();
      const fullName    = (rosterData[i][0] || '').toString().trim().toLowerCase();
      if (displayName === nameL || fullName === nameL) {
        rosterRowIdx = i;
        break;
      }
    }

    if (rosterRowIdx === -1) {
      errors.push(name + ' (not in Roster)');
      return;
    }

    // Find event column
    const eventColIdx = rosterHeaders.indexOf(event);
    if (eventColIdx === -1) {
      errors.push(name + ' / ' + event + ' (event column not found)');
      return;
    }

    const currentPR = rosterData[rosterRowIdx][eventColIdx];
    const hasNoPR = (currentPR === '' || currentPR === null || currentPR === undefined || currentPR === '-');

    // Check if this is a PR
    if (hasNoPR || isBetter(result, currentPR, event)) {
      updates.push({
        name: name,
        event: event,
        oldPR: hasNoPR ? '(none)' : formatCellValue(currentPR),
        newPR: formatCellValue(result),
        newPRRaw: result,  // keep raw value for writing
        rowIdx: rosterRowIdx,
        colIdx: eventColIdx
      });
    }
  });

  // Handle errors
  if (errors.length > 0) {
    statusCell.setValue("❌ Athletes not in Roster").setBackground("#f4cccc").setFontColor("#990000");
    ui.alert('❌ Error: Athletes Not Found',
      'The following athletes from Data_Entry were not found in the Roster:\n\n' +
      errors.join('\n') + '\n\n' +
      'Please add them to the Roster tab first (use the exact Display Name).',
      ui.ButtonSet.OK);
    return;
  }

  // No updates needed
  if (updates.length === 0) {
    statusCell.setValue("ℹ️ No PRs to update").setBackground("#d9ead3").setFontColor("#274e13");
    ui.alert('ℹ️ No PRs to Update',
      'All results from Meet #' + meetNum + ' / ' + gender + ' are already recorded or not PRs.',
      ui.ButtonSet.OK);
    return;
  }

  // Show PRs in batches to avoid truncation in alert dialogs
  const BATCH_SIZE = 10;
  const totalBatches = Math.ceil(updates.length / BATCH_SIZE);
  
  for (let batch = 0; batch < totalBatches; batch++) {
    const start = batch * BATCH_SIZE;
    const end = Math.min(start + BATCH_SIZE, updates.length);
    const batchUpdates = updates.slice(start, end);
    
    const previewLines = batchUpdates.map(u =>
      u.name + ' | ' + u.event + ': ' + u.oldPR + ' → ' + u.newPR
    );
    
    const isLastBatch = (batch === totalBatches - 1);
    const title = totalBatches > 1
      ? '🎯 PRs to Update (' + (start + 1) + '-' + end + ' of ' + updates.length + ')'
      : '🎯 Update ' + updates.length + ' PR' + (updates.length === 1 ? '' : 's') + '?';
    
    const message = previewLines.join('\n') + 
      (isLastBatch ? '\n\nProceed with all ' + updates.length + ' update' + (updates.length === 1 ? '' : 's') + '?' : '');
    
    const buttons = isLastBatch ? ui.ButtonSet.YES_NO : ui.ButtonSet.OK_CANCEL;
    const response = ui.alert(title, message, buttons);
    
    // If user cancels on any batch, abort
    if (response === ui.Button.CANCEL || response === ui.Button.NO) {
      statusCell.setValue("⚠️ Update canceled").setBackground("#fff2cc").setFontColor("#7f6000");
      return;
    }
  }

  // Apply updates
  updates.forEach(u => {
    // Update in-memory array (for subsequent updates in this batch)
    rosterData[u.rowIdx][u.colIdx] = u.newPRRaw;
    // Write to sheet (row is 1-indexed, col is 1-indexed) - use raw value not formatted string
    rosterSheet.getRange(u.rowIdx + 1, u.colIdx + 1).setValue(u.newPRRaw);
  });

  statusCell.setValue("✅ " + updates.length + " PR" + (updates.length === 1 ? '' : 's') + " updated")
    .setBackground("#d9ead3").setFontColor("#274e13").setFontWeight("bold");
  SpreadsheetApp.flush();

  ui.alert('✅ PRs Updated',
    updates.length + ' PR' + (updates.length === 1 ? '' : 's') + ' updated successfully in the Roster tab.',
    ui.ButtonSet.OK);
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

// ── 5. TOP MARKS (YTD) ────────────────────────────────────────

/**
 * Generate Top N Marks report showing best boys and girls for each event
 * based on current year Data_Entry results. N is configurable on Home tab (B5).
 */
function generateTopMarks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Top_Marks');
  const home = ss.getSheetByName('Home');
  const entryData = ss.getSheetByName('Data_Entry').getDataRange().getValues();
  
  // Get configurable Top N value from Home tab (default to 7 if not set)
  let topN = home.getRange("B5").getValue();
  if (!topN || isNaN(topN) || topN < 1) {
    topN = 7; // default
  }
  topN = Math.floor(topN); // ensure integer
  
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 9).clearNote();
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);
  
  // Title row with dynamic Top N
  sheet.getRange("A1:D1").merge()
    .setValue("TOP " + topN + " MARKS — YEAR TO DATE")
    .setFontWeight("bold").setFontSize(16)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#674ea7").setFontColor("white");
  sheet.setRowHeight(1, 36);
  
  let row = 3;
  
  // Helper function to parse mark values for comparison
  const parseVal = (val) => {
    if (val instanceof Date) val = formatCellValue(val);
    const str = val.toString().trim();
    if (str.includes("'")) {
      const parts = str.split("'");
      const feet = parseFloat(parts[0]) || 0;
      const inches = parts[1] ? parseFloat(parts[1].replace('"', '')) || 0 : 0;
      return (feet * 12) + inches;
    }
    if (str.includes(':')) {
      const parts = str.split(':');
      return (parseFloat(parts[0]) * 60) + parseFloat(parts[1]);
    }
    return parseFloat(str.replace(/[^\d.]/g, '')) || 0;
  };
  
  // Process each event
  PRINT_EVT.forEach(ev => {
    // Skip relay events
    if (isRelayEvent(ev)) return;
    
    const isField = ev.includes("Jump") || ev.includes("Put") || ev.includes("Discus") || ev.includes("High Jump");
    
    // Get all boys results for this event
    const boysResults = entryData.slice(1)
      .filter(r => r[1] === 'Boys' && r[2] === ev && r[5] && !isNoMark(r[5]))
      .map(r => ({ name: r[3].toString().trim(), mark: r[5] }));
    
    // Get all girls results for this event
    const girlsResults = entryData.slice(1)
      .filter(r => r[1] === 'Girls' && r[2] === ev && r[5] && !isNoMark(r[5]))
      .map(r => ({ name: r[3].toString().trim(), mark: r[5] }));
    
    if (boysResults.length === 0 && girlsResults.length === 0) return; // skip if no results
    
    // Group by athlete and keep only their best mark
    const getBestByAthlete = (results) => {
      const athleteMap = {};
      results.forEach(r => {
        const existing = athleteMap[r.name];
        if (!existing) {
          athleteMap[r.name] = r;
        } else {
          const existingVal = parseVal(existing.mark);
          const newVal = parseVal(r.mark);
          const isBetter = isField ? (newVal > existingVal) : (newVal < existingVal);
          if (isBetter) {
            athleteMap[r.name] = r;
          }
        }
      });
      return Object.values(athleteMap);
    };
    
    const boysBest = getBestByAthlete(boysResults);
    const girlsBest = getBestByAthlete(girlsResults);
    
    // Sort by best mark and take top N athletes
    const sortBoys = boysBest.sort((a, b) => {
      const aVal = parseVal(a.mark);
      const bVal = parseVal(b.mark);
      return isField ? bVal - aVal : aVal - bVal; // field: bigger better, track: smaller better
    }).slice(0, topN);
    
    const sortGirls = girlsBest.sort((a, b) => {
      const aVal = parseVal(a.mark);
      const bVal = parseVal(b.mark);
      return isField ? bVal - aVal : aVal - bVal;
    }).slice(0, topN);
    
    // Event header
    sheet.getRange(row, 1, 1, 4).merge()
      .setValue(ev.toUpperCase())
      .setFontWeight("bold").setFontSize(12)
      .setBackground("#e0e0e0");
    row++;
    
    // Column headers
    sheet.getRange(row, 1).setValue("BOYS").setFontWeight("bold").setBackground("#cfe2f3");
    sheet.getRange(row, 2).setValue("Mark").setFontWeight("bold").setBackground("#cfe2f3");
    sheet.getRange(row, 3).setValue("GIRLS").setFontWeight("bold").setBackground("#f4cccc");
    sheet.getRange(row, 4).setValue("Mark").setFontWeight("bold").setBackground("#f4cccc");
    row++;
    
    // Write top N athletes for both
    const maxRows = Math.max(sortBoys.length, sortGirls.length, topN);
    for (let i = 0; i < maxRows; i++) {
      if (sortBoys[i]) {
        sheet.getRange(row, 1).setValue(sortBoys[i].name);
        sheet.getRange(row, 2).setValue(formatCellValue(sortBoys[i].mark));
      }
      if (sortGirls[i]) {
        sheet.getRange(row, 3).setValue(sortGirls[i].name);
        sheet.getRange(row, 4).setValue(formatCellValue(sortGirls[i].mark));
      }
      row++;
    }
    
    row++; // blank line between events
  });
  
  // Set column widths
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 100);
  
  SpreadsheetApp.getUi().alert('✅ Top Marks report generated!');
}

// ── 6. ATHLETE RECAPS (YEAR-END) ──────────────────────────────

/**
 * Generate year-end recap for all athletes showing meet-by-meet results
 * and year-by-year PR tables. Format matches the PDF template provided.
 */
function generateAllAthleteRecaps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Athlete_Recaps');
  const schedData = ss.getSheetByName('Schedule').getDataRange().getValues();
  const entryData = ss.getSheetByName('Data_Entry').getDataRange().getValues();
  const rosterData = ss.getSheetByName('Roster').getDataRange().getValues();
  const histData = ss.getSheetByName('Historical_PRs').getDataRange().getValues();
  
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), 10).clearNote();
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);
  
  let currentRow = 1;
  
  // Get all unique athletes from roster
  const athletes = rosterData.slice(1).filter(r => r[0]); // has athlete name
  
  if (athletes.length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No athletes found in Roster tab.');
    return;
  }
  
  // Calculate rows needed and ensure sheet has enough rows
  // Estimate ~50 rows per athlete (3-column layout + reduced spacing between athletes)
  const estimatedRowsNeeded = athletes.length * 50;
  const currentMaxRows = sheet.getMaxRows();
  
  if (estimatedRowsNeeded > currentMaxRows) {
    const rowsToAdd = estimatedRowsNeeded - currentMaxRows;
    sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
    SpreadsheetApp.flush();
  }
  
  athletes.forEach((athlete, athleteIdx) => {
    const athleteName = athlete[0].toString().trim();
    const displayName = (athlete[1] || athleteName).toString().trim();
    
    // Check if we're running out of rows (Google Sheets max is ~10000)
    if (currentRow > 9000) {
      SpreadsheetApp.getUi().alert('⚠️ Approaching row limit. Generated ' + athleteIdx + ' of ' + athletes.length + ' athletes.');
      return;
    }
    
    try {
      // Header: School name, Year, Athlete name
      sheet.getRange(currentRow, 1, 1, 8).merge()
        .setValue("Our Lady of Lourdes Track")
        .setFontWeight("bold").setFontSize(14)
        .setHorizontalAlignment("center");
      currentRow++;
      
      sheet.getRange(currentRow, 1, 1, 8).merge()
        .setValue("2026") // current year
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("center");
      currentRow++;
      
      sheet.getRange(currentRow, 1, 1, 8).merge()
        .setValue(displayName)
        .setFontWeight("bold").setFontSize(12)
        .setHorizontalAlignment("center");
      currentRow++;
      currentRow++; // blank line
    
      const startRow = currentRow;
      let col1Row = startRow; // Column 1: cols A-B
      let col2Row = startRow; // Column 2: cols D-E
      
      // Split meets between two columns for better page layout
      const meets = schedData.slice(1).filter(r => r[0]); // has meet #
      const midPoint = Math.ceil(meets.length / 2);
      const col1Meets = meets.slice(0, midPoint);
      const col2Meets = meets.slice(midPoint);
      
      // COLUMN 1: First half of meets (cols A-B)
      col1Meets.forEach(meet => {
        const meetNum = meet[0];
        const meetName = meet[6] || "Meet";
        // Format date properly (don't use formatCellValue for actual dates)
        let meetDate = "";
        if (meet[1]) {
          if (meet[1] instanceof Date) {
            meetDate = Utilities.formatDate(meet[1], Session.getScriptTimeZone(), "M/d/yyyy");
          } else {
            meetDate = meet[1].toString();
          }
        }
        
        // Get this athlete's results for this meet
        const athleteResults = entryData.slice(1).filter(r => {
          const entryName = r[3] ? r[3].toString().trim().toLowerCase() : "";
          const matchName = athleteName.toLowerCase();
          const matchDisplay = displayName.toLowerCase();
          return r[0] == meetNum && (entryName === matchName || entryName === matchDisplay);
        });
        
        // Meet box header
        sheet.getRange(col1Row, 1, 1, 2).merge()
          .setValue(meetName + " " + meetDate)
          .setFontWeight("bold").setFontSize(9)
          .setBorder(true, true, true, true, false, false)
          .setBackground("#d9d9d9");
        col1Row++;
        
        if (athleteResults.length === 0) {
          sheet.getRange(col1Row, 1, 1, 2).merge()
            .setValue("-")
            .setHorizontalAlignment("center")
            .setBorder(true, true, true, true, false, false);
          col1Row++;
        } else {
          // List each event and result
          athleteResults.forEach(res => {
            const event = res[2];
            const isRelay = RELAY_EVTS.includes(event);
            
            // For relays, show both team time and individual split
            let markDisplay = "";
            if (isRelay) {
              const teamTime = res[5] ? formatCellValue(res[5]) : "";
              const split = res[6] ? formatCellValue(res[6]) : "";
              if (teamTime && split) {
                markDisplay = teamTime + " (" + split + ")";
              } else if (teamTime) {
                markDisplay = teamTime;
              } else if (split) {
                markDisplay = split;
              } else {
                markDisplay = "-";
              }
            } else {
              markDisplay = formatCellValue(res[5]);
            }
            
            sheet.getRange(col1Row, 1).setValue(event).setFontSize(8)
              .setBorder(true, true, true, true, false, false);
            sheet.getRange(col1Row, 2).setValue(markDisplay).setFontSize(8)
              .setBorder(true, true, true, true, false, false);
            col1Row++;
          });
        }
        
        col1Row++; // space between meet boxes
      });
      
      // COLUMN 2: Second half of meets (cols D-E)
      col2Meets.forEach(meet => {
        const meetNum = meet[0];
        const meetName = meet[6] || "Meet";
        // Format date properly (don't use formatCellValue for actual dates)
        let meetDate = "";
        if (meet[1]) {
          if (meet[1] instanceof Date) {
            meetDate = Utilities.formatDate(meet[1], Session.getScriptTimeZone(), "M/d/yyyy");
          } else {
            meetDate = meet[1].toString();
          }
        }
        
        // Get this athlete's results for this meet
        const athleteResults = entryData.slice(1).filter(r => {
          const entryName = r[3] ? r[3].toString().trim().toLowerCase() : "";
          const matchName = athleteName.toLowerCase();
          const matchDisplay = displayName.toLowerCase();
          return r[0] == meetNum && (entryName === matchName || entryName === matchDisplay);
        });
        
        // Meet box header
        sheet.getRange(col2Row, 4, 1, 2).merge()
          .setValue(meetName + " " + meetDate)
          .setFontWeight("bold").setFontSize(9)
          .setBorder(true, true, true, true, false, false)
          .setBackground("#d9d9d9");
        col2Row++;
        
        if (athleteResults.length === 0) {
          sheet.getRange(col2Row, 4, 1, 2).merge()
            .setValue("-")
            .setHorizontalAlignment("center")
            .setBorder(true, true, true, true, false, false);
          col2Row++;
        } else {
          // List each event and result
          athleteResults.forEach(res => {
            const event = res[2];
            const isRelay = RELAY_EVTS.includes(event);
            
            // For relays, show both team time and individual split
            let markDisplay = "";
            if (isRelay) {
              const teamTime = res[5] ? formatCellValue(res[5]) : "";
              const split = res[6] ? formatCellValue(res[6]) : "";
              if (teamTime && split) {
                markDisplay = teamTime + " (" + split + ")";
              } else if (teamTime) {
                markDisplay = teamTime;
              } else if (split) {
                markDisplay = split;
              } else {
                markDisplay = "-";
              }
            } else {
              markDisplay = formatCellValue(res[5]);
            }
            
            sheet.getRange(col2Row, 4).setValue(event).setFontSize(8)
              .setBorder(true, true, true, true, false, false);
            sheet.getRange(col2Row, 5).setValue(markDisplay).setFontSize(8)
              .setBorder(true, true, true, true, false, false);
            col2Row++;
          });
        }
        
        col2Row++; // space between meet boxes
      });
      
      // RIGHT COLUMN: Year-by-year PR tables (cols G-H)
      let rightRow = startRow;
      
      // PR table header
      sheet.getRange(rightRow, 7, 1, 2).merge()
        .setValue("Personal Records")
        .setFontWeight("bold").setFontSize(11)
        .setHorizontalAlignment("center");
      rightRow++;
      
      // Current year (2026) from Roster
      sheet.getRange(rightRow, 7).setValue("Event").setFontWeight("bold").setBackground("#e0e0e0");
      sheet.getRange(rightRow, 8).setValue("2026").setFontWeight("bold").setBackground("#e0e0e0");
      rightRow++;
      
      const rosterHeaders = rosterData[0];
      PRINT_EVT.forEach(ev => {
        const evIdx = rosterHeaders.indexOf(ev);
        if (evIdx >= 0) {
          const pr = athlete[evIdx] ? formatCellValue(athlete[evIdx]) : "-";
          if (pr && pr !== "-") {
            sheet.getRange(rightRow, 7).setValue(ev).setFontSize(8);
            sheet.getRange(rightRow, 8).setValue(pr).setFontSize(8);
            rightRow++;
          }
        }
      });
      
      rightRow++; // blank line
      
      // Historical years from Historical_PRs tab
      const histHeaders = histData[0];
      const athleteHistRows = histData.slice(1).filter(r => {
        const hName = (r[0] || '').toString().trim().toLowerCase();
        const hDisplay = (r[1] || '').toString().trim().toLowerCase();
        return hName === athleteName.toLowerCase() || hDisplay === displayName.toLowerCase();
      });
      
      // Get unique years, sorted descending
      const years = [...new Set(athleteHistRows.map(r => r[2]))].sort((a, b) => b - a);
      
      years.forEach(year => {
        const yearRow = athleteHistRows.find(r => r[2] === year);
        if (!yearRow) return;
        
        sheet.getRange(rightRow, 7).setValue("Event").setFontWeight("bold").setBackground("#e0e0e0");
        sheet.getRange(rightRow, 8).setValue(year).setFontWeight("bold").setBackground("#e0e0e0");
        rightRow++;
        
        PRINT_EVT.forEach(ev => {
          const evIdx = histHeaders.indexOf(ev);
          if (evIdx >= 0) {
            const pr = yearRow[evIdx] ? formatCellValue(yearRow[evIdx]) : "-";
            if (pr && pr !== "-") {
              sheet.getRange(rightRow, 7).setValue(ev).setFontSize(8);
              sheet.getRange(rightRow, 8).setValue(pr).setFontSize(8);
              rightRow++;
            }
          }
        });
        
        rightRow++; // blank line between years
      });
      
      // Move to next athlete (force page break at 47 rows per page)
      const nextRow = Math.max(col1Row, col2Row, rightRow);
      const pageSize = 47;
      // Calculate which page we're on (pages: 1-47, 48-94, 95-141, etc.)
      const currentPageEnd = Math.ceil(nextRow / pageSize) * pageSize;
      currentRow = currentPageEnd + 1;
    
    } catch(err) {
      // Log error but continue with next athlete
      Logger.log('Error generating recap for ' + displayName + ': ' + err.message);
    }
    
    // Flush after each athlete to avoid API rate limits
    SpreadsheetApp.flush();
  });
  
  // Set column widths for 3-column layout
  sheet.setColumnWidth(1, 140);  // Meet col 1 - event name
  sheet.setColumnWidth(2, 100);  // Meet col 1 - mark (wider for relay times)
  sheet.setColumnWidth(3, 15);   // gutter
  sheet.setColumnWidth(4, 140);  // Meet col 2 - event name
  sheet.setColumnWidth(5, 100);  // Meet col 2 - mark (wider for relay times)
  sheet.setColumnWidth(6, 15);   // gutter
  sheet.setColumnWidth(7, 140);  // PR - event name
  sheet.setColumnWidth(8, 80);   // PR - mark
  
  SpreadsheetApp.getUi().alert('✅ Athlete Recaps generated for ' + athletes.length + ' athletes!');
}
