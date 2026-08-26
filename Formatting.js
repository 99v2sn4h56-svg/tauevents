/************ SHEET FORMATTING / VISUAL POLISH ************/
//
// Same rule as Code.js: every lookup here goes through the header text
// (findHeaderColumn_ / normaliseKey_, both defined in Code.js and shared
// across this Apps Script project), never a hardcoded column number.
//
// Run refreshAllFormatting() once from the Apps Script editor (or use the
// "TAU Tools > Refresh sheet formatting" menu this file adds), and again
// any time a new yearly Requests-style tab is created and needs the same
// look, or whenever the sheet data changes and the masthead stats/date
// should catch up (they're formulas, so they update on their own — this
// is only needed for structure/formatting, not the numbers themselves).

const UPCOMING_EVENTS_SHEET_NAME = 'TAU Upcoming Events';
const STATUS_HEADER = 'Status';
const MASTHEAD_MARKER_ = 'TAU EVENTS';

// Blue / maroon brand palette for the masthead and stat tiles.
const BRAND_MAROON_ = '#5C1B2E';   // banner background
const BRAND_MAROON_DEEP_ = '#3E1220'; // stat number colour (deeper, reads on light bg)
const BRAND_BLUE_ = '#2E5090';     // secondary accent (used sparingly)
const BRAND_BANNER_TEXT_ = '#E7ECF5'; // light blue-white, reads on maroon
const BRAND_TILE_BG_ = '#EEF2F8';  // cool light blue-grey stat tile background
const BRAND_TILE_LABEL_ = '#51607A'; // muted blue-grey label text

const STATUS_STYLES_ = [
  { startsWith: '1.',                 background: '#FCE8B2', color: '#7F5B00', accent: '#D9A400' }, // Request received
  { startsWith: '2.',                 background: '#3C78D8', color: '#FFFFFF', accent: '#3C78D8' }, // Planning
  { startsWith: '3.',                 background: '#4C9959', color: '#FFFFFF', accent: '#3F8F52' }, // Completed
  { startsWith: 'Declined',           background: '#434B54', color: '#FFFFFF', accent: '#6B7280' },
  { startsWith: 'Forwarded contacts', background: '#0F9D6F', color: '#FFFFFF', accent: '#0F9D6F' },
  { startsWith: 'Cancelled',          background: '#A31515', color: '#FFFFFF', accent: '#A31515' }
];


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TAU Tools')
    .addItem('Refresh sheet formatting', 'refreshAllFormatting')
    .addItem('Test Upcoming Events formatting (safe copy)', 'testUpcomingEventsFormattingOnCopy')
    .addToUi();
}


function refreshAllFormatting() {
  formatRequestsWorkflow();
  formatUpcomingEventsSummary();

  SpreadsheetApp.getUi().alert('Formatting refreshed for "Requests" and "' + UPCOMING_EVENTS_SHEET_NAME + '".');
}


/**
 * Tidies the tick-box workflow columns on the Requests sheet:
 * - groups them with a soft background so they read as one block,
 *   separate from the event-info columns either side
 * - centres the checkboxes
 * - tints a box light green the moment it's ticked, so progress across
 *   a row is visible at a glance without reading every header
 * - freezes the header row and the Timestamp/Status columns so status
 *   stays visible while scrolling through the tick boxes
 */
function formatRequestsWorkflow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet named "${REQUESTS_SHEET_NAME}" was not found.`);
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  const startCol = findHeaderColumn_(headers, 'Initial email');
  const endCol = findHeaderColumn_(headers, 'Post-Event Evaluation');

  if (startCol === 0 || endCol === 0 || endCol < startCol) {
    throw new Error(
      'Could not find the workflow tick-box columns by header name ' +
      '("Initial email" ... "Post-Event Evaluation"). Check those headers ' +
      'still exist in row 1 of "Requests" — they may have been renamed.'
    );
  }

  const workflowRange = sheet.getRange(2, startCol, lastRow - 1, endCol - startCol + 1);

  workflowRange
    .setBackground('#F3F6FC')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const firstCellA1 = sheet.getRange(2, startCol).getA1Notation();

  const checkedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=${firstCellA1}=TRUE`)
    .setBackground('#CDEACB')
    .setRanges([workflowRange])
    .build();

  setConditionalFormatRulesForRange_(sheet, workflowRange, [checkedRule]);

  sheet.setFrozenRows(1);

  const statusCol = getOrCreateStatusHeader_(sheet, headers, 2, 1); // confirmed by inspection: column B, header row 1
  sheet.setFrozenColumns(statusCol);
}


/**
 * Turns "TAU Upcoming Events" into the executive-facing landing page:
 * a dark masthead banner + live stat tiles above the table (this is the
 * page higher-ups actually open, so the branding goes here, not on a
 * separate web page), the same colour-coded status chips as the mockup,
 * lighter borders, soft row banding, and frozen header/status column.
 *
 * Idempotent: re-running this does not duplicate the masthead — it
 * detects the existing banner and just refreshes formatting/formulas.
 */
function formatUpcomingEventsSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UPCOMING_EVENTS_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet named "${UPCOMING_EVENTS_SHEET_NAME}" was not found.`);
  }

  applyUpcomingEventsFormatting_(sheet);
}


/**
 * Duplicates "TAU Upcoming Events" and runs the exact same formatting
 * logic against the copy only. Use this to verify the masthead/stat-tile
 * insert works before ever pointing it at the real tab — nothing here
 * touches the live sheet. Delete the copy from the tab bar when done.
 */
function testUpcomingEventsFormattingOnCopy() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const source = spreadsheet.getSheetByName(UPCOMING_EVENTS_SHEET_NAME);

  if (!source) {
    throw new Error(`Sheet named "${UPCOMING_EVENTS_SHEET_NAME}" was not found.`);
  }

  const testName = UPCOMING_EVENTS_SHEET_NAME + ' TEST COPY';
  const existingTestCopy = spreadsheet.getSheetByName(testName);
  if (existingTestCopy) {
    spreadsheet.deleteSheet(existingTestCopy); // start fresh each time this is run
  }

  const copy = source.copyTo(spreadsheet);
  copy.setName(testName);
  spreadsheet.setActiveSheet(copy);

  applyUpcomingEventsFormatting_(copy);

  SpreadsheetApp.getUi().alert(
    'Formatted "' + copy.getName() + '" — check it over. The real "' +
    UPCOMING_EVENTS_SHEET_NAME + '" tab was not touched. Once it looks ' +
    'right, run "Refresh sheet formatting" for real, then delete this ' +
    'test copy tab.'
  );
}


/**
 * The actual masthead + formatting logic, factored out so it can run
 * against either the real sheet or a disposable test copy.
 */
function applyUpcomingEventsFormatting_(sheet) {
  // Clear any frozen rows/columns left over from a previous run before
  // inserting or merging anything — Sheets refuses to merge a range that
  // straddles a frozen/non-frozen column boundary, which is exactly what
  // a full-width banner merge does if column A is already frozen.
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  SpreadsheetApp.flush(); // make sure the freeze change actually lands before merging

  ensureUpcomingEventsMasthead_(sheet);

  const headerRow = findHeaderRow_(sheet, 'Name of Event:');

  if (headerRow === 0) {
    throw new Error(
      'Could not find the header row (looked for "Name of Event:") on "' +
      UPCOMING_EVENTS_SHEET_NAME + '". It may have been renamed or moved.'
    );
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), headerRow + 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];

  const statusCol = getOrCreateStatusHeader_(sheet, headers, 1, headerRow); // confirmed by inspection: column A
  const statusRange = sheet.getRange(headerRow + 1, statusCol, lastRow - headerRow, 1);

  statusRange.setBackground(null).setFontColor(null).setFontWeight('normal').setFontLine('none');

  const statusRules = STATUS_STYLES_.map(style =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(style.startsWith)
      .setBackground(style.background)
      .setFontColor(style.color)
      .setBold(true)
      .setRanges([statusRange])
      .build()
  );

  setConditionalFormatRulesForRange_(sheet, statusRange, statusRules);

  const dataRange = sheet.getRange(headerRow, 1, lastRow - headerRow + 1, lastColumn);
  dataRange.setBorder(false, false, false, false, false, false);
  dataRange.setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
  dataRange.setFontFamily('Public Sans');

  applyRowBanding_(sheet, dataRange);
  applyEventCardStyling_(sheet, headerRow, lastRow, lastColumn, statusCol);

  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(statusCol);
}


/**
 * Makes each data row read as its own event "card" rather than a plain
 * table row: a bold, larger event name, generous auto-fit row height,
 * and a coloured left-edge accent stripe matching that row's status —
 * the closest a spreadsheet grid can get to the mockup's timeline cards
 * without breaking the row-per-event data structure.
 */
function applyEventCardStyling_(sheet, headerRow, lastRow, lastColumn, statusCol) {
  const firstDataRow = headerRow + 1;
  const numDataRows = lastRow - headerRow;

  if (numDataRows <= 0) return;

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const nameCol = findHeaderColumn_(headers, 'Name of Event:');

  if (nameCol > 0) {
    sheet.getRange(firstDataRow, nameCol, numDataRows, 1)
      .setFontSize(12)
      .setFontWeight('bold');
  }

  const statusValues = sheet.getRange(firstDataRow, statusCol, numDataRows, 1).getValues();

  statusValues.forEach((row, i) => {
    const text = String(row[0] || '');
    const style = STATUS_STYLES_.find(s => text.indexOf(s.startsWith) === 0);
    const accent = style ? style.accent : '#D9D9D9';

    sheet.getRange(firstDataRow + i, 1, 1, lastColumn)
      .setBorder(null, true, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  // Guarded: autoResizeRows/getRowHeight availability has varied across
  // Apps Script API versions — a gap here shouldn't break everything else
  // that already applied successfully above.
  try {
    sheet.autoResizeRows(firstDataRow, numDataRows);

    // autoResizeRows tends to run a little tight for wrapped multi-line
    // notes/descriptions — pad each row height slightly so text has room
    // to breathe like a card, not a cramped table cell.
    for (let r = firstDataRow; r < firstDataRow + numDataRows; r++) {
      const current = sheet.getRowHeight(r);
      sheet.setRowHeight(r, current + 14);
    }
  } catch (err) {
    console.error('Row auto-resize skipped: ' + err.message);
  }
}


/**
 * Inserts the dark masthead banner + live stat-tile rows above the
 * existing header, once. Detects an already-inserted masthead via the
 * "TAU EVENTS" marker in A1 so re-running this is a no-op for structure
 * (only the formulas/formatting below get refreshed).
 */
function ensureUpcomingEventsMasthead_(sheet) {
  const existingA1 = String(sheet.getRange(1, 1).getValue() || '');

  if (existingA1.indexOf(MASTHEAD_MARKER_) === 0) {
    return;
  }

  sheet.insertRowsBefore(1, 3);

  const lastColumn = Math.max(sheet.getLastColumn(), 4);
  const headerRow = 4; // just shifted the old row 1 down by 3

  const nameHeaderCol = findHeaderColumn_(
    sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0],
    'Name of Event:'
  );

  const statusColLetter = columnToLetter_(1); // status lives in column A, per formatUpcomingEventsSummary
  const nameColLetter = columnToLetter_(nameHeaderCol > 0 ? nameHeaderCol : 2);
  const firstDataRow = headerRow + 1;

  // --- Row 1: banner ---
  const bannerRange = sheet.getRange(1, 1, 1, lastColumn);
  bannerRange.merge();
  sheet.getRange(1, 1).setFormula(
    '="' + MASTHEAD_MARKER_ + '   ·   Executive Summary — The Arts Unit   ·   As of " & TEXT(NOW(), "ddd d mmm yyyy")'
  );
  bannerRange
    .setBackground(BRAND_MAROON_)
    .setFontColor(BRAND_BANNER_TEXT_)
    .setFontWeight('bold')
    .setFontSize(15)
    .setFontFamily('Fraunces')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // --- Rows 2–3: stat tiles (number, then label underneath) ---
  const groups = splitColumnsIntoGroups_(lastColumn, 4);

  const tiles = [
    { label: 'Upcoming events',   formula: `=SUMPRODUCT(--ISTEXT(${nameColLetter}${firstDataRow}:${nameColLetter}${firstDataRow + 300}))` },
    { label: 'In planning',       formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"2.*")` },
    { label: 'Awaiting response', formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"1.*")` },
    { label: 'Completed',         formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"3.*")` }
  ];

  groups.forEach((group, i) => {
    const tile = tiles[i];
    if (!tile || group.span === 0) return;

    const numberCell = sheet.getRange(2, group.startCol, 1, group.span);
    numberCell.merge();
    sheet.getRange(2, group.startCol).setFormula(tile.formula);
    numberCell
      .setBackground(BRAND_TILE_BG_)
      .setFontColor(BRAND_MAROON_DEEP_)
      .setFontWeight('bold')
      .setFontSize(22)
      .setFontFamily('Fraunces')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0');

    const labelCell = sheet.getRange(3, group.startCol, 1, group.span);
    labelCell.merge();
    sheet.getRange(3, group.startCol).setValue(tile.label);
    labelCell
      .setBackground(BRAND_TILE_BG_)
      .setFontColor(BRAND_TILE_LABEL_)
      .setFontWeight('normal')
      .setFontSize(10)
      .setFontFamily('Public Sans')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('top');
  });

  sheet.setRowHeight(2, 34);
  sheet.setRowHeight(3, 20);
}


/**
 * Splits `totalColumns` into `groupCount` contiguous column groups, as
 * even as possible (extra columns go to the earliest groups). Used to
 * lay out the stat tiles across however many columns the sheet has.
 */
function splitColumnsIntoGroups_(totalColumns, groupCount) {
  const count = Math.min(groupCount, totalColumns);
  const base = Math.floor(totalColumns / count);
  const remainder = totalColumns % count;

  const groups = [];
  let col = 1;

  for (let i = 0; i < count; i++) {
    const span = base + (i < remainder ? 1 : 0);
    groups.push({ startCol: col, span: span });
    col += span;
  }

  while (groups.length < groupCount) {
    groups.push({ startCol: col, span: 0 });
  }

  return groups;
}


/** 1 -> "A", 2 -> "B", 27 -> "AA", etc. */
function columnToLetter_(column) {
  let letter = '';
  let n = column;

  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }

  return letter;
}


/**
 * Scans the first several rows of a sheet for one containing the given
 * header text in any column, and returns that row number (or 0 if not
 * found). Lets formatting code find the real header row even after rows
 * have been inserted above it for a masthead.
 */
function findHeaderRow_(sheet, headerText) {
  const scanRows = Math.min(sheet.getLastRow(), 10);
  if (scanRows === 0) return 0;

  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, scanRows, lastColumn).getValues();
  const target = normaliseKey_(headerText);

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (values[r][c] && normaliseKey_(values[r][c]) === target) {
        return r + 1;
      }
    }
  }

  return 0;
}


/**
 * Finds a column by header text; if the header cell is genuinely blank,
 * labels it "Status" at the given column (only used for the two status
 * columns confirmed by direct inspection to sit there with no header),
 * so every other lookup — here and on future runs — goes through the
 * header name from then on.
 */
function getOrCreateStatusHeader_(sheet, headers, knownColumnIfBlank, headerRow) {
  const existing = findHeaderColumn_(headers, STATUS_HEADER);

  if (existing > 0) {
    return existing;
  }

  const currentValue = headers[knownColumnIfBlank - 1];

  if (currentValue) {
    throw new Error(
      `Expected column ${knownColumnIfBlank} of "${sheet.getName()}" to be the status ` +
      `column but its header is "${currentValue}", not blank or "Status". Please check ` +
      'and re-run.'
    );
  }

  sheet.getRange(headerRow, knownColumnIfBlank).setValue(STATUS_HEADER);
  headers[knownColumnIfBlank - 1] = STATUS_HEADER;

  return knownColumnIfBlank;
}


/**
 * Replaces any existing conditional format rules that target exactly this
 * range with a new set, so re-running formatting functions is idempotent
 * instead of stacking duplicate rules on every run.
 */
function setConditionalFormatRulesForRange_(sheet, range, newRules) {
  const targetA1 = range.getA1Notation();

  const remainingRules = sheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getA1Notation() === targetA1)
  );

  sheet.setConditionalFormatRules(remainingRules.concat(newRules));
}


/**
 * Applies light alternating row shading to a range, replacing any
 * existing banding on the sheet so this stays idempotent.
 */
function applyRowBanding_(sheet, range) {
  range.getBandings().forEach(banding => banding.remove());

  const banding = range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banding.setFirstRowColor('#FFFFFF').setSecondRowColor('#F5F7FA');
}
