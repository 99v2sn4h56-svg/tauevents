/************ SHEET FORMATTING / VISUAL POLISH ************/
//
// Same rule as Code.js: every lookup here goes through the header text
// (findHeaderColumn_ / normaliseKey_, both defined in Code.js and shared
// across this Apps Script project), never a hardcoded column number.
//
// Run refreshAllFormatting() once from the Apps Script editor (or use the
// "TAU Tools > Refresh sheet formatting" menu this file adds), and again
// any time a new yearly Requests-style tab is created and needs the same
// look.

const UPCOMING_EVENTS_SHEET_NAME = 'TAU Upcoming Events';
const STATUS_HEADER = 'Status';

const STATUS_STYLES_ = [
  { startsWith: '1.',                 background: '#FCE8B2', color: '#7F5B00' }, // Request received
  { startsWith: '2.',                 background: '#3C78D8', color: '#FFFFFF' }, // Planning
  { startsWith: '3.',                 background: '#4C9959', color: '#FFFFFF' }, // Completed
  { startsWith: 'Declined',           background: '#434B54', color: '#FFFFFF' },
  { startsWith: 'Forwarded contacts', background: '#0F9D6F', color: '#FFFFFF' },
  { startsWith: 'Cancelled',          background: '#A31515', color: '#FFFFFF' }
];


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TAU Tools')
    .addItem('Refresh sheet formatting', 'refreshAllFormatting')
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

  const statusCol = getOrCreateStatusHeader_(sheet, headers, 2); // confirmed by inspection: column B
  sheet.setFrozenColumns(statusCol);
}


/**
 * Makes the "TAU Upcoming Events" tab read as an executive summary rather
 * than a raw data grid:
 * - status column gets the same colour-coded chip look the dropdown
 *   already uses on the Requests sheet (it can't inherit those chips
 *   automatically because these cells are formulas pulling from Requests)
 * - clears the one-off manual colouring that only existed on some rows,
 *   so every row follows the same rule from now on, including new ones
 * - lighter borders and soft row banding instead of a heavy black grid
 * - freezes the header row and status column
 */
function formatUpcomingEventsSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UPCOMING_EVENTS_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet named "${UPCOMING_EVENTS_SHEET_NAME}" was not found.`);
  }

  const lastColumn = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  const statusCol = getOrCreateStatusHeader_(sheet, headers, 1); // confirmed by inspection: column A
  const statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1);

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

  const dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
  dataRange.setBorder(false, false, false, false, false, false);
  dataRange.setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

  applyRowBanding_(sheet, dataRange);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(statusCol);
}


/**
 * Finds a column by header text; if the header cell is genuinely blank,
 * labels it "Status" at the given column (only used for the two status
 * columns confirmed by direct inspection to sit there with no header),
 * so every other lookup — here and on future runs — goes through the
 * header name from then on.
 */
function getOrCreateStatusHeader_(sheet, headers, knownColumnIfBlank) {
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

  sheet.getRange(1, knownColumnIfBlank).setValue(STATUS_HEADER);
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
