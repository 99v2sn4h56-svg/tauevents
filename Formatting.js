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

// Blue-forward brand palette. Blue is the primary/heading colour; maroon
// is kept only as a minimal secondary accent (a thin rule under the
// banner) rather than a competing second "main" colour. One typeface
// throughout (Public Sans) — hierarchy comes from size/weight/colour,
// not from mixing a display serif in.
const BRAND_BLUE_ = '#2E5090';        // masthead banner background (primary)
const BRAND_BLUE_DEEP_ = '#1F3B66';   // stat tile numbers / accents
const BRAND_MAROON_ = '#5C1B2E';      // secondary accent only — banner underline rule
const BRAND_BANNER_TEXT_ = '#EDF1F9'; // light blue-white, reads on the blue banner
const CARD_CANVAS_BG_ = '#F4F6FB';    // soft backdrop the cards float on, replacing table gridlines
const CARD_TEXT_PRIMARY_ = '#1F2937'; // event name colour — neutral ink, not brand blue

// Status colours are a separate, reserved palette — never the brand
// blue/maroon — so a status is always readable as itself, not as "part
// of the theme." Soft tinted background + matching darker text, like a
// modern tag/chip, rather than a solid saturated fill.
//
// These match against the CLEANED label (numeric prefix like "1. "
// already stripped for display — see buildEventCard_) not the raw
// sheet value, so they key off "Planning" / "Completed" etc.
const STATUS_STYLES_ = [
  { startsWith: 'Request received',   background: '#FCEFCB', color: '#8A5D07' },
  { startsWith: 'Planning',           background: '#DCE6FA', color: '#2E5090' },
  { startsWith: 'Completed',          background: '#DCF0E1', color: '#227A45' },
  { startsWith: 'Declined',           background: '#E7E9ED', color: '#4B5563' },
  { startsWith: 'Forwarded contacts', background: '#D6F3E9', color: '#0F7A57' },
  { startsWith: 'Cancelled',          background: '#FADBDB', color: '#A31515' }
];
const DATE_BADGE_BG_ = '#EAF0FB'; // soft blue tint for the per-event date badge

// Card layout: 3 content rows per event (title+date / venue·org / description)
// plus 1 blank spacer row, so consecutive cards read as separate blocks
// floating on the canvas rather than adjoining table rows.
const CARD_ROWS_PER_EVENT_ = 4;
const CARD_META_KEY_ = 'tauCardBlockRowCount';


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
 * a blue masthead banner + live stat tiles, then every event rendered
 * as its own floating card (coloured status spine, big event name,
 * venue/org line, description line) on a soft grey canvas — not a
 * bordered grid. See rebuildEventCards_ for how the cards are built.
 *
 * Idempotent: re-running this does not duplicate the masthead or the
 * card block — it detects both and rebuilds them fresh in place.
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
 * logic against the copy only. Use this to verify the masthead/card
 * rebuild works before ever pointing it at the real tab — nothing here
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
 * The actual masthead + card-rebuild logic, factored out so it can run
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

  const lastColumn = Math.max(sheet.getLastColumn(), 7);
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const statusCol = getOrCreateStatusHeader_(sheet, headers, 1, headerRow); // confirmed by inspection: column A

  // Full rebuild below, so start from a clean slate rather than trying to
  // reconcile old conditional-format rules against the new card ranges.
  sheet.setConditionalFormatRules([]);
  sheet.setHiddenGridlines(true);

  rebuildEventCards_(sheet, headerRow, statusCol);

  sheet.setFrozenRows(3); // keep the banner + stat tiles visible while scrolling the cards
  sheet.setFrozenColumns(0);
}


/**
 * Rebuilds the event listing as stacked cards instead of a flat table.
 * The original one-row-per-event data (however it gets populated —
 * this doesn't assume or change that) is kept completely intact, just
 * hidden, directly under the masthead; every card is a live formula
 * pointing back at its hidden data row. To add or edit an event: unhide
 * those rows, make the change, then run "Refresh sheet formatting"
 * again to regenerate the cards — the footer note left at the bottom
 * of the card block says the same thing.
 *
 * Idempotent: the row-count of the previously generated card block is
 * stored as sheet developer metadata (survives row inserts elsewhere on
 * the sheet), so re-running this deletes the old card rows before
 * rebuilding rather than stacking duplicates or losing track of where
 * the real data ends and the generated cards begin.
 */
function rebuildEventCards_(sheet, headerRow, statusCol) {
  const previousCardRows = getCardBlockRowCount_(sheet);
  const priorLastRow = sheet.getLastRow();
  const dataLastRow = previousCardRows > 0 ? priorLastRow - previousCardRows : priorLastRow;

  if (previousCardRows > 0) {
    sheet.deleteRows(dataLastRow + 1, previousCardRows);
  }

  const lastColumn = sheet.getLastColumn();
  const dataRowCount = dataLastRow - headerRow;

  if (dataRowCount <= 0) {
    setUpcomingEventsCountTile_(sheet, lastColumn, 0);
    setCardBlockRowCount_(sheet, 0);
    return;
  }

  // Reset then hide the header + data rows — they stay exactly as they
  // are, just out of view, so nothing that reads or writes them by
  // header/column position (this script or anything else) is affected.
  sheet.showRows(headerRow, dataRowCount + 1);
  sheet.hideRows(headerRow, dataRowCount + 1);

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const dateCol = findHeaderColumn_(headers, 'Event Date(s):');
  const nameCol = findHeaderColumn_(headers, 'Name of Event:');
  const venueCol = findHeaderColumn_(headers, 'Venue:');
  const orgCol = findHeaderColumn_(headers, 'Name of Organisation:');
  const descCol = findHeaderColumn_(headers, 'Short Description of Event:');
  const notesCol = findHeaderColumn_(headers, 'TAU Notes');

  if (!dateCol || !nameCol || !venueCol || !orgCol || !descCol) {
    throw new Error(
      'Could not find one or more expected columns ("Event Date(s):", "Name of Event:", ' +
      '"Venue:", "Name of Organisation:", "Short Description of Event:") on "' +
      sheet.getName() + '". Check the headers in row ' + headerRow + ' — they may have been renamed.'
    );
  }

  const cols = {
    statusLetter: columnToLetter_(statusCol),
    dateLetter: columnToLetter_(dateCol),
    nameLetter: columnToLetter_(nameCol),
    venueLetter: columnToLetter_(venueCol),
    orgLetter: columnToLetter_(orgCol),
    descLetter: columnToLetter_(descCol),
    notesLetter: notesCol ? columnToLetter_(notesCol) : null
  };

  // Narrower fixed column widths so the cards read as a contained list,
  // not a full-bleed document — column A is the date badge, B the
  // status tag, C onward the text content.
  sheet.setColumnWidth(1, 88);
  sheet.setColumnWidth(2, 92);
  for (let c = 3; c <= lastColumn; c++) {
    sheet.setColumnWidth(c, 128);
  }

  const totalCardRows = dataRowCount * CARD_ROWS_PER_EVENT_ + 1; // +1 footer instruction row
  sheet.insertRowsAfter(dataLastRow, totalCardRows);

  const tagRanges = [];
  const nameRanges = [];
  const descRows = [];
  let firstCardTitleRow = null;

  for (let i = 0; i < dataRowCount; i++) {
    const dataRow = headerRow + 1 + i;
    const cardTop = dataLastRow + 1 + i * CARD_ROWS_PER_EVENT_;
    const built = buildEventCard_(sheet, cardTop, dataRow, lastColumn, cols);
    tagRanges.push(built.tagRange);
    nameRanges.push(built.nameRange);
    descRows.push(built.descRow);
    if (firstCardTitleRow === null) firstCardTitleRow = built.titleRow;
  }

  // Status tag chip: coloured by matching its own (cleaned) text.
  const tagStatusRules = STATUS_STYLES_.map(style =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(style.startsWith)
      .setBackground(style.background)
      .setFontColor(style.color)
      .setRanges(tagRanges)
      .build()
  );

  // Event-title bar: coloured by matching column B (the status tag) on
  // the SAME row — a formula-based rule, since the title cell's own
  // text is the event name, not the status. The row number here is
  // just the anchor for the first card; Sheets resolves the relative
  // row per range the same way a normal formula fills down.
  const titleStatusRules = STATUS_STYLES_.map(style =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$B${firstCardTitleRow}="${style.startsWith}"`)
      .setBackground(style.color)
      .setFontColor('#FFFFFF')
      .setRanges(nameRanges)
      .build()
  );

  sheet.setConditionalFormatRules(
    sheet.getConditionalFormatRules().concat(tagStatusRules).concat(titleStatusRules)
  );

  // Auto-fit each description row to its actual wrapped content, then put
  // the spacer rows' deliberate height back (auto-resize would otherwise
  // collapse those blank rows to a default height and undo the gaps
  // between cards).
  try {
    descRows.forEach(row => sheet.autoResizeRows(row, 1));
  } catch (err) {
    console.error('Description row auto-resize skipped: ' + err.message);
  }
  for (let i = 0; i < dataRowCount; i++) {
    const spacerRow = dataLastRow + 1 + i * CARD_ROWS_PER_EVENT_ + 3;
    sheet.setRowHeight(spacerRow, 16);
  }

  const footerRow = dataLastRow + 1 + dataRowCount * CARD_ROWS_PER_EVENT_;
  const footerRange = sheet.getRange(footerRow, 1, 1, lastColumn);
  footerRange.setBorder(false, false, false, false, false, false);
  footerRange.merge();
  sheet.getRange(footerRow, 1).setValue(
    '✎  Event details live in hidden rows ' + headerRow + '–' + dataLastRow + ' above — unhide them to add ' +
    'or edit an event, then run TAU Tools → Refresh sheet formatting to update these cards.'
  );
  footerRange
    .setBackground(CARD_CANVAS_BG_)
    .setFontFamily('Public Sans')
    .setFontSize(9)
    .setFontWeight('normal')
    .setFontStyle('italic')
    .setFontColor('#8A93A6')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(footerRow, 22);

  // Soft grey backdrop for any dead space below the card block, so the
  // sheet doesn't trail off into stark white past the last card.
  const maxRows = sheet.getMaxRows();
  if (maxRows > footerRow) {
    sheet.getRange(footerRow + 1, 1, maxRows - footerRow, lastColumn).setBackground(CARD_CANVAS_BG_);
  }

  // The "Upcoming events" tile is written directly here (not left as a
  // formula scanning a wide range) — now that the card rows themselves
  // contain formula-driven text in the same columns as the hidden data,
  // any range-scanning formula double-counts card content alongside the
  // real rows. This function already knows the true count.
  setUpcomingEventsCountTile_(sheet, lastColumn, dataRowCount);

  setCardBlockRowCount_(sheet, totalCardRows);
}


/**
 * Builds one 3-row event card starting at sheet row `top`, sourced
 * entirely from live formulas pointing back at `dataRow` (the hidden
 * flat-table row for this event) — never copied values, so the card
 * always reflects whatever is in the data row as of the last refresh.
 * Every formula is wrapped in TO_TEXT() so a hyperlinked source cell
 * (e.g. an event name Code.js has linked to its generated doc) shows
 * as plain text here — this card controls its own look, not whatever
 * rich formatting happens to live on the source cell.
 *
 * Layout, left to right: a dedicated date badge (its own tinted block,
 * spanning the card's full height — the "clear date card" per event),
 * then a small status tag on the title row, then the event content
 * (name / venue·org / description), all left-aligned like a normal
 * list item rather than centred like a printed page.
 *
 * Returns the status-tag range so the caller can batch all cards' tags
 * into one set of conditional format rules.
 */
function buildEventCard_(sheet, top, dataRow, lastColumn, cols) {
  const titleRow = top;
  const metaRow = top + 1;
  const descRow = top + 2;
  const spacerRow = top + 3;
  const contentCol = 3; // column C onward — column A is the date badge, B the status tag
  const contentWidth = lastColumn - contentCol + 1;

  // insertRowsAfter copies the anchor row's formatting into every new row —
  // borders, bold, italics, whatever the original flat table happened to
  // have. Wipe the whole block back to a blank slate before applying any
  // of this card's own styling, or stray inherited formatting (a grid of
  // black borders, unintentionally-bold text) bleeds straight through.
  const wholeCardRange = sheet.getRange(titleRow, 1, 3, lastColumn);
  wholeCardRange.setBorder(false, false, false, false, false, false);
  wholeCardRange.setFontWeight('normal').setFontStyle('normal');

  // --- Date badge: three stacked lines sharing one tinted block — day +
  // date number (thin, all-caps) on top, month (bold, all-caps) below
  // it, and a live "N days to go" / "N days ago" countdown underneath.
  // Not merged (each line needs its own weight/size), but the shared
  // tint reads as one badge.
  const dateTextExpr = `TO_TEXT(${cols.dateLetter}${dataRow})`;
  const monthExpr = `TRIM(RIGHT(SUBSTITUTE(${dateTextExpr}," ",REPT(" ",100)),100))`;
  const dayDateExpr = `TRIM(SUBSTITUTE(TRIM(LEFT(${dateTextExpr},LEN(${dateTextExpr})-LEN(${monthExpr}))),",",""))`;
  // Drop a leading "Wed , " weekday so DATEVALUE has a clean "26 Aug" to
  // parse; if the source text doesn't match that shape (e.g. a "16-17
  // January 2027" range) this just passes it through unchanged and
  // DATEVALUE below fails safely into a blank countdown via IFERROR.
  const weekdayStrippedExpr = `REGEXREPLACE(${dateTextExpr},"^[A-Za-z]+\\s*,\\s*","")`;
  const dayDiffExpr = `(INT(DATEVALUE(${weekdayStrippedExpr}))-INT(TODAY()))`;

  const dayDateCell = sheet.getRange(titleRow, 1, 1, 1);
  dayDateCell.setFormula(`=UPPER(${dayDateExpr})`);
  dayDateCell
    .setBackground(DATE_BADGE_BG_)
    .setFontFamily('Public Sans')
    .setFontSize(8.5)
    .setFontWeight('normal')
    .setFontStyle('normal')
    .setFontColor(BRAND_BLUE_DEEP_)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('bottom')
    .setWrap(true);

  const monthCell = sheet.getRange(metaRow, 1, 1, 1);
  monthCell.setFormula(`=UPPER(${monthExpr})`);
  monthCell
    .setBackground(DATE_BADGE_BG_)
    .setFontFamily('Public Sans')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontStyle('normal')
    .setFontColor(BRAND_BLUE_DEEP_)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top')
    .setWrap(true);

  const countdownCell = sheet.getRange(descRow, 1, 1, 1);
  countdownCell.setFormula(
    `=IFERROR(IF(${dayDiffExpr}=0,"Today",IF(${dayDiffExpr}>0,${dayDiffExpr}&" day"&IF(${dayDiffExpr}=1,"","s")&" to go",(0-(${dayDiffExpr}))&" day"&IF((0-(${dayDiffExpr}))=1,"","s")&" ago")),"")`
  );
  countdownCell
    .setBackground(DATE_BADGE_BG_)
    .setFontFamily('Public Sans')
    .setFontSize(7.5)
    .setFontWeight('normal')
    .setFontStyle('italic')
    .setFontColor('#5B7CB0')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top')
    .setWrap(true);

  // --- Status tag: small soft-tinted chip, title row only — not a
  // solid block running the height of the card ---
  const tagRange = sheet.getRange(titleRow, 2, 1, 1);
  tagRange.setFormula(`=REGEXREPLACE(TO_TEXT(${cols.statusLetter}${dataRow}),"^\\d+\\.\\s*","")`);
  tagRange
    .setBackground('#FFFFFF')
    .setFontFamily('Public Sans')
    .setFontSize(8.5)
    .setFontWeight('bold')
    .setFontStyle('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // Meta/description rows of the status column stay blank canvas.
  sheet.getRange(metaRow, 2, 2, 1).setBackground(CARD_CANVAS_BG_);

  // --- Title row: event name on a solid, status-coloured bar with white
  // text — an inverse treatment (dark fill / light text) that contrasts
  // against the plain white body below it and gives each card a strong,
  // scannable colour identity per event. Falls back to a neutral dark
  // navy fill until the conditional rule below matches. ---
  const nameRange = sheet.getRange(titleRow, contentCol, 1, contentWidth);
  nameRange.merge();
  sheet.getRange(titleRow, contentCol).setFormula(`=TO_TEXT(${cols.nameLetter}${dataRow})`);
  nameRange
    .setBackground(BRAND_BLUE_DEEP_)
    .setFontFamily('Public Sans')
    .setFontSize(13)
    .setFontWeight('bold')
    .setFontStyle('normal')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // --- Meta row: venue · organisation — a distinct middle tier: still
  // bold so it reads as a label, but smaller and in a muted ink rather
  // than competing with the event name ---
  const metaRange = sheet.getRange(metaRow, contentCol, 1, contentWidth);
  metaRange.merge();
  sheet.getRange(metaRow, contentCol).setFormula(
    `=TO_TEXT(${cols.venueLetter}${dataRow})&IF(AND(${cols.venueLetter}${dataRow}<>"",${cols.orgLetter}${dataRow}<>"")," · ","")&TO_TEXT(${cols.orgLetter}${dataRow})`
  );
  metaRange
    .setFontFamily('Public Sans')
    .setFontSize(9)
    .setFontWeight('bold')
    .setFontStyle('normal')
    .setFontColor('#44546F')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // --- Description row (+ TAU Notes, if present) — the quietest tier:
  // regular weight, italic, lightest ink, clearly supplementary ---
  const descRange = sheet.getRange(descRow, contentCol, 1, contentWidth);
  descRange.merge();
  const notesFormula = cols.notesLetter
    ? `&IF(${cols.notesLetter}${dataRow}<>"",CHAR(10)&"TAU notes: "&TO_TEXT(${cols.notesLetter}${dataRow}),"")`
    : '';
  sheet.getRange(descRow, contentCol).setFormula(`=TO_TEXT(${cols.descLetter}${dataRow})${notesFormula}`);
  descRange
    .setFontFamily('Public Sans')
    .setFontSize(9)
    .setFontWeight('normal')
    .setFontStyle('italic')
    .setFontColor('#8A93A6')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);

  // --- Card body surface: plain white, no border, for the meta/desc
  // rows only — the title row keeps its own status-coloured fill above.
  // White-on-grey contrast against the canvas is what reads as a card,
  // not a ruled box, so this stays free of the "boxed dialog" look ---
  sheet.getRange(metaRow, contentCol, 2, contentWidth).setBackground('#FFFFFF');

  sheet.setRowHeight(titleRow, 26);
  sheet.setRowHeight(metaRow, 20);
  sheet.setRowHeight(descRow, 32);
  sheet.setRowHeight(spacerRow, 16);

  // Spacer row: matches the page backdrop, no border — the whitespace
  // gap between one card and the next.
  sheet.getRange(spacerRow, 1, 1, lastColumn)
    .setBackground(CARD_CANVAS_BG_)
    .setBorder(false, false, false, false, false, false);

  return { tagRange, nameRange, descRow, titleRow };
}


/** Reads the row-count of the previously generated card block, or 0. */
function getCardBlockRowCount_(sheet) {
  const entries = sheet.createDeveloperMetadataFinder().withKey(CARD_META_KEY_).find();
  if (entries.length === 0) return 0;
  return parseInt(entries[0].getValue(), 10) || 0;
}


/** Stores the row-count of the just-generated card block, replacing any previous entry. */
function setCardBlockRowCount_(sheet, count) {
  sheet.createDeveloperMetadataFinder().withKey(CARD_META_KEY_).find()
    .forEach(entry => entry.remove());
  sheet.addDeveloperMetadata(CARD_META_KEY_, String(count));
}


/**
 * Inserts the blue masthead banner + live stat-tile rows above the
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

  const statusColLetter = columnToLetter_(1); // status lives in column A, per formatUpcomingEventsSummary
  const firstDataRow = headerRow + 1;

  // --- Row 1: banner (blue — the primary heading colour) ---
  const bannerRange = sheet.getRange(1, 1, 1, lastColumn);
  bannerRange.merge();
  sheet.getRange(1, 1).setFormula(
    '="' + MASTHEAD_MARKER_ + '   ·   Executive Summary — The Arts Unit   ·   As of " & TEXT(NOW(), "ddd d mmm yyyy")'
  );
  bannerRange
    .setBackground(BRAND_BLUE_)
    .setFontColor(BRAND_BANNER_TEXT_)
    .setFontWeight('bold')
    .setFontSize(13)
    .setFontFamily('Public Sans')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(null, null, true, null, null, null, BRAND_MAROON_, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 34);

  // --- Rows 2–3: stat tiles — small bold ALL-CAPS label under a coloured
  // top rule, then the number below in a large weight. Each label/rule
  // is tinted a different accent so the row itself carries some colour
  // variety instead of reading as one flat blue block. "Upcoming events"
  // is left without a formula here — rebuildEventCards_ writes the real
  // count directly once it knows it (see setUpcomingEventsCountTile_).
  const groups = splitColumnsIntoGroups_(lastColumn, 4);

  const tiles = [
    { label: 'Upcoming events',   accent: BRAND_BLUE_DEEP_, formula: null },
    { label: 'In planning',       accent: '#2E5090', formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"2.*")` },
    { label: 'Awaiting response', accent: '#8A5D07', formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"1.*")` },
    { label: 'Completed',         accent: '#227A45', formula: `=COUNTIF(${statusColLetter}${firstDataRow}:${statusColLetter},"3.*")` }
  ];

  groups.forEach((group, i) => {
    const tile = tiles[i];
    if (!tile || group.span === 0) return;

    const labelCell = sheet.getRange(2, group.startCol, 1, group.span);
    labelCell.merge();
    sheet.getRange(2, group.startCol).setValue(tile.label.toUpperCase());
    labelCell
      .setBackground('#FFFFFF')
      .setFontColor(tile.accent)
      .setFontWeight('bold')
      .setFontStyle('normal')
      .setFontSize(8.5)
      .setFontFamily('Public Sans')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('bottom')
      .setBorder(true, null, null, null, null, null, tile.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    const numberCell = sheet.getRange(3, group.startCol, 1, group.span);
    numberCell.merge();
    if (tile.formula) {
      sheet.getRange(3, group.startCol).setFormula(tile.formula);
    }
    numberCell
      .setBackground('#FFFFFF')
      .setFontColor(BRAND_BLUE_DEEP_)
      .setFontWeight('bold')
      .setFontStyle('normal')
      .setFontSize(19)
      .setFontFamily('Public Sans')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('top')
      .setNumberFormat('0');
  });

  sheet.setRowHeight(2, 20);
  sheet.setRowHeight(3, 30);
}


/**
 * Writes the true "Upcoming events" count directly into its tile cell.
 * Kept separate from a formula because the card rows below now put
 * formula-driven text in the same columns the hidden data lives in, so
 * a range-scanning COUNTA/ISTEXT formula would double-count the cards
 * themselves alongside the real data.
 */
function setUpcomingEventsCountTile_(sheet, lastColumn, count) {
  const groups = splitColumnsIntoGroups_(lastColumn, 4);
  const first = groups[0];
  if (!first || first.span === 0) return;
  sheet.getRange(3, first.startCol).setValue(count);
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
