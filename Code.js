/************ CONFIG ************/
//
// Convention for this file: NEVER read or write the Requests sheet by a
// hardcoded column number. Always look up columns by their header text
// (see findHeaderColumn_ / getValueByNormalisedHeader_ / getOrCreateColumn_
// below) so the script keeps working if columns get reordered or new
// tracking columns get added to the sheet.

const REQUESTS_SHEET_NAME = 'Requests';
const TEMPLATE_SPREADSHEET_ID = '1PfGJGLYH6wpdw18XIKZp0Kc10LbspnVwsiV_AXfqkOM';
const DESTINATION_FOLDER_ID = '1Hq3vMsEEQFHEx-69fEX5kZoZx3soXMjB';
const FILE_NAME_PREFIX = 'Event Request - ';
const STATUS_COLUMN_HEADER = 'Processing Status';

const NOTIFICATION_EMAIL = 'TAUevents@det.nsw.edu.au,Sonja.Benson@det.nsw.edu.au,Lisa.Vandendolder@det.nsw.edu.au';


function setupEventRequestTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createEventRequestSpreadsheet') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('createEventRequestSpreadsheet')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}


function createEventRequestSpreadsheet(e) {
  // Captured first so it's available for error reporting even if
  // everything else below fails.
  const submittedRow = e.range.getRow();

  try {
    const responseSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = responseSpreadsheet.getSheetByName(REQUESTS_SHEET_NAME);

    if (!requestsSheet) {
      throw new Error(`Sheet named "${REQUESTS_SHEET_NAME}" was not found.`);
    }

    const lastColumn = requestsSheet.getLastColumn();

    const headers = requestsSheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0];

    const rowData = requestsSheet
      .getRange(submittedRow, 1, 1, lastColumn)
      .getValues()[0];

    const data = {};

    headers.forEach((header, index) => {
      if (header) {
        data[String(header).trim()] = rowData[index];
      }
    });

    const eventName =
      getValueByNormalisedHeader_(data, 'Name of Event:') ||
      `Submission ${submittedRow}`;

    const eventDate =
      getValueByNormalisedHeader_(data, 'Event Date(s):') ||
      '';

    const formattedEventDate = formatEventDate_(eventDate);

    const folderName = `${eventName} - ${formattedEventDate}`;
    const newFileName = `${FILE_NAME_PREFIX}${eventName}`;

    const templateFile = DriveApp.getFileById(TEMPLATE_SPREADSHEET_ID);
    const parentFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

    const eventFolder = parentFolder.createFolder(folderName);
    const copiedFile = templateFile.makeCopy(newFileName, eventFolder);

    const newSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());

    populateTemplateByMatchingLabels_(newSpreadsheet, data);
    addOriginalRequestTab_(newSpreadsheet, data);

    sendEventRequestEmail_(data, eventFolder, copiedFile);

    writeBackHyperlink_(requestsSheet, headers, submittedRow, eventName, eventFolder.getUrl());
    setRowStatus_(requestsSheet, headers, submittedRow, `Processed ${formatTimestamp_()}`);

  } catch (err) {
    handleProcessingError_(submittedRow, err);
    // Re-throw so this still shows up as a failed execution in the
    // Apps Script Executions log, not just as a quietly-handled error.
    throw err;
  }
}


/**
 * Writes the "Name of Event:" cell in the Requests row as a link to the
 * generated event folder. Column is found by header text, not position.
 */
function writeBackHyperlink_(requestsSheet, headers, submittedRow, eventName, folderUrl) {
  const eventNameColumn = findHeaderColumn_(headers, 'Name of Event:');

  if (eventNameColumn > 0) {
    requestsSheet
      .getRange(submittedRow, eventNameColumn)
      .setFormula(
        `=HYPERLINK("${folderUrl}","${escapeFormulaText_(eventName)}")`
      );
  }
}


/**
 * Writes a human-readable status (success or error) into the
 * STATUS_COLUMN_HEADER column for this row, creating that column by
 * header name if it doesn't exist yet. Never assumes a fixed column
 * number.
 */
function setRowStatus_(requestsSheet, headers, submittedRow, statusText) {
  const statusColumn = getOrCreateColumn_(requestsSheet, headers, STATUS_COLUMN_HEADER);

  requestsSheet
    .getRange(submittedRow, statusColumn)
    .setValue(statusText);
}


/**
 * Finds a column by header text; if it isn't present, appends a new
 * header with that name at the end of row 1 and returns its column
 * index. Keeps the passed-in headers array in sync so repeated calls
 * within the same run don't create duplicate columns.
 */
function getOrCreateColumn_(sheet, headers, headerName) {
  const existingColumn = findHeaderColumn_(headers, headerName);

  if (existingColumn > 0) {
    return existingColumn;
  }

  const newColumn = headers.length + 1;
  sheet.getRange(1, newColumn).setValue(headerName);
  headers.push(headerName);

  return newColumn;
}


/**
 * Central error handler for a failed submission. Leaves a visible error
 * note on the Requests row (best-effort — a failure here must not stop
 * the alert email from going out) and emails the TAU notification list
 * so a failed submission never just disappears into the execution logs.
 */
function handleProcessingError_(submittedRow, err) {
  const message = (err && err.message) || String(err);

  console.error(`Row ${submittedRow} failed to process: ${message}`, err && err.stack);

  try {
    const responseSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = responseSpreadsheet.getSheetByName(REQUESTS_SHEET_NAME);

    if (requestsSheet) {
      const lastColumn = requestsSheet.getLastColumn();
      const headers = requestsSheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0];

      setRowStatus_(requestsSheet, headers, submittedRow, `Error ${formatTimestamp_()}: ${message}`);
    }
  } catch (statusErr) {
    console.error('Additionally failed to write the error status to the sheet:', statusErr);
  }

  sendProcessingFailureEmail_(submittedRow, message, err && err.stack);
}


function sendProcessingFailureEmail_(submittedRow, message, stack) {
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const subject = `⚠️ Event Request processing FAILED - row ${submittedRow}`;

  const body = `An event request submission failed to process automatically and needs manual follow-up.

Row: ${submittedRow}
Sheet: ${spreadsheetUrl}

Error: ${message}
${stack ? '\nDetails:\n' + stack : ''}`;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}


function formatTimestamp_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd MMM yyyy HH:mm'
  );
}


function sendEventRequestEmail_(data, eventFolder, copiedFile) {
  const eventName =
    getValueByNormalisedHeader_(data, 'Name of Event:') ||
    'Untitled Event';

  const eventDate =
    getValueByNormalisedHeader_(data, 'Event Date(s):') ||
    'Date TBC';

  const organisation =
    getValueByNormalisedHeader_(data, 'Name of Organisation:') ||
    '';

  const venue =
    getValueByNormalisedHeader_(data, 'Venue') ||
    '';

  const performanceType =
    getValueByNormalisedHeader_(data, 'Type of Performance') ||
    '';

  const subject = `New Event Request - ${eventName}`;

  const htmlBody = `
    <p>A new event request has been submitted.</p>

    <p>
      <strong>Event Date(s):</strong><br>
      <strong>${escapeHtml_(eventDate)}</strong>
    </p>

    <p>
      <strong>Name of Event:</strong><br>
      <strong>${escapeHtml_(eventName)}</strong>
    </p>

    <p>
      <strong>Name of Organisation:</strong><br>
      <strong>${escapeHtml_(organisation)}</strong>
    </p>

    <p>
      <strong>Venue:</strong><br>
      <strong>${escapeHtml_(venue)}</strong>
    </p>

    <p>
      <strong>Type of Performance:</strong><br>
      <strong>${escapeHtml_(performanceType)}</strong>
    </p>

    <p>
      <strong>Event Folder:</strong><br>
      <a href="${eventFolder.getUrl()}">Open Event Folder</a>
    </p>

    <p>
      <strong>Generated Spreadsheet:</strong><br>
      <a href="${copiedFile.getUrl()}">Open Spreadsheet</a>
    </p>
  `;

  const plainBody = `
A new event request has been submitted.

Event Date(s):
${eventDate}

Name of Event:
${eventName}

Name of Organisation:
${organisation}

Venue:
${venue}

Type of Performance:
${performanceType}

Event Folder:
${eventFolder.getUrl()}

Generated Spreadsheet:
${copiedFile.getUrl()}
`;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}


function formatEventDate_(eventDate) {
  if (!eventDate) {
    return 'Date TBC';
  }

  const date = new Date(eventDate);

  if (isNaN(date.getTime())) {
    return String(eventDate);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMM yyyy'
  );
}


function populateTemplateByMatchingLabels_(spreadsheet, data) {
  const normalisedData = {};

  Object.keys(data).forEach(key => {
    const cleanKey = normaliseKey_(key);
    normalisedData[cleanKey] = data[key];
  });

  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    const values = range.getValues();

    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const cellValue = values[row][col];

        if (!cellValue) continue;

        const cleanLabel = normaliseKey_(cellValue);

        if (normalisedData.hasOwnProperty(cleanLabel)) {
          sheet
            .getRange(row + 1, col + 2)
            .setValue(normalisedData[cleanLabel]);
        }
      }
    }
  });
}


function addOriginalRequestTab_(spreadsheet, data) {
  const existingSheet = spreadsheet.getSheetByName('Original Request');

  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }

  const sheet = spreadsheet.insertSheet('Original Request');

  const rows = Object.entries(data).map(([question, answer]) => [
    question,
    answer
  ]);

  sheet
    .getRange(1, 1, 1, 2)
    .setValues([
      ['Question', 'Response']
    ]);

  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, 2)
      .setValues(rows);
  }

  sheet
    .getRange('A1:B1')
    .setFontWeight('bold')
    .setBackground('#7600f1')
    .setFontColor('#ffffff');

  sheet.autoResizeColumns(1, 2);
  sheet.setFrozenRows(1);
}


function findHeaderColumn_(headers, headerName) {
  const cleanHeaderName = normaliseKey_(headerName);

  for (let i = 0; i < headers.length; i++) {
    if (normaliseKey_(headers[i]) === cleanHeaderName) {
      return i + 1;
    }
  }

  return 0;
}


function getValueByNormalisedHeader_(data, headerName) {
  const cleanHeader = normaliseKey_(headerName);

  for (const key in data) {
    if (normaliseKey_(key) === cleanHeader) {
      return data[key];
    }
  }

  return '';
}


function normaliseKey_(text) {
  return String(text)
    .toLowerCase()
    .replace(/:/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function escapeFormulaText_(text) {
  return String(text).replace(/"/g, '""');
}


function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
