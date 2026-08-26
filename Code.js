/************ CONFIG ************/

const REQUESTS_SHEET_NAME = 'Requests';
const TEMPLATE_SPREADSHEET_ID = '1PfGJGLYH6wpdw18XIKZp0Kc10LbspnVwsiV_AXfqkOM';
const DESTINATION_FOLDER_ID = '1Hq3vMsEEQFHEx-69fEX5kZoZx3soXMjB';
const FILE_NAME_PREFIX = 'Event Request - ';

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
  const responseSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = responseSpreadsheet.getSheetByName(REQUESTS_SHEET_NAME);

  const submittedRow = e.range.getRow();
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

  const eventNameColumn = findHeaderColumn_(headers, 'Name of Event:');

  if (eventNameColumn > 0) {
    requestsSheet
      .getRange(submittedRow, eventNameColumn)
      .setFormula(
        `=HYPERLINK("${eventFolder.getUrl()}","${escapeFormulaText_(eventName)}")`
      );
  }
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