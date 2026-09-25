/**
 * CORA AI — Google Apps Script Lead API
 *
 * Dieses Script kann als eigenständiges Apps-Script-Projekt
 * verwendet werden. Es schreibt direkt in die Cora-Leads-Tabelle.
 */

const SHEET_ID = "1WTw7peU05h0AJ2i2jTkDL14wcsjX6jSFcOXLxXs4cK4";
const SHEET_NAME = "Cora Leads";
const NOTIFICATION_EMAIL = ""; // Optional

const HEADERS = [
  "Zeitpunkt",
  "Lead-ID",
  "Name",
  "Unternehmen",
  "E-Mail",
  "Telefon",
  "Unternehmensgröße",
  "Leistung",
  "Anliegen",
  "Quelle",
  "Status"
];

function doGet() {
  return json_({
    ok: true,
    service: "Cora AI Lead API"
  });
}

function doPost(e) {
  try {
    const sheet = getLeadSheet_();
    ensureHeaders_(sheet);

    const p = e && e.parameter ? e.parameter : {};
    const leadId = Utilities.getUuid();

    const row = [
      new Date(),
      leadId,
      clean_(p.name),
      clean_(p.company),
      clean_(p.email),
      clean_(p.phone),
      clean_(p.company_size),
      clean_(p.service),
      clean_(p.message),
      "Cora Website",
      "Neu"
    ];

    sheet.appendRow(row);

    if (NOTIFICATION_EMAIL) {
      sendNotification_(row);
    }

    return json_({
      ok: true,
      leadId: leadId
    });

  } catch (error) {
    console.error(error);

    return json_({
      ok: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function getLeadSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  const firstRow =
    sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];

  const needsHeaders = HEADERS.some(function(header, index) {
    return firstRow[index] !== header;
  });

  if (needsHeaders) {
    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS]);

    sheet.setFrozenRows(1);
  }
}

function clean_(value) {
  return String(value || "").trim().slice(0, 5000);
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sendNotification_(row) {
  const subject = "Cora AI — Neuer Lead";

  const body =
    "Neuer Lead über die Cora Website\n\n" +
    "Name: " + row[2] + "\n" +
    "Unternehmen: " + row[3] + "\n" +
    "E-Mail: " + row[4] + "\n" +
    "Telefon: " + row[5] + "\n" +
    "Unternehmensgröße: " + row[6] + "\n" +
    "Leistung: " + row[7] + "\n" +
    "Anliegen: " + row[8] + "\n" +
    "Lead-ID: " + row[1];

  MailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
}
