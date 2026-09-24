const SHEET_ID = "1WTw7peU05h0AJ2i2jTkDL14wcsjX6jSFcOXLxXs4cK4";

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    sheet.appendRow([
      new Date(),
      p.name || "",
      p.company || "",
      p.email || "",
      p.phone || "",
      p.company_size || "",
      p.service || "",
      p.message || ""
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, service: "Cora AI Lead Endpoint" });
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
