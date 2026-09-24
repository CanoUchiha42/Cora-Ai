/**
 * CORA AI – Google-Formular + Google-Sheets Lead-Erfassung
 *
 * Anleitung:
 * 1. Öffne https://script.google.com/
 * 2. Neues Projekt erstellen.
 * 3. Diesen Code einfügen und "setupCoraLeadForm" ausführen.
 * 4. Google-Berechtigungen bestätigen.
 * 5. Danach werden Formular-URL, Bearbeitungs-URL und Tabellen-URL ausgegeben.
 *
 * Die Antworten landen automatisch in Google Sheets.
 */

function setupCoraLeadForm() {
  const form = FormApp.create('Cora AI – Anfrage / Beratung');

  form.setDescription(
    'Kontaktanfrage für Cora AI. Bitte füllen Sie die folgenden Felder aus. ' +
    'Ihre Angaben werden zur Bearbeitung Ihrer Anfrage verarbeitet.'
  );
  form.setConfirmationMessage(
    'Vielen Dank. Ihre Anfrage ist eingegangen. Cora AI meldet sich zeitnah bei Ihnen.'
  );
  form.setCollectEmail(false);
  form.setProgressBar(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);

  form.addTextItem()
    .setTitle('Name')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Unternehmen')
    .setRequired(true);

  form.addTextItem()
    .setTitle('E-Mail-Adresse')
    .setRequired(true);

  form.addTextItem()
    .setTitle('Telefonnummer')
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle('Unternehmensgröße')
    .setChoiceValues([
      '1–10 Mitarbeitende',
      '11–50 Mitarbeitende',
      '51–250 Mitarbeitende',
      '251–1.000 Mitarbeitende',
      '1.001+ Mitarbeitende'
    ])
    .setRequired(false);

  form.addMultipleChoiceItem()
    .setTitle('Gewünschte Leistung')
    .setChoiceValues([
      'Cora Basic',
      'Cora Pro',
      'Cora Enterprise',
      'Beratung / individuelle Lösung'
    ])
    .setRequired(false);

  form.addParagraphTextItem()
    .setTitle('Anliegen')
    .setRequired(true);

  // Antwort-Tabelle in Google Sheets erstellen und mit dem Formular verbinden.
  const sheet = SpreadsheetApp.create('Cora AI – Leads');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  Logger.log('=== CORA AI LEAD-FORMULAR ERSTELLT ===');
  Logger.log('Formular zum Ausfüllen: ' + form.getPublishedUrl());
  Logger.log('Formular bearbeiten: ' + form.getEditUrl());
  Logger.log('Google-Sheets-Leads: ' + sheet.getUrl());
  Logger.log('Formular-ID: ' + form.getId());
  Logger.log('======================================');
}
