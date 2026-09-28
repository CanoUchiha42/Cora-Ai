/**
 * CORA AI — Google Apps Script Lead API
 * Verwendet weiterhin dasselbe Google-Sheets-Dokument und den Tab "Cora Leads".
 */
const SHEET_ID = "1WTw7peU05h0AJ2i2jTkDL14wcsjX6jSFcOXLxXs4cK4";
const SHEET_NAME = "Cora Leads";
const NOTIFICATION_EMAIL = "";

const HEADERS = [
  "Zeitpunkt","Lead-ID","Name","Unternehmen","E-Mail","Telefon",
  "Unternehmensgröße","Leistung","Anliegen","Quelle","Status",
  "Website","Branche","Ziel","Website-Anfragen/Monat","Seite"
];

function doGet(){return json_({ok:true,service:"Cora AI Lead API"});}
function doPost(e){
  try{
    const sheet=getLeadSheet_(); ensureHeaders_(sheet);
    const p=e&&e.parameter?e.parameter:{};
    const leadId=Utilities.getUuid();
    const row=[
      new Date(),leadId,clean_(p.name),clean_(p.company),clean_(p.email),
      clean_(p.phone),clean_(p.company_size),clean_(p.service),clean_(p.message),
      clean_(p.source)||"Cora Website","Neu",clean_(p.website),clean_(p.industry),
      clean_(p.goal),clean_(p.monthly_requests),clean_(p.page)
    ];
    sheet.appendRow(row);
    if(NOTIFICATION_EMAIL)sendNotification_(row);
    return json_({ok:true,leadId:leadId});
  }catch(error){
    console.error(error);
    return json_({ok:false,error:String(error&&error.message?error.message:error)});
  }
}
function getLeadSheet_(){
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName(SHEET_NAME);
  if(!sheet)sheet=ss.insertSheet(SHEET_NAME);
  return sheet;
}
function ensureHeaders_(sheet){
  const first=sheet.getRange(1,1,1,HEADERS.length).getValues()[0];
  const needs=HEADERS.some((h,i)=>first[i]!==h);
  if(needs){sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);sheet.setFrozenRows(1);}
}
function clean_(value){return String(value||"").trim().slice(0,5000);}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function sendNotification_(row){
  MailApp.sendEmail(
    NOTIFICATION_EMAIL,
    "Cora AI — Neuer Lead",
    "Neuer Lead über die Cora Website\n\n"+
    "Name: "+row[2]+"\nUnternehmen: "+row[3]+"\nE-Mail: "+row[4]+"\nTelefon: "+row[5]+
    "\nBranche: "+row[12]+"\nZiel: "+row[13]+"\nAnliegen: "+row[8]+"\nLead-ID: "+row[1]
  );
}