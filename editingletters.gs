function generateletters() {createletters();}
function createletters() {
  const templatedoc = "14hRyX6yyfcyh8amy1XR0PNJ6kNcMzOCOyr1xqGpX_vo";
  const folderid = "1O_ZuzZ0ljynipKblZbjMqtnmBnSXH6L2";
  const spreadsheet = SpreadsheetApp.openById("1GG48_TqcS2Y2F2RfV5r-QAeR1pz0wsUZ4HvfFD-rmXo");
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Sheet1"));
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  //folder where letters will be stored
  const folder = DriveApp.getFolderById(folderid);
  for (let i = 1; i < data.length; i++) { //skipping header row
    const rownumber = i + 1; //numbering skips the header row 
    const fullname = data[i][0];  //accessing full name column
    const address = data[i][1];  //address column
    const firstname = data[i][2]; //first name column

    //creating a copy of the template doc
    const title = `${rownumber} ${fullname} donation letter`; //making the title for each letter
    const newdoc = templateFile.makeCopy(`${fullname} letter`, folder);
    const newdocid = newdoc.getId();
    const doc = DocumentApp.openById(newdocid);
    const body = doc.getBody();
    //replacing placeholders with data from the google sheet
    body.replaceText("{{full name}}", fullname || "");
    body.replaceText("{{address}}", address ? address.replace(", ", "\n") : "");
    body.replaceText("{{first name}}", firstname || "");
    //changing title
    doc.setName(title);
    doc.saveAndClose();}
  Logger.log('all docs complete');}
