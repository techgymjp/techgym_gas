function postMail() {
  file = DriveApp.getFilesByName('test_file.pdf').next();
  sheet_url = 'スプレッドシートのURL';
  sheet_name = 'シート1';
  spreadsheet = SpreadsheetApp.openByUrl(sheet_url);
  sheet = spreadsheet.getSheetByName(sheet_name);

  subject = sheet.getRange(1, 2).getValue();
  options = {
    bcc: sheet.getRange(2, 2).getValue(),
    name: sheet.getRange(3, 2).getValue(),
    attachments: [file]
  };

  lists = readLists();
  base_body = readBody();

  for(var i in lists){
    body = '';
    body = base_body.replace('[[会社名]]', lists[i]['to_company']);
    body = body.replace('[[名前]]', lists[i]['to_name']);

    GmailApp.sendEmail(lists[i]['to_mail'], subject, body, options);
  }
}

function readBody() {
  doc_url = 'GoogleドキュメントのURL';
  doc = DocumentApp.openByUrl(doc_url);
  return doc.getBody().getText();
}

function readLists() {
  sheet = SpreadsheetApp.getActiveSheet();
  last_row = sheet.getLastRow();

  lists = [];
  for(let i = 0; i < last_row; i++){
    lists[i] = [];
    lists[i]['to_company'] = sheet.getRange(i+1, 1).getValue();
    lists[i]['to_name'] = sheet.getRange(i+1, 2).getValue();
    lists[i]['to_mail'] = sheet.getRange(i+1, 3).getValue();
  }
  return lists;
}

