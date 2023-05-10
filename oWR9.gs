list_file = SpreadsheetApp.openById('<スプレッドシートのID>'); // 各自作成したスプレッドシートのIDを入れる
form_item = list_file.getSheetByName('フォーム入力項目');
seminar_list = list_file.getSheetByName('セミナーリスト');
place_list = list_file.getSheetByName('会場リスト');
client_list = list_file.getSheetByName('顧客リスト');
entry_list = list_file.getSheetByName('参加者リストフォーマット');
invitation = DocumentApp.openByUrl('<ドキュメントのURL>');

// セミナー参加の申し込みフォームを作成しよう
function createForm() {
  form = FormApp.getActiveForm();
  form.setTitle('2022年5月セミナー申し込みフォーム');
  form.setDescription(
    '参加希望の方は下記の項目にご回答ください'
  );
  questions = getQuestions();
  deleteQuestion(form);
  setQuestion(questions);
}

function getQuestions() {
  questions = [];
  question_values = form_item.getDataRange().getValues();
  question_values.shift();
  for (i = 0; i < question_values.length; i++) {
    questions[i] = [];
    questions[i]['title'] = question_values[i][0];
    questions[i]['type'] = question_values[i][1];
    questions[i]['validation'] = question_values[i][2];
  }
  return questions;
}

function setQuestion(questions) {
  questions.forEach(function (question) {
    setTextInputForm(question);
  });
}

function setTextInputForm(question) {
  if (question['validation'] == 'メールアドレス') {
    help_text = '';
    error_message = '入力されたメールアドレスは有効ではありません。';
    validation = FormApp.createTextValidation()
      .requireTextIsEmail()
      .setHelpText(error_message)
      .build();
  } else if (question['validation'] == '') {
    help_text = '';
    error_message = '';
    validation = FormApp.createTextValidation()
      .setHelpText(error_message)
      .build();
  }
  form
    .addTextItem()
    .setTitle(question['title'])
    .setHelpText(help_text)
    .setRequired(true)
    .setValidation(validation);
}

function deleteQuestion(form) {
  form_questions = form.getItems();
  form_questions.forEach(function (form_question) {
    form.deleteItem(form_question);
  });
}

// 第1節　申し込みフォームに項目を追加しよう
function readSeminarLists() {
  seminar_lists = [];
  for(i = 0; i < seminar_list.getLastRow() - 1; i++){
    seminar_lists[i] = [];
    seminar_lists[i]['id'] = seminar_list.getRange(i + 2, 1).getValue();
    seminar_lists[i]['date'] = seminar_list.getRange(i + 2, 2).getValue();
    seminar_lists[i]['time'] = seminar_list.getRange(i + 2, 3).getValue();
    seminar_lists[i]['theme'] = seminar_list.getRange(i + 2, 4).getValue();
    seminar_lists[i]['place'] = seminar_list.getRange(i + 2, 5).getValue();
    seminar_lists[i]['status'] = seminar_list.getRange(i + 2, 6).getValue();
  }
  return seminar_lists;
}

function getReceptionSeminars(seminar_lists){
  reception_seminars = seminar_lists.filter(seminar => {
    if(seminar.status === '受付開始'){
      return true;
    }
  });
  return reception_seminars;
}

function setSeminarInfo() {
  seminar_lists = readSeminarLists();
  reception_seminars = getReceptionSeminars(seminar_lists);
  seminars = [];
  reception_seminars.forEach(function(seminar){
    seminars.push('【ID】' + seminar['id'] + ' 【日時】' + Utilities.formatDate(seminar['date'], 'JST', 'M月d日') + seminar['time'] + ' 【テーマ】' + seminar['theme'] + ' 【会場】' +  place_list.getRange(seminar['place'] + 1, 2).getValue());
  })
  return seminars;
}

function addItem() {
  form = FormApp.getActiveForm();
  seminars = setSeminarInfo(); 
  form
    .addCheckboxItem()
    .setTitle('参加セミナー')
    .setChoiceValues(seminars)
    .setRequired(true);
}


// 第2節　セミナー情報を表にまとめて案内状に記載しよう
function setTableTitles() {
  table_titles = [];
  table_titles.push('日程');
  table_titles.push('時間');
  table_titles.push('テーマ');
  table_titles.push('会場');
  return table_titles;
}

function setTableContents(reception_seminars) {
  table_contents = [];
  reception_seminars.forEach(function(seminar){
    rows = [];
    rows.push(Utilities.formatDate(seminar['date'], 'JST', 'M月d日'));
    rows.push(seminar['time']);
    rows.push(seminar['theme']);
    rows.push(place_list.getRange(seminar['place'] + 1, 2).getValue());
    table_contents.push(rows);
    rows = [];
  })
  return table_contents;
}

function insertTable(copy_file, reception_seminars) {
  table_titles = setTableTitles();
  table_contents = setTableContents(reception_seminars);
  table = [];
  table.push(table_titles);
  table_contents.forEach(function(content){
    table.push(content);
  })
  body = copy_file.getBody();
  body.insertTable(10, table);
}

function makeFolder(){
  folder_name = 'セミナー案内状';
  folder_id = DriveApp.createFolder(folder_name).getId();
  folder = DriveApp.getFolderById(folder_id);
  return folder;
}

function makeInvitation() {
  seminar_lists = readSeminarLists();
  reception_seminars = getReceptionSeminars(seminar_lists);
  place_lists = readPlaceLists();
  place_id = getPlaceID(reception_seminars);
  client_lists = readClientLists();
  folder = makeFolder();
  for (i = 0; i < client_lists.length; i++){
    file = DriveApp.getFileById(invitation.getId());
    copy = file.makeCopy(folder);
    file_name = '【案内状】' + client_lists[i]['company'] + client_lists[i]['name'] + '様';
    copy.setName(file_name);
    copy_file = DocumentApp.openById(copy.getId());
    insertTable(copy_file, reception_seminars);
    createQR(copy_file);
    appendMap(place_lists,client_lists, i, copy_file);
    replaceBody(i, copy_file);
  }
}

// 第3節　申し込みフォームのQRコードを作成しよう
function createQR(copy_file) {
  form_url = FormApp.getActiveForm().getPublishedUrl();
  qr = UrlFetchApp.fetch('http://chart.apis.google.com/chart?cht=qr&chs=150x150&chl=' + form_url);
  image = qr.getBlob().setName('申し込みフォーム');
  copy_file.insertImage(19, image);
}

// 第4節　会場付近の地図を案内状に挿入しよう
function readPlaceLists() {
  place_lists = [];
  for(i = 0; i < place_list.getLastRow() - 1 ; i++){
    place_lists[i] = [];
    place_lists[i]['id'] = place_list.getRange(i + 2, 1).getValue();
    place_lists[i]['place'] = place_list.getRange(i + 2, 2).getValue();
    place_lists[i]['facility'] = place_list.getRange(i + 2, 3).getValue();
    place_lists[i]['address'] = place_list.getRange(i + 2, 4).getValue();
    place_lists[i]['url'] = place_list.getRange(i + 2, 5).getValue();
  }
  return place_lists;
}

function getPlaceID(reception_seminars) {
  place_ids = [];
  reception_seminars.forEach(function(seminar){
    place_ids.push(seminar['place']);
  })
  place_ids = place_ids.filter(function(id, index, array){
    return array.indexOf(id) === index;
  });
  return place_ids;
}

function appendMap(place_lists,client_lists,i, copy_file) {
  place_id.forEach(function(id){
    seminar_address = place_lists[id - 1]['address'];
    map = Maps.newStaticMap()
      .setLanguage('ja')
      .setSize(600,300)
      .setZoom(16)
      .setCenter(seminar_address)
      .addMarker(seminar_address);
    copy_file.appendParagraph('■' + place_lists[id -1]['place'] +'会場（' + place_lists[id - 1]['facility'] + '）' +seminar_address);
    copy_file.appendImage(map);
    getAccess(place_lists,client_lists, id, i, copy_file);
  })
}

// 第5節　顧客の最寄駅から会場までの所要時間と交通費を表示しよう
function readClientLists() {
  client_lists = [];
  for(i = 0; i < client_list.getLastRow() - 1; i++){
    client_lists[i] = [];
    client_lists[i]['company'] = client_list.getRange(i + 2, 1).getValue();
    client_lists[i]['name'] = client_list.getRange(i + 2, 2).getValue();
    client_lists[i]['address'] = client_list.getRange(i + 2, 3).getValue();
    client_lists[i]['sales'] = client_list.getRange(i + 2, 4).getValue();
  }
  return client_lists;
}

function getAccess(place_lists,client_lists, id, i, copy_file) {
  seminar_address = place_lists[id - 1]['address'];
  client_address = client_lists[i]['address'];
  directions = Maps.newDirectionFinder()
                  .setLanguage('ja')
                  .setOrigin(client_address)
                  .setDestination(seminar_address)
                  .setMode(Maps.DirectionFinder.Mode.TRANSIT)
                  .setArrive(new Date(2023, 4, 1, 10))
                  .getDirections();
  route = directions.routes[0];
  if (route['fare']) {
    access = '貴社の最寄駅からの所要時間' + route.legs[0].duration.text + '（交通費' + route.fare.text + '）'
    copy_file.appendParagraph(access);
  } else {
    access = '貴社の最寄駅からの所要時間' + route.legs[0].duration.text + '（交通費¥0）'
    copy_file.appendParagraph(access);
  }
}

// 第6節　顧客毎に案内状の文章を変えよう
function replaceBody(i, copy_file) {
  company = client_lists[i]['company'];
  name = client_lists[i]['name'];
  sales = client_lists[i]['sales'];
  form_url = FormApp.getActiveForm().getPublishedUrl();
  body = copy_file.getBody();
  body = body.replaceText('{会社名}', company);
  body = body.replaceText('{氏名}', name);
  body = body.replaceText('{担当}', sales);
  body = body.replaceText('{申し込みフォーム}', form_url);
}

// 第7節　顧客分の案内状を作成し、PDF化しよう
function getFolderID(){
  folder = DriveApp.getFoldersByName('セミナー案内状');
  folder_id = folder.next().getId();  
  return folder_id;
}

function getFileInfo(){
  folder_id = getFolderID();
  files = DriveApp.getFolderById(folder_id).getFiles();
  file_info = [];
  while(files.hasNext()){
    file = files.next();
    file_info.push([file.getId(), file.getName()]);
  }
  return file_info;
}

function makePDF() {
  folder_id = getFolderID();
  file_info = getFileInfo();
  file_info.forEach(function(info){
    url = 'https://docs.google.com/document/d/'+ info[0] +'/export?';
    opts = {
    exportFormat: 'pdf',
    format:       'pdf',
    size:         'A4',
    portrait:     'true',
    };
    urlExts = [];
    for(optName in opts){
      urlExts.push(optName + '=' + opts[optName]);
    }
    options  = urlExts.join('&');
    token    = ScriptApp.getOAuthToken();
    response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });
    blob = response.getBlob().setName(info[1] + '.pdf');
    DriveApp.getFolderById(folder_id).createFile(blob);
  })
}

// 第8節　参加者管理のためのシートを複製しよう
function getSeminarID(seminar_lists) {
  seminar_ids = [];
  seminar_lists = readSeminarLists()
  seminar_lists.forEach(function(seminar){
    seminar_ids.push(seminar['id']);
  })
  return seminar_ids;
}

function copyEntryList() {
  seminar_lists = readSeminarLists();
  seminar_id = getSeminarID(seminar_lists);
  reception_seminars = getReceptionSeminars(seminar_lists);
  reception_seminars.forEach(function(seminar){
    seminar_name = Utilities.formatDate(seminar['date'], 'JST', 'MMdd') + seminar['theme'] + place_list.getRange(seminar['place'] + 1, 2).getValue();
    newSheet = entry_list.copyTo(list_file);
    newSheet.setName(seminar_name);
    seminar_list.getRange(seminar_id.indexOf(seminar['id']) + 2, 7).setValue(list_file.getUrl() + '#gid=' + newSheet.getSheetId());  
  })
}

// 第9節　参加申し込みがあったら、参加者リストへ追加しよう
function addEntryList(e) {
  seminar_lists = readSeminarLists();
  seminar_id = getSeminarID(seminar_lists);
  itemResponses = e.response.getItemResponses();
  entry_seminar = itemResponses[3].getResponse();
  before = '【ID】';
  after = ' 【日時】';
  entry_seminar.forEach(function(seminar){
    regexp = new RegExp( '(?<=' + before + ').*?(?=' + after + ')' );
    id = Number(seminar.match(regexp)[0]);
    entry_url = seminar_list.getRange(seminar_id.indexOf(id) + 2, 7).getValue();
    sheets = list_file.getSheets();
    sheets.forEach(function(sheet){
      if (entry_url.match(sheet.getSheetId()) != null){
        sheet_number = sheets.indexOf(sheet);
        entry_list = SpreadsheetApp.openByUrl(entry_url).getSheets()[sheet_number];
        last_row = entry_list.getLastRow();
        for (i = 0; i < itemResponses.length - 1; i++) {
          entry_list.getRange(last_row + 1, 1 + i).setValue(itemResponses[i].getResponse());
        }
      }
    })
  })
}

// 第10節　申し込みを受け付けた旨をメールで通知しよう
function sendThanksMail(e) {
  place_lists = readPlaceLists();
  itemResponses = e.response.getItemResponses();
  entry_seminar = itemResponses[3].getResponse();
  recipient = itemResponses[2].getResponse();
  subject = '申し込み完了のお知らせ';
  body = '';
  body += '以下の内容で申し込みを受け付けました\n';
  body += '会社名：' + itemResponses[0].getResponse() + '\n';
  body += '氏名：' +  itemResponses[1].getResponse() + ' 様\n';
  for (i = 0;  i < itemResponses[3].getResponse().length; i++){
    body += '参加セミナー' + (i + 1)  +'：' +  itemResponses[3].getResponse()[i] + '\n';
  }
  body += '\n';
  body += '各会場へのアクセスは以下の通りです\n';
  place_lists.forEach(function(place){
    body += place['place'] + '会場：' + place['url'] + '\n';
  })
  GmailApp.sendEmail(recipient, subject, body);
}
