list_file = SpreadsheetApp.openById('<スプレッドシートのID>'); // 各自作成したスプレッドシートのIDを入れる
form_item = list_file.getSheetByName('フォーム入力項目');
seminar_list = list_file.getSheetByName('セミナーリスト');
place_list = list_file.getSheetByName('会場リスト');
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
  folder = makeFolder();
  file = DriveApp.getFileById(invitation.getId());
  copy = file.makeCopy(folder);
  copy.setName('案内状');
  copy_file = DocumentApp.openById(copy.getId());
  insertTable(copy_file, reception_seminars);
  createQR(copy_file);
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
    place_lists[i]['id'] = 
    place_lists[i]['place'] = 
    place_lists[i]['facility'] = 
    place_lists[i]['address'] = 
    place_lists[i]['url'] = 
  }
  return place_lists;
}

function getPlaceID(reception_seminars) {
  place_ids = [];
  reception_seminars.forEach(function(seminar){
    place_ids.push(seminar['place']);
  })
  //重複しないようにするための処理を書く

  return place_ids;
}

function appendMap(place_lists, copy_file) {

}