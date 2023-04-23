list_file = SpreadsheetApp.openById('<スプレッドシートのID>'); // 各自作成したスプレッドシートのIDを入れる
form_item = list_file.getSheetByName('フォーム入力項目');

// 0節　セミナー参加の申し込みフォームを作成しよう
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