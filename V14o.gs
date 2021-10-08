function createForm() {
  form = FormApp.getActiveForm();
  form.setTitle('アンケートツールを作ろう');
  form.setDescription(
    'GASとスプレッドシートを利用して、Googleフォームでアンケートツールを作りました。ご回答くださいますと幸いです。'
  );
  sheet = fetchSheet('シート1');
  questions = getQuestions(sheet);
  deleteQuestion(form);
  setQuestion(form, questions);
}

function fetchSheet(sheet_name) {
  spreadsheet = SpreadsheetApp.openById('スプレッドシートのID');
  sheet = spreadsheet.getSheetByName(sheet_name);
  return sheet;
}

function getQuestions(sheet) {
  questions = [];
  question_values = sheet.getDataRange().getValues();
  question_values.shift();
  for (i = 0; i < question_values.length; i++) {
    questions[i] = [];
    questions[i]['title'] = question_values[i][0];
    questions[i]['type'] = question_values[i][1];
    questions[i]['validation'] = question_values[i][2];
    choices = question_values[i].slice(3, question_values[i].length);
    questions[i]['choices'] = choices.filter(isntBlank);
  }
  return questions;
}

function isntBlank(value) {
  return value != '';
}

function setQuestion(form, questions) {
  questions.forEach(function (question) {
    if (question['type'] == '記述式') {
      setTextInputForm(question);
    } else if (question['type'] == 'ラジオボタン') {
      form
        .addMultipleChoiceItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    } else if (question['type'] == 'プルダウン') {
      form
        .addListItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    } else if (question['type'] == 'チェックボックス') {
      form
        .addCheckboxItem()
        .setTitle(question['title'])
        .setChoiceValues(question['choices'])
        .setRequired(true);
    }
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
  } else if (question['validation'].match(/〜/)) {
    help_text = '30歳の場合の記入例: 30';
    error_message = '有効な年齢を入力してください。';
    [min, max] = question['validation'].split('〜');
    validation = FormApp.createTextValidation()
      .requireNumberBetween(min, max)
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
