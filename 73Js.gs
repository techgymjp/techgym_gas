function createForm() {
  form = FormApp.getActiveForm();
  form.setTitle('アンケートツールを作ろう');
  form.setDescription(
    'GASとスプレッドシートを利用して、Googleフォームでアンケートツールを作りました。ご回答くださいますと幸いです。'
  );
  sheet = fetchSheet('質問');
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
    } else if (question['type'].match(/ラジオボタン/)) {
      question['choices'] = updateLimitQuestions(question);
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

function updateLimitQuestions(question) {
  if (!question['type'].match(/制限付き/)) {
    return question['choices'];
  }
  question['choices'].forEach(function (choice, i) {
    limit = choice.replace(/.+残り/, '');
    if (limit == 0) {
      question['choices'].splice(i, 1);
    }
  });
  return question['choices'];
}

function reflectAnswer(e) {
  form = FormApp.getActiveForm();
  answers = e.response.getItemResponses();
  reduceLimitQuestion(form, answers);
  saveAnswers(form, answers);
}

function reduceLimitQuestion(form, answers) {
  form_questions = form.getItems();
  sheet = fetchSheet('質問');
  questions = getQuestions(sheet);
  for (i = 0; i < questions.length; i++) {
    if (!questions[i]['type'].match(/制限付き/)) {
      continue;
    }
    answer = answers[i].getResponse();
    choices = getQuestionChoices(form_questions[i]);
    answer_number = choices.indexOf(answer);
    if (answer_number == -1) {
      continue;
    }
    limit = choices[answer_number].replace(/.+残り/, '');
    limit -= 1;
    replaced_choice = choices[answer_number].replace(
      /残り[1-9]{1,}/,
      '残り' + limit
    );
    choices.splice(answer_number, 1, replaced_choice);
    questions[i]['choices'] = choices;
    questions[i]['choices'] = updateLimitQuestions(questions[i]);
    form_questions[i]
      .asMultipleChoiceItem()
      .setChoiceValues(questions[i]['choices'])
      .setRequired(true);
  }
}

function getQuestionChoices(form_question) {
  choice_values = [];
  if (form_question.getType() == 'MULTIPLE_CHOICE') {
    form_choices = form_question.asMultipleChoiceItem().getChoices();
  } else if (form_question.getType() == 'LIST') {
    form_choices = form_question.asListItem().getChoices();
  }
  form_choices.forEach(function (form_choice) {
    choice = form_choice.getValue();
    choice_values.push(choice);
  });
  return choice_values;
}

function saveAnswers(form, answers) {
  sheet = fetchSheet('回答');
  form_questions = form.getItems();
  sex_choices = getQuestionChoices(form_questions[2]).indexOf(
    answers[2].getResponse()
  );
  area_choices = getQuestionChoices(form_questions[3]).indexOf(
    answers[3].getResponse()
  );
  start_row = 2;
  start_column = 2;
  range = sheet.getRange(area_choices + start_row, sex_choices + start_column);
  range.setValue(range.getValue() + 1);
}

function deleteQuestion(form) {
  form_questions = form.getItems();
  form_questions.forEach(function (form_question) {
    form.deleteItem(form_question);
  });
}
