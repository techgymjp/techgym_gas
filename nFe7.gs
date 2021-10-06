function createForm() {
  form = FormApp.getActiveForm();
  form.setTitle('アンケートツールを作ろう');
  form.setDescription(
    'GASとスプレッドシートを利用して、Googleフォームでアンケートツールを作りました。ご回答くださいますと幸いです。'
  );
  deleteQuestion(form);
  setQuestion(form);
}

function setQuestion(form) {
  email_error_message = '入力されたメールアドレスは有効ではありません。';
  email_validation = FormApp.createTextValidation()
    .requireTextIsEmail()
    .setHelpText(email_error_message)
    .build();
  age_error_message = '有効な年齢を入力してください。';
  age_validation = FormApp.createTextValidation()
    .requireNumberBetween(10, 100)
    .setHelpText(age_error_message)
    .build();
  form
    .addTextItem()
    .setTitle('メールアドレスを入力してください。')
    .setRequired(true)
    .setValidation(email_validation);
  form
    .addTextItem()
    .setTitle('年齢を入力してください。')
    .setHelpText('30歳の場合の記入例: 30')
    .setRequired(true)
    .setValidation(age_validation);
  sex_labels = ['男', '女', 'その他'];
  form
    .addMultipleChoiceItem()
    .setTitle('性別を選んでください。')
    .setChoiceValues(sex_labels)
    .setRequired(true);
  area_labels = [
    '北海道',
    '東北',
    '関東',
    '中部',
    '近畿',
    '中国',
    '四国',
    '九州',
  ];
  form
    .addListItem()
    .setTitle('出身地を選んでください。')
    .setChoiceValues(area_labels)
    .setRequired(true);
  subject_labels = ['国語', '数学', '英語', '理科', '社会'];
  form
    .addCheckboxItem()
    .setTitle('学生時代好きだった教科を選んでください。')
    .setChoiceValues(subject_labels)
    .setRequired(true);
}

function deleteQuestion(form) {
  form_questions = form.getItems();
  form_questions.forEach(function (form_question) {
    form.deleteItem(form_question);
  });
}
