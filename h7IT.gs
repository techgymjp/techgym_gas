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
  form
    .addTextItem()
    .setTitle('メールアドレスを入力してください。')
    .setRequired(true)
    .setValidation(email_validation);
  form
    .addTextItem()
    .setTitle('年齢を入力してください。')
    .setHelpText('30歳の場合の記入例: 30')
    .setRequired(true);
}

function deleteQuestion(form) {
  form_questions = form.getItems();
  form_questions.forEach(function (form_question) {
    form.deleteItem(form_question);
  });
}
