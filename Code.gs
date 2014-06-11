function myFunction() {
}

function sendMail() {
  var email = Session.getActiveUser().getEmail();
  
  MailApp.sendEmail(email , "初めてのGoogle Apps Script", "Hello World. Welcome Google Apps Script!");
}

function onOpen() {

  SpreadsheetApp.getUi() //①UIオブジェクトの取得
  .createMenu("メール配信") //②Spreadsheetに「メール配信」メニューを追加
  .addItem("メールを送信", "sendMail") //③メニュー内にアイテムを追加、sendMail関数を呼ぶように指定
  .addToUi();  //④実際に"Spreadsheet"へ追加 ※ここを呼ばないと追加されません.

} 
