function myFunction() {
}

function sendMail() {
  
  //メールデータ
  var mailForm = {
    subject : "テスト To ${氏名}さん",
    body : "Hello ${氏名}さん"
  };
  
  //① Spreadsheetの「メール配信」シートから配信先の取得
  var sheet = SpreadsheetApp.getActive().getSheetByName("メール配信");
  
  //② sheet.getDataRange()はデータが入力されている範囲を取得します。
  //range.getValues()でその範囲に入力されているデータを2次元配列で取得します。
  var values = sheet.getDataRange().getValues();
  
  //1行目はヘッダーなので, データ行は2行目からです。※headerは使いません。
  var header = values[0];
  for(var i = 1; i < values.length; i++) {
    
    //③ 行を取得し、1列目(index = 0)のメールアドレスを取得
    var row = values[i];
    var to = row[0];
    
    //④ 氏名置換
    var nameReplaceRegex = /\$\{氏名\}/gm;
    
    var body = mailForm.body.replace(nameReplaceRegex, row[1]);
    var subject = mailForm.subject.replace(nameReplaceRegex, row[1]);
    
    //メール配信
    MailApp.sendEmail(to, subject, body);
    
    //⑤ 連続で送るとエラーになるので少し待たせます。
    Utilities.sleep(100);    
  }
  
  
  //⑤ 完了した旨をSpreadsheetに表示します。
  SpreadsheetApp.getUi().alert("メール送信が完了しました");
}

function onOpen() {

  SpreadsheetApp.getUi() //①UIオブジェクトの取得
  .createMenu("メール配信") //②Spreadsheetに「メール配信」メニューを追加
  .addItem("メールを送信", "sendMail") //③メニュー内にアイテムを追加、sendMail関数を呼ぶように指定
  .addToUi();  //④実際に"Spreadsheet"へ追加 ※ここを呼ばないと追加されません.

} 
