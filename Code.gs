function myFunction() {
}

function showSidebar() {
  
  //HTMLファイルを読込み ※引数のファイル名は.htmlを省略可能です。
  var sidebarHtml = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("メール配信");
 
  //SpreadsheetApp上で表示
  SpreadsheetApp.getUi().showSidebar(sidebarHtml);
  
}

function sendMail(mailForm) { //←mailFormというプロパティにsubject、body(共にsidebar.htmlのinputのname)を持つオブジェクトを引数にする

//↓不要になるので削除
//メールデータ
//  var mailForm = {
//    subject : "テスト To ${氏名}さん",
//    body : "Hello ${氏名}さん"
//  };

  //パラメータチェック
  validateMailForm_(mailForm);

  //↓以降は当分そのまま
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

  //↑ここまでそのまま 値を返却するように 修正(ただし未使用)
  return {message : "メール送信が完了しました"};
}

/**
 * Spreadsheet表示時にメニューを追加します。
 * 
 * onOpen関数はGASの中では特殊な関数で、
 * この名前で関数を作成すると"Spreadsheet"起動時に
 * 誰が起動したとしても自動的に呼び出されるようになります。
 */
function onOpen() {

  SpreadsheetApp.getUi()           //UIオブジェクトの取得
  .createMenu("メール配信")    //Spreadsheetに「メール配信」メニューを追加
  .addItem("サイドバーを表示", "showSidebar")  //追加したメニュー内にアイテムを追加、"showSidebar"関数を呼ぶように指定
  .addItem("About", "showAbout")   //Appendix2) Aboudを表示
  .addToUi();                      //実際に"Spreadsheet"へ追加 ※ここを呼ばないと追加されません.
}

/**
 * Aboutダイアログを表示します。
 */
function showAbout() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile("about"), "About this.");
}

/**
 * テストメール送信します。テストメールは利用しているユーザ宛に送信されます。(サイドバーから呼び出されます。)
 * @param {object} mailForm メールの内容 subject:タイトル, body:本文
 * @return {object} 送信結果
 */
function sendTestMail(mailForm) {
  validateMailForm_(mailForm);
  
  //Session.getActiveUser().getEmail()でこのスクリプトを起動したユーザのEmailが取得できる。
  //ただしスクリプトを作成した本人のみ
  var me = Session.getActiveUser().getEmail();
  
  var body = mailForm.body.replace(nameReplaceRegex, me);
  var subject = mailForm.subject.replace(nameReplaceRegex, me);
  
  MailApp.sendEmail(me, subject, body);
  
  SpreadsheetApp.getUi().alert("テストメール送信が完了しました");
  return {message: "テストメール送信が完了しました"};
}

/**
 * メールの内容をチェックします。
 * @param {object} mailForm メールの内容 subject:タイトル, body:本文
 */
//関数名の最後に_(アンダースコア)が付いているとその関数はプライベート関数になります。
//プライベート関数は上部の関数呼び出しSelectBoxに表示されない、google.script.runで呼び出せないなどの特徴があります。
function validateMailForm_(mailForm) {

  if (mailForm.subject == ""){
    throw new Error("メールタイトルは必須です。");
  }
  
  if (mailForm.body == ""){
    throw new Error("メール本文は必須です。");
  }
}

