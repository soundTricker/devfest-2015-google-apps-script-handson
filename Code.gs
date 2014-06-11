
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
 * サイドバーを表示します。
 */
function showSidebar() {
  
  //Appendix 5)
  //前回の送信で保存したメールタイトル、本文を取得
  var userProp = PropertiesService.getUserProperties();
  
  
  var prevMailForm = {
    subject : userProp.getProperty("subject") || "",
    body : userProp.getProperty("body") || ""
  };
  
//  //HTMLファイルを読込み ※引数のファイル名は.htmlを省略可能です。
//  var sidebarHtml = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("メール配信");

  //Appendix 5)
  //sidebarをテンプレートして扱う
  var sidebarHtml = HtmlService.createTemplateFromFile("sidebar");
  
  sidebarHtml.prevMailForm = prevMailForm;

  //SpreadsheetApp上で表示
  SpreadsheetApp.getUi().showSidebar(sidebarHtml.evaluate().setTitle("メール配信"));
  
}
/**
 * メール送信します。(サイドバーから呼び出されます。)
 * @param {object} mailForm メールの内容 subject:タイトル, body:本文
 * @return {object} 送信結果
 */
function sendMail(mailForm) {

  //パラメータチェック
  validateMailForm_(mailForm);
  
  //Spreadsheetの「メール配信」シートから配信先の取得
  var sheet = SpreadsheetApp.getActive().getSheetByName("メール配信");
  
  //sheet.getDataRange()はデータが入力されている範囲を取得します。
  var dataRange = sheet.getDataRange();
  
  //range.getValues()でその範囲に入力されているデータを2次元配列で取得します。
  var values = dataRange.getValues();
  
  //1行目はヘッダーなので, データ行は2行目からです。
  var header = values[0];
  for(var i = 1; i < values.length; i++) {
    
    var row = values[i];
    var to = row[0];
    
    //Appendix4) 置換 
    var body = mailForm.body + "";
    var subject = mailForm.subject + "";
    
    //ヘッダーのカラム名で全て置換
    for(var colIndex = 0; colIndex < header.length; colIndex++) {
      var replaceRegex = new RegExp("\\$\\{" + header[colIndex] + "\\}", "gm");
      body = body.replace(replaceRegex, row[colIndex]);
      subject = subject.replace(replaceRegex, row[colIndex]);
    }
    
    //メール配信
    MailApp.sendEmail(to, subject, body);
    
    //現在処理中の行の末尾のカラム2列に送信完了日とステータスを入れる Appendix 3)
    //getRangeで使うrow,columnは1から始まりです。
    sheet.getRange(i + 1, row.length - 1, 1 , 2).setValues([["済", new Date()]]);
    
    //連続で送るとエラーになるので少し待たせます。
    Utilities.sleep(100);
  }
  
  
  //完了した旨をSpreadsheetに表示します。
  SpreadsheetApp.getUi().alert("メール送信が完了しました");
  
  //今回のメールタイトル、本文を保存 Appendix 5)
  var userProp = PropertiesService.getUserProperties();
  userProp.setProperty("subject", mailForm.subject);
  userProp.setProperty("body", mailForm.body);
  
  return {message: "メール送信が完了しました"};
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