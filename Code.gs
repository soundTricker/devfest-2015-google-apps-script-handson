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

function onOpen() {

  SpreadsheetApp.getUi() //①UIオブジェクトの取得
  .createMenu("メール配信") //②Spreadsheetに「メール配信」メニューを追加
  .addItem("サイドバーを表示", "showSidebar") //③メニュー内にアイテムを追加、showSidebar関数を呼ぶように指定
  .addToUi();  //④実際に"Spreadsheet"へ追加 ※ここを呼ばないと追加されません.

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

