function myFunction() {
}

function sendMail() {
  var email = Session.getActiveUser().getEmail();
  
  MailApp.sendEmail(email , "初めてのGoogle Apps Script", "Hello World. Welcome Google Apps Script!");
}
