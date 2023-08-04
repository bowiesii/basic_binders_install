//日１早朝
function dailyMorning() {

  //発注警告メール
  var orderM = orderDailyMail();

  //以下は数を報告メール
  var simeiS = simeiLogDaily();
  var wtaskS = wtaskLogDaily();
  var fcheckS = fcheckLogDaily();
  var editS = editCountDaily();

  var address = "youseimale@gmail.com";//宛先
  var subject = "前日の統計情報"; //件名
  var body = "前日の統計情報です。";
  body = body + "\n氏名ログ数：" + simeiS;
  body = body + "\n週タスクログ数：" + wtaskS;
  body = body + "\n鮮度チェックログ数：" + fcheckS;
  body = body + "\n総編集ログ数（管理者以外）：" + editS;
  if (orderM == "m") {
    body = body + "\n発注一部未報告のためメール発信しました。";
  } else {
    body = body + "\n発注は終わっています。";
  }
  body = body + "\n※このメールは自動配信です。";

  MailApp.sendEmail(address, subject, body);

}

//日１昼
function dailyNoon() {
  orderDaily();
}

//週１日曜早朝
function weeklySunday() {
  wtaskLogWeekly();
}
