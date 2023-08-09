//日１早朝（メール報告も）※ゆくゆくはbot
function dailyMorning() {

  //使用統計ログ
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_useSumDay);

  //発注警告メール
  var orderM = orderDailyMail();

  //実行しログ数把握
  var simeiS = simeiLogDaily();
  var wtaskS = wtaskLogDaily();
  var fcheckS = fcheckLogDaily();
  var sinjinS = sinjinLogDaily();
  var editS = editCountDaily();
  var botS = botLogDaily();

  //前回の記録の日時
  var prevTime = sheetLog.getRange(2, 1).getDisplayValue();

  //共有登録者数
  var shareS = shareSheet.getLastRow() - 2;
  var shareS_p = shareS - sheetLog.getRange(2, 2).getDisplayValue();

  //ユーザープロパティ（氏名）数
  var userpropS = bbsLib.getSheetByIdGid(id_bb, gid_h_simeiNow).getLastRow() - 1;
  var userpropS_p = userpropS - sheetLog.getRange(2, 3).getDisplayValue();

  //新人表の数
  var sinjinNum = sinjinNumDaily();
  var sinjinNum_p = sinjinNum - sheetLog.getRange(2, 4).getDisplayValue();

  //統計ログ書き込み
  var logary = [today_ymddhm, shareS, userpropS, sinjinNum, editS, simeiS, botS, orderM, wtaskS, sinjinS, fcheckS];
  bbsLib.addLogFirst(sheetLog, 2, [logary], 11, 10000);

  //統計報告メール
  var subject = "笠間店統計情報"; //件名
  var body = "〇" + today_ymddhm + " 時点集計の日報です。";
  if (orderM == 1) {
    body = body + "\n★" + today_md + " 朝締め発注一部未報告でした。";
    body = body + "https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=648587868?openExternalBrowser=1";
  } else {
    body = body + "\n" + today_md + " 朝締め発注は終わっています。";
  }
  body = body + "\n共有登録者数：" + shareS + "（前日比" + shareS_p + "）";
  body = body + "\n氏名数（推定）：" + userpropS + "（前日比" + userpropS_p + "）";
  body = body + "\n新人表の数：" + sinjinNum + "（前日比" + sinjinNum_p + "）";
  body = body + "\n〇前回日報（" + prevTime + "）からの増加ログ";
  body = body + "\n氏名ログ数：" + simeiS;
  body = body + "\n週タスクログ数：" + wtaskS;
  body = body + "\n新人教育ログ数：" + sinjinS;
  body = body + "\n鮮度チェックログ数：" + fcheckS;
  body = body + "\n総編集ログ数（管理者以外）：" + editS;
  body = body + "\n" + "bot受信回数：" + botS;
  body = body + "\n" + "使用統計（日１更新）" + "https://docs.google.com/spreadsheets/d/17bZ83U_NeHXLT__NOV0zfHd2B8XZIBEgKfd_akNZDuY/edit#gid=733648789?openExternalBrowser=1";

  sheetLog.getRange(1, 1).setNote(subject + "\n" + body);//bot用にメールの内容をとっておく

  MailApp.sendEmail("youseimale@gmail.com", subject, body);

}

//日１昼
function dailyNoon() {
  orderDaily();
}

//週１日曜早朝
function weeklySunday() {
  wtaskLogWeekly();
}
