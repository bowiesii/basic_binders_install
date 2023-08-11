//日１早朝（メール報告も）※ゆくゆくはbot
function dailyMorning() {

  //使用統計ログ
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_useSumDay);

  //発注警告メール
  var orderM = orderDailyMorning();

  //実行しログ数など把握
  var simeiS = simeiLogDaily();
  var botS = botLogDaily();
  var { r_red, r_red_p } = fcheckLogDaily();//色付け・色付け統計のみ、赤割合返す★この書き方の場合元関数と同じシンボル名でないといけないらしい。
  var { allS, orderS, wtaskS, sinjinS, freshS } = intLogDaily();//統合ログ

  //前回の記録の日時
  var prevTime = sheetLog.getRange(2, 1).getDisplayValue();

  //共有登録者数
  var shareS = shareSheet.getLastRow() - 2;
  var shareS_p = shareS - sheetLog.getRange(2, 2).getDisplayValue();//前回からの変化

  //ユーザープロパティ（氏名）数
  var userpropS = bbsLib.getSheetByIdGid(id_bb, gid_h_simeiNow).getLastRow() - 1;
  var userpropS_p = userpropS - sheetLog.getRange(2, 3).getDisplayValue();//前回からの変化

  //新人表の数
  var sinjinNum = sinjinNumDaily();
  var sinjinNum_p = sinjinNum - sheetLog.getRange(2, 4).getDisplayValue();//前回からの変化

  //botフォロワー数-ブロック数
  var botUserS = botUserNumDaily();
  var botUserS_p = botUserS - sheetLog.getRange(2, 5).getDisplayValue();//前回からの変化

  //使用統計ログ書き込み
  var logary = [today_ymddhm, shareS, userpropS, sinjinNum, botUserS, simeiS, botS, orderM, allS, orderS, wtaskS, sinjinS, freshS];
  bbsLib.addLogFirst(sheetLog, 2, [logary], 13, 10000);

  //統計報告通知
  var body = "〇" + today_ymddhm + " 時点集計の日報です。";
  if (orderM == 1) {
    body = body + "\n★" + today_md + " 朝締め発注一部未報告でした。";
    body = body + "https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=648587868";
  } else {
    body = body + "\n" + today_md + " 朝締め発注は終わっています。";
  }
  body = body + "\n共有登録者数：" + shareS + "（前回比" + shareS_p + "）";
  body = body + "\n氏名数（推定）：" + userpropS + "（前回比" + userpropS_p + "）";
  body = body + "\n新人表の数：" + sinjinNum + "（前回比" + sinjinNum_p + "）";
  body = body + "\n" + "botユーザー数：" + botUserS + "（前回比" + botUserS_p + "）";
  body = body + "\n鮮度表の赤割合：" + r_red + "%（前回比" + r_red_p + "%）";
  body = body + "\n〇前回日報（" + prevTime + "）からの増加ログ";
  body = body + "\n氏名ログ数：" + simeiS;
  body = body + "\n" + "bot受信回数：" + botS;
  body = body + "\n統合ログ数（管理者以外）：" + allS;
  body = body + "\n発注ログ数：" + orderS;
  body = body + "\n週タスクログ数：" + wtaskS;
  body = body + "\n新人教育ログ数：" + sinjinS;
  body = body + "\n鮮度チェックログ数：" + freshS;
  body = body + "\n" + "使用統計（日１朝更新）" + "https://docs.google.com/spreadsheets/d/17bZ83U_NeHXLT__NOV0zfHd2B8XZIBEgKfd_akNZDuY/edit#gid=733648789";

  mail_summary(body);

}

//日１昼
function dailyNoon() {
  orderDailyNoon();
}

//週１日曜早朝
function weeklySunday() {
  wtaskLogWeekly();
}
