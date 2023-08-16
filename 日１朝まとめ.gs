//★これは月曜のみ、朝５～６時、日報後、週１週報作成
function weeklyMonday() {

  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_makeLogWeek);
  var date1 = new Date(sheetLog.getRange(2, 1).getValue());//前回作成日時

  //作成
  var { editN, ptN, newSpS } = makeSummalyInPeriod(date1, today, "1DCXss-JSW9z12z7Xx7dFtJ8x_6yrsccL", "週報");

  var url = "";
  var filename = "";
  var logAry = [today_ymddhm, 0, 0, ""];
  var body = today_ymddhm + "週報作成されませんでした。";

  if (newSpS != null) {//作成された
    url = bbsLib.toUrl(newSpS.getId(), "0");
    filename = newSpS.getName();
    logAry = [today_ymddhm, editN, ptN, url];

    body = today_ymddhm + "作成週報";
    body = body + "\n編集数" + editN;
    body = body + "\nポイント数" + ptN;
    body = body + "\n" + filename;
    body = body + "\n" + url;
  }

  body = body + "\n週報フォルダ\n" + "https://drive.google.com/drive/folders/1DCXss-JSW9z12z7Xx7dFtJ8x_6yrsccL";

  bbsLib.addLogFirst(sheetLog, 2, [logAry], 4, 10000);
  sheetLog.getRange(1, 1).setNote(body);//bot用にメモしておく
  mail_summalyMonday(body);

}

//氏名ログ移動（日１、朝）
function simeiLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_simei);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_simei);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return sum;
}

//botログ移動（日１、朝）※これだけ一時はログスプシ。
function botLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bbLog, gid_botTemp);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_botDay);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return sum;
}

//botフォロワー-ブロックの数を取得
function botUserNumDaily() {
  var { followS, blockS } = botLib.getFBNum();
  var sum = followS - blockS;
  return sum;
}

//統合ログ移動（日１、朝）
function intLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_log);
  var allS = sheetTempLog.getLastRow() - 1;

  //発注、週タスク、新人、鮮度のログ数
  var orderS = 0;
  var wtaskS = 0;
  var sinjinS = 0;
  var freshS = 0;
  var cleanS = 0;

  if (allS != 0) {//getvaluesのエラー防止
    var snAry = sheetTempLog.getRange(2, 7, allS, 1).getValues();//シート名の列
    for (let row = 0; row <= snAry.length - 1; row++) {
      if (snAry[row][0].includes("【新】")) {//自由にシート名つけられえるのでこれが最初
        sinjinS++;
      } else if (snAry[row][0] == "発注") {
        orderS++;
      } else if (snAry[row][0] == "鮮度") {
        freshS++;
      } else if (snAry[row][0] == "清掃") {
        cleanS++;
      } else if (snAry[row][0].includes("週バ")) {
        wtaskS++;
      }
    }
  }

  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_intLog);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return { allS, orderS, wtaskS, sinjinS, freshS, cleanS };
}

//新人人数カウント（日１、朝）
function sinjinNumDaily() {

  var sum = 0;
  var sheetS = bbSpreadSheet.getSheets();//すべてのシートが配列に
  for (let na = 0; na <= sheetS.length - 1; na++) {
    if (sheetS[na].getSheetName().includes("【新】")) {
      sum = sum + 1;
    }
  }

  return sum;
}
