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

//統合ログ移動（日１、朝）
function intLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_log);
  var allS = sheetTempLog.getLastRow() - 1;

  //発注、週タスク、新人、鮮度のログ数
  var orderS = 0;
  var wtaskS = 0;
  var sinjinS = 0;
  var freshS = 0;

  if (allS != 0) {//getvaluesのエラー防止
    var snAry = sheetTempLog.getRange(2, 3, allS, 1).getValues();
    for (let row = 0; row <= snAry.length - 1; row++) {
      if (snAry[row][0].includes("【新】")) {//自由にシート名つけられえるのでこれが最初
        sinjinS++;
      } else if (snAry[row][0] == "発注") {
        orderS++;
      } else if (snAry[row][0] == "鮮度") {
        freshS++;
      } else if (snAry[row][0].includes("週バ")) {
        wtaskS++;
      }
    }
  }

  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_intLog);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return { allS, orderS, wtaskS, sinjinS, freshS };
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
