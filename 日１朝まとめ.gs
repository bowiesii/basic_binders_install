//氏名ログ移動（日１、朝）
function simeiLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_simei);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_simei);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return sum;

}

//新人ログ移動（日１、朝）
function sinjinLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_sinjin);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_sinjinDay);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return sum;

}

//週タスクログ移動（日１、朝）
function wtaskLogDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_wtask);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_wtaskDay);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  return sum;

}

//編集数ログ移動と統計追加（日１、朝）
function editCountDaily() {

  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_edit);
  var sum = sheetTempLog.getLastRow() - 1;
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_editDay);
  const sheetLogSum = bbsLib.getSheetByIdGid(id_bbLog, gid_editDaySum);
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  var logary = [today_ymddhm, sum];
  bbsLib.addLogFirst(sheetLogSum, 2, [logary], 2, 10000);//統計追加

  return sum;

}