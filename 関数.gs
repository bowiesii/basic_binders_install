//Date型→統合ログ管理者以外（日１朝更新）の２行目から何行目までが範囲か。
function dateToRow(date) {

  const sheetRaw = bbsLib.getSheetByIdGid(id_bbLog, gid_intLog);//統合ログ
  var dNum = 1;//統合ログ何行目までがこの週のデータか。
  if (sheetRaw.getLastRow() >= 2) {
    var ymdAry = sheetRaw.getRange(2, 1, sheetRaw.getLastRow() - 1, 1).getValues();
    for (let row = 0; row <= ymdAry.length - 1; row++) {
      var dateR = new Date(ymdAry[row][0]);
      if ((dateR - date) < 0) {//これのひとつ前まで。
        dNum = row + 1;
        break;
      } else {
        dNum = row + 2;
      }
    }
  }
  return dNum;

}

//配列（★１次元）の合計を返す（数字じゃなければプラス０）
function rowSum(ary) {
  if (ary == []) { return 0; }
  var sum = 0;

  for (let r = 0; r <= ary.length - 1; r++) {
    if (ary[r].isFinite()) {//数値だったらプラス
      sum = sum + ary[r];
    }
  }

  return sum;
}

//emailが共有リストシートに含まれるかどうか（含まれる→true）
function shareingOrNot(email) {

  //共有リストスシート
  var sheet = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg').getSheetByName("登録している人");

  var ary = sheet.getRange(3, 3, sheet.getLastRow() - 2, 1).getValues();
  for (let row = 0; row <= ary.length - 1; row++) {
    if (ary[row][0] == email) {
      return true;
    }
  }
  return false;

}

//シートをフォルダ内の新規ファイルにコピー
function copyToNewSpreadsheet(fromSheet, toFolderID, newSpreadSheetName) {

  let toFolder = DriveApp.getFolderById(toFolderID);
  let newFile = DriveApp.getFileById(SpreadsheetApp.create(newSpreadSheetName).getId());
  let newSpreadsheet = SpreadsheetApp.openById(newFile.moveTo(toFolder).getId());
  let newSheet = fromSheet.copyTo(newSpreadsheet);
  newSheet.setName(fromSheet.getSheetName());
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);
  return newSpreadsheet;

}

//配列を、新規スプシファイルを作ってコピー（gid=0）
//aryの１行目は必ずフルで。
function copyAryToNewSpreadSheet(ary, toFolderID, newSpreadSheetName, newSheetName) {

  let toFolder = DriveApp.getFolderById(toFolderID);
  let newFile = DriveApp.getFileById(SpreadsheetApp.create(newSpreadSheetName).getId());
  let newSpreadsheet = SpreadsheetApp.openById(newFile.moveTo(toFolder).getId());
  let sheet = bbsLib.getSheetBySperadGid(newSpreadsheet, "0");//最初のGIDは必ず０になる
  sheet.setName(newSheetName);
  sheet.getRange(1, 1, ary.length, ary[0].length).setValues(ary);
  return newSpreadsheet;//スプシオブジェクトを返す

}

//配列を、既存のスプレッドシート内に新規シートを作ってコピー
//aryの１行目は必ずフルで。
function copyAryToSpreadSheet(spreadSheet, ary, newSheetName) {

  var newSheet = spreadSheet.insertSheet();
  newSheet.setName(newSheetName);
  newSheet.getRange(1, 1, ary.length, ary[0].length).setValues(ary);
  return newSheet;//シートを返す

}

//シート１からシート２へ、シート保護したうえで非保護レンジをコピーする
//灰色セルでひとつひとつ設定すると、シート２が重くなるので。。
function copySheetProtection(sheet1, sheet2) {

  var protection1 = sheet1.protect();
  var ranges1 = protection1.getUnprotectedRanges();//非保護レンジの配列
  var protection2 = sheet2.protect();
  protection2.removeEditors(protection2.getEditors());//編集者をすべて削除
  if (ranges1 == []) { return; }
  var ranges2 = [];

  for (r = 0; r <= ranges1.length - 1; r++) {
    let row = ranges1[r].getRow();
    let col = ranges1[r].getColumn();
    let rowN = ranges1[r].getNumRows();
    let colN = ranges1[r].getNumColumns();
    //Logger.log(r + ">" + row + " " + col + " " + rowN + " " + colN);
    ranges2.push(sheet2.getRange(row, col, rowN, colN));
  }

  protection2.setUnprotectedRanges(ranges2);

}


