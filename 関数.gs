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
function copyToNewSpreadsheet(fromSheet, toFolderID, newSpreadsheetName) {

  let toFolder = DriveApp.getFolderById(toFolderID);
  let newFile = DriveApp.getFileById(SpreadsheetApp.create(newSpreadsheetName).getId());
  let newSpreadsheet = SpreadsheetApp.openById(newFile.moveTo(toFolder).getId());
  let newSheet = fromSheet.copyTo(newSpreadsheet);
  newSheet.setName(fromSheet.getSheetName());
  newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);
  return newSpreadsheet;

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


