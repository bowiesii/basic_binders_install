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

//シート全体を保護、ただし灰色セル以外は編集可能に
function protectExceptGray(sheet) {

  var protection = sheet.protect();
  protection.removeEditors(protection.getEditors());//編集者をすべて削除
  var noProtRanges = [];//保護しない範囲の配列を作る

  var maxrow = sheet.getLastRow();
  var maxcol = sheet.getLastColumn();
  var bgcary = sheet.getRange(1, 1, maxrow, maxcol).getBackgrounds();
  for (let r = 1; r <= maxrow; r++) {
    for (let c = 1; c <= maxcol; c++) {
      if (bgcary[r - 1][c - 1] != "#b7b7b7") {//灰色以外は保護解除
        noProtRanges.push(sheet.getRange(r, c));//範囲を追加
      }
    }
  }

  Logger.log(noProtRanges.length);
  protection.setUnprotectedRanges(noProtRanges);//範囲配列を保護の例外にする

}

