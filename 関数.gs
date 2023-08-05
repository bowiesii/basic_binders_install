//emailが共有リストシートに含まれるかどうか（含まれる→true）
function shareingOrNot(email) {

  //共有リストスシート
  var sheet = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg').getSheetByName('シート1');

  var ary = sheet.getRange(3, 3, sheet.getLastRow() - 2, 1).getValues();
  for (let row = 0; row <= ary.length - 1; row++) {
    if (ary[row][0] == email) {
      return true;
    }
  }
  return false;

}

