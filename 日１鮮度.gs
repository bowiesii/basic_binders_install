//鮮度ログ移動（日１、朝）、色付け
function fcheckLogDaily() {

  const sheet = bbsLib.getSheetByIdGid(id_bb, gid_fcheck);
  const sheetTempLog = bbsLib.getSheetByIdGid(id_bb, gid_h_fcheck);//一時ログ
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_fcheckDay);//ログ
  const sheetSum = bbsLib.getSheetByIdGid(id_bbLog, gid_fcheckDaySum);//日ごと統計

  var didnum = sheetTempLog.getLastRow() - 1;//ログの数
  bbsLib.replaceLogFirst(sheetTempLog, sheetLog);//ログ移動

  //本体ファイル色付けの更新
  var maxrow = sheet.getLastRow();
  var maxcol = sheet.getLastColumn();
  var sheet_val = sheet.getRange(1, 1, maxrow, maxcol).getDisplayValues();
  var sheet_note = sheet.getRange(1, 1, maxrow, maxcol).getNotes();
  var n_white = 0;
  var n_pink = 0;
  var n_red = 0;

  for (var r = 0; r <= maxrow - 1; r++) {
    for (var c = 0; c <= maxcol - 1; c++) {
      if (sheet_val[r][c] != "未") {//氏名入力用セルは気にしなくてよい。
        continue;
      }
      if (sheet_note[r][c] != "") {
        var notestr = sheet_note[r][c];
        var string_date = notestr.match(/\d{4}\/\d{1,2}\/\d{1,2}/);//マッチしなければnullを返す、複数あれば最初のにマッチする
        if (string_date != null) {
          var ddate = new Date(string_date);
          var ddate_t = new Date(today_ymd);
          var dd = (ddate_t - ddate) / 86400000;//経過日数（ミリ秒を日に変換）
          Logger.log(r + " " + c + " " + string_date + " " + dd);

          //最終作業日15日～45日後にかけて徐々に赤くなる
          if (dd <= 15) {
            sheet.getRange(r + 1, c + 1).setBackground(null);
            n_white = n_white + 1;
          } else if (dd >= 45) {
            sheet.getRange(r + 1, c + 1).setBackgroundRGB(255, 15, 15);
            n_red = n_red + 1;
          } else {
            sheet.getRange(r + 1, c + 1).setBackgroundRGB(255, 255 - (dd - 15) * 8, 255 - (dd - 15) * 8);
            n_pink = n_pink + 1;
          }

        } else {//メモはあるが日付にマッチしない場合真っ赤
          sheet.getRange(r + 1, c + 1).setBackgroundRGB(255, 15, 15);
          n_red = n_red + 1;
        }
      } else {//メモがない場合真っ赤
        sheet.getRange(r + 1, c + 1).setBackgroundRGB(255, 15, 15);
        n_red = n_red + 1;
      }

    }
  }

  //統計書き込み
  var n_all = n_white + n_pink + n_red;
  var r_white = Math.round((n_white / n_all) * 100);
  var r_pink = Math.round((n_pink / n_all) * 100);
  var r_red = Math.round((n_red / n_all) * 100);
  var sumary = [today_ymddhm, didnum, n_all, n_white, r_white, n_pink, r_pink, n_red, r_red];
  bbsLib.addLogFirst(sheetSum, 2, [sumary], 9, 10000);

  return didnum;

}
