//鮮度の★赤色付け（日１、朝）、色統計（※実行ログは統合ログを見よ）
function fcheckLogDaily() {

  const sheet = bbsLib.getSheetByIdGid(id_bb, gid_fcheck);
  const sheetSum = bbsLib.getSheetByIdGid(id_bbLog, gid_fcheckDaySum);//日ごと統計

  var r_red_p = sheetSum.getRange(2, 8).getValue();//前回の赤割合

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
  var sumary = [today_ymddhm, n_all, n_white, r_white, n_pink, r_pink, n_red, r_red];
  bbsLib.addLogFirst(sheetSum, 2, [sumary], 8, 10000);

  r_red_p = r_red - r_red_p;

  //赤割合、前回比を返す
  return { r_red, r_red_p };

}
