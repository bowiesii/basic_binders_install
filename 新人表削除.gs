//対象フォームへトリガー設定（１回走ればよし。）
//function setTrigger2() {
//var file = FormApp.openById("1pnmKX6FyroPPrG3-0daILMH_3RYkQNUruhGpC4m0NQ8");//削除フォーム：済
//var functionName = "delSinjin"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

function delSinjin(e) {

  var answer = e.response.getItemResponses();
  var delSheetName = answer[0].getResponse();//削除者が入力する氏名には【新】も含めないといけないようにする（ミス防止）
  var email = e.response.getRespondentEmail();

  //「鮮度」とか打ち込まれると困るので【新】のないシート名ははじく。
  if (delSheetName.includes("【新】") == false) {
    Logger.log("invalid sheetname");
    MailApp.sendEmail(email, "笠間店より", "シート名指定エラー");
    return;
  }

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, delSheetName, 2);//本人
    return;
  }

  Logger.log("del " + delSheetName + " " + email);
  var sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_sinjinList);//新人リスト

  var sheetS = bbSpreadSheet.getSheets();//すべてのシートが配列に
  for (let na = 0; na <= sheetS.length - 1; na++) {
    if (sheetS[na].getSheetName() == delSheetName) {
      var gid = sheetS[na].getSheetId();

      Logger.log(sheetS[na].getSheetName() + " 移動と削除とログ記録");
      var newfilename = sheetS[na].getSheetName() + "（" + sheetS[na].getRange(3, 4).getNote() + "最終更新、手動削除）";
      copyToNewSpreadsheet(sheetS[na], "12QZoEbx8TU6LpHUnZEaykx4Y__MWEOMG", newfilename);//移動
      bbSpreadSheet.deleteSheet(sheetS[na]);//シート削除

      //削除をログ記録
      var gidary = sheetLog.getRange(2, 4, sheetLog.getLastRow() - 1, 1).getDisplayValues();
      for (nb = 0; nb <= gidary.length - 1; nb++) {
        if (gidary[nb][0] == gid) {
          sheetLog.getRange(nb + 2, 4).setValue("削除・移動済み");//GID情報削除
          sheetLog.getRange(nb + 2, 5).setValue(today_ymddhm);//削除日時
          sheetLog.getRange(nb + 2, 6).setValue(email);//表削除者
        }
      }

      mail_sinjin(email, delSheetName, 7);//本人
      mail_sinjin(email, delSheetName, 5);//bot
      return;
    }
  }

  //シートが見つからないとき返信メール
  mail_sinjin(email, delSheetName, 4);//本人
  Logger.log("cannot find sheet");

}