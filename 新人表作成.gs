//対象フォームへトリガー設定（１回走ればよし。）
//function setTrigger() {
//var file = FormApp.openById("1trWYICYzaVGK_5FXjzMupKGYhSbFHMW92AZllNbG5xY");//作成フォーム：済
//var functionName = "newSinjin"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

//新人作成フォームからトリガー
function newSinjin(e) {

  var answer = e.response.getItemResponses();
  var sinjinN = answer[0].getResponse();
  var email = e.response.getRespondentEmail();
  Logger.log("new " + sinjinN + " " + email);
  var newSheetName = "【新】" + sinjinN;

  if (sinjinN.length <= 1 || sinjinN.length >= 9) {
    Logger.log("氏名文字数エラー");
    return;
  }

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, sinjinN, 2);//本人
    return;
  }

  //既に新人氏名のシートがあった場合スルー→メール
  if (bbsLib.getGIDbysheetname(bbSpreadSheet, newSheetName) != null) {
    Logger.log("作成済み");
    mail_sinjin(email, sinjinN, 1);//本人
    return;
  }

  var sheetOri = bbsLib.getSheetByIdGid(id_sinjinOri, gid_sinjinOri);//原本
  var sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_sinjinList);//新人リスト

  //長期経過新人シート→自動削除と移動・ログ記録
  var sheetS = bbSpreadSheet.getSheets();//すべてのシートが配列に
  for (let na = 0; na <= sheetS.length - 1; na++) {
    if (sheetS[na].getSheetName().includes("【新】")) {

      var date_s = new Date(sheetS[na].getRange(3, 4).getDisplayValue());//特定セルに最終更新が記録すみ
      var date_t = new Date(today_ymd);
      var dd = (date_t - date_s) / 86400000;//経過日数（ミリ秒を日に変換）
      if (dd >= 30) {//最終更新から30日以上経過→削除と移動・ログ記録
        var gid = sheetS[na].getSheetId();

        Logger.log(sheetS[na].getSheetName() + " 移動と削除とログ記録");
        var newfilename = sheetS[na].getSheetName() + "（" + sheetS[na].getRange(3, 4).getDisplayValue() + "最終更新、自動削除）";
        copyToNewSpreadsheet(sheetS[na], "12QZoEbx8TU6LpHUnZEaykx4Y__MWEOMG", newfilename);//移動
        bbSpreadSheet.deleteSheet(sheetS[na]);//シート削除

        //削除をログ記録
        var gidary = sheetLog.getRange(2, 4, sheetLog.getLastRow() - 1, 1).getDisplayValues();
        for (nb = 0; nb <= gidary.length - 1; nb++) {
          if (gidary[nb][0] == gid) {
            sheetLog.getRange(nb + 2, 4).setValue("削除・移動済み");//GID情報削除
            sheetLog.getRange(nb + 2, 5).setValue(today_ymddhm);//削除日時
            sheetLog.getRange(nb + 2, 6).setValue("自動");//表削除者
          }
        }
      } else {
        Logger.log(sheetS[na].getSheetName() + " 新人シートだが残す");
      }

    }
  }

  //シートコピー・追加をログ記録・保護
  Logger.log("コピー段階");
  var newSheet = sheetOri.copyTo(bbSpreadSheet);//コピー
  newSheet.setName(newSheetName);//シート名など設定
  newSheet.getRange(2, 2).setValue(sinjinN);//新人氏名
  newSheet.getRange(3, 2).setValue(today_ymd);//作成日
  newSheet.getRange(3, 4).setValue(today_ymd);//最終更新
  newSheet.getRange(4, 4).setValue(newSheet.getSheetId());//GID

  //保護設定を原本からコピー
  copySheetProtection(sheetOri, newSheet);

  var logary = [today_ymddhm, sinjinN, email, newSheet.getSheetId(), ""];
  bbsLib.addLogFirst(sheetLog, 2, [logary], 5, 10000);

  mail_sinjin(email, sinjinN, 6);//本人
  mail_sinjin(email, sinjinN, 3);//bot

}

