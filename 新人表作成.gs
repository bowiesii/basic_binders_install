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
  Logger.log(sinjinN + " " + email);
  var newSheetName = "【新】" + sinjinN;

  if (sinjinN.length <= 1 || sinjinN.length >= 9) {
    Logger.log("氏名文字数エラー");
    return;
  }

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, sinjinN, 2);
    return;
  }

  //既に新人氏名のシートがあった場合スルー→メール
  if (bbsLib.getGIDbysheetname(bbSpreadSheet, newSheetName) != null) {
    Logger.log("作成済み");
    mail_sinjin(email, sinjinN, 1);
    return;
  }

  var sheetOri = bbsLib.getSheetByIdGid(id_sinjinOri, gid_sinjinOri);//原本
  var sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_sinjinList);//新人リスト

  //長期経過新人シート→削除と移動・ログ記録
  var sheetS = bbSpreadSheet.getSheets();//すべてのシートが配列に
  for (let na = 0; na <= sheetS.length - 1; na++) {
    if (sheetS[na].getSheetName().includes("【新】")) {

      var date_s = new Date(sheetS[na].getRange(3, 4).getDisplayValue());//特定セルに最終更新が記録すみ
      var date_t = new Date(today_ymd);
      var dd = (date_t - date_s) / 86400000;//経過日数（ミリ秒を日に変換）
      if (dd >= 30) {//最終更新から30日以上経過→削除と移動・ログ記録
        var gid = sheetS[na].getSheetId();

        Logger.log(sheetS[na].getSheetName() + " 移動と削除とログ記録");
        var newfilename = sheetS[na].getSheetName() + "（" + sheetS[na].getRange(3, 4).getDisplayValue() + "最終更新）";
        copyToNewSpreadsheet(sheetS[na], "12QZoEbx8TU6LpHUnZEaykx4Y__MWEOMG", newfilename);//移動
        bbSpreadSheet.deleteSheet(sheetS[na]);//シート削除

        //ログ記録
        var gidary = sheetLog.getRange(2, 4, sheetLog.getLastRow() - 1, 1).getDisplayValues();
        for (nb = 0; nb <= gidary.length - 1; nb++) {
          if (gidary[nb][0] == gid) {
            sheetLog.getRange(nb + 2, 4).setValue("削除・移動済み");//GID情報削除
            sheetLog.getRange(nb + 2, 5).setValue(today_ymddhm);//削除日時
          }
        }
      } else {
        Logger.log(sheetS[na].getSheetName() + " 新人シートだが残す");
      }

    }
  }

  //シートコピー・ログ記録・保護
  Logger.log("コピー段階");
  var newSheet = sheetOri.copyTo(bbSpreadSheet);//コピー
  newSheet.setName(newSheetName);//シート名など設定
  newSheet.getRange(2, 2).setValue(sinjinN);
  newSheet.getRange(3, 2).setValue(today_ymd);
  newSheet.getRange(3, 4).setValue(today_ymd);

  var logary = [today_ymddhm, sinjinN, email, newSheet.getSheetId(), ""];
  bbsLib.addLogFirst(sheetLog, 2, [logary], 5, 10000);

  protectExceptGray(newSheet);//シートを保護、灰色セル以外は編集可
  mail_sinjin(email, sinjinN, 3);//youseimaleに報告メール（ゆくゆくはbot？）

}


//opt1=シート名重複
//opt2=共有登録なし
function mail_sinjin(address, sinjinN, opt) {

  var subject = "";
  var body = "";

  if (opt == 1) {//同氏名があったとき

    subject = '同じ氏名の新人教育表ファイルがあるようです'; //件名
    body = `氏名「` + sinjinN + `」の新人用シートが既にあるようなので、新しいシートを作成しませんでした。

１，既にほかの誰かがシートを作成済み、もしくは
２，過去の新人とたまたま同じ氏名
の可能性があります。

とりあえず以下のシートを確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit

２，の場合は別の区別可能な氏名で再作成して下さい。（性だけでなく名も含めるなど）
新人教育表作成フォーム
https://docs.google.com/forms/d/e/1FAIpQLSc0yBXDQc6dxrZxiMApc5tT0KgOCCHvvKeQuMmowoUGxQXPKw/viewform

※このメールは自動配信です。
`;

  } else if (opt == 2) {//共有未登録のとき

    subject = 'ファイルの共有登録を行って下さい。'; //件名
    body = `新人教育シートを作成する前に、お使いのGoogleアカウントでの笠間店ファイルの共有登録が必要です。

共有登録は以下から。
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

※このメールは自動配信です。
`;

  } else if (opt == 3) {//新人作成したとき→★youseimaleにメール

    subject = "新人表が作成されました。"; //件名
    body = address + "さんが新人表（" + sinjinN + "さん用）を作成。";
    address = "youseimale@gmail.com";

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  MailApp.sendEmail(address, subject, body);

}
