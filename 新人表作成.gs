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

  if (sinjinN.length <= 1 || sinjinN.length >= 9) {
    Logger.log("氏名文字数エラー");
    return;
  }

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    sendmail(email, sinjinN, 2);
    return;
  }

  //既に新人氏名のシートがあった場合スルー→メール
  if (bbsLib.getGIDbysheetname(bbSpreadSheet, "【新】" + sinjinN) != null) {
    Logger.log("作成済み");
    sendmail(email, sinjinN, 1);
    return;
  }

  var sheetOri = bbsLib.getSheetBySperadGid(bbSpreadSheet, gid_sinjinOri);//原本
  var sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_sinjinList);//新人リスト

  //長期経過新人シート→削除と移動・ログ記録
  if (sheetLog.getLastRow() - 1 >= 1) {//０だとgetrangeがエラーになるので

    var sinjinL = sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, sheetLog.getLastColumn()).getDisplayValues();//新人の情報
    var sheetS = bbSpreadSheet.getSheets();//すべてのシートが配列に

    for (let na = 0; na <= sheetS.length - 1; na++) {
      var gid = sheetS[na].getSheetId();
      for (let nb = 0; nb <= sinjinL.length - 1; nb++) {
        if (gid == sinjinL[nb][2]) {//このGIDは新人シート
          var date_s = new Date(sheetS[na].getRange(3, 4).getDisplayValue());//特定セルに最終更新を記録
          var date_t = new Date(today_ymd);
          var dd = (date_t - date_s) / 86400000;//経過日数（ミリ秒を日に変換）
          if (dd >= 45) {//最終更新から４５日以上→削除と移動・ログ記録
            Logger.log(sheetS[na].getSheetName() + " " + gid + " 削除と移動とログ記録");




          } else {
            Logger.log(sheetS[na].getSheetName() + " " + gid + " 新人シートだが残す");
          }
        }
      }
    }

  }


  //シートコピー・ログ記録
  Logger.log("コピー段階");
  var newSheet = sheetOri.copyTo(bbSpreadSheet);//コピー
  newSheet.setName("【新人】" + sinjinN);
  newSheet.getRange(2, 2).setValue(sinjinN);
  newSheet.getRange(3, 2).setValue(today_ymd);
  newSheet.getRange(3, 4).setValue(today_ymd);
  var logary = [today_ymddhm, sinjinN, newSheet.getSheetId(), "", ""];
  bbsLib.addLogFirst(sheetLog, 2, [logary], 5, 10000);

  //シート→保護、灰色以外のセル→保護解除
  var protection = newSheet.protect();
  protection.removeEditors(protection.getEditors());
  var noProtRanges = [];//保護しない範囲の配列を作る
  var maxrow = newSheet.getLastRow();
  var maxcol = newSheet.getLastColumn();
  var bgcary = newSheet.getRange(1, 1, maxrow, maxcol).getBackgrounds();
  for (let r = 1; r <= maxrow; r++) {
    for (let c = 1; c <= maxcol; c++) {
      if (bgcary[r - 1][c - 1] != "#b7b7b7") {//灰色以外は保護解除
        noProtRanges.push(newSheet.getRange(r, c));
      }
    }
  }
  Logger.log(noProtRanges.length)
  protection.setUnprotectedRanges(noProtRanges);





}


//opt1=シート名重複
//opt2=共有登録なし
function mail_sinjin(address, sinjinN, opt) {

  var subject = "";
  var body = "";

  if (opt == 1) {

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

  } else if (opt == 2) {

    subject = 'ファイルの共有登録を行って下さい。'; //件名
    body = `新人教育シートを作成する前に、お使いのGoogleアカウントでの笠間店ファイルの共有登録が必要です。

共有登録は以下から。
https://docs.google.com/forms/d/e/1FAIpQLSe4Y8moCb8cNc2Kg-fISMFaxXwS2mF11WapzGBVmMCDLPGcYA/viewform

※このメールは自動配信です。
`;

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  MailApp.sendEmail(address, subject, body);

}
