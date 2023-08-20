//対象フォームへトリガー設定（１回走ればよし。）
//function setTrigger() {
//var file = FormApp.openById("1Mu-TllN67tOQ9c56m_uEhW95TKaLTgty8L5DcYHVCxA");//登録フォーム：済
//var file = FormApp.openById("1pvjWB4EZ7Em4udc-vz2ciwC3rY82sXQGRtJNztRSdlw");//停止フォーム：済
//var functionName = "pushMail_stop"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

//プッシュ通知登録
function pushMail_sub(e) {
  var email = e.response.getRespondentEmail();
  Logger.log(email);

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, "", 2);//本人
    return;
  }

  var sheetPushList = bbsLib.getSheetByIdGid(id_bbLog, gid_pushList);//プッシュ登録者リスト
  var listRow = bbsLib.searchInCol(sheetPushList, 2, email);

  if (listRow != -1) {//登録済み
    mail_push_substop(2, email);

  } else {//未登録
    var logAry = [today_ymddhm, email]
    bbsLib.addLogFirst(sheetPushList, 2, [logAry], 2, 10000);
    mail_push_substop(1, email);

  }

}

//プッシュ通知解除
function pushMail_stop(e) {
  var email = e.response.getRespondentEmail();
  Logger.log(email);

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, "", 2);//本人
    return;
  }

  var sheetPushList = bbsLib.getSheetByIdGid(id_bbLog, gid_pushList);//プッシュ登録者リスト
  var listRow = bbsLib.searchInCol(sheetPushList, 2, email);

  if (listRow != -1) {//登録済み
    sheetPushList.deleteRow(listRow);
    mail_push_substop(3, email);

  } else {//未登録
    mail_push_substop(4, email);

  }

}
