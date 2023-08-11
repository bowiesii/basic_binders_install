//統計★botにも
function mail_summary(body) {
  var address = "youseimale@gmail.com";
  var subject = "笠間店統計情報";

  //通知
  MailApp.sendEmail(address, subject, body);
  botLib.pushSB(subject, body);

  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_useSumDay);
  sheetLog.getRange(1, 1).setNote(subject + "\n" + body);//bot返信用にメールの内容をとっておく
}

//発注警告★botにも
function mail_order() {
  var address = "youseimale@gmail.com";
  const subject = '本日朝締め発注が終わっていない可能性'; //件名
  let body = `本日朝締めの発注の一部または全部が未報告であり、終わっていない可能性があります。
確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=648587868
`;

  //通知
  MailApp.sendEmail(address, subject, body);
  botLib.pushSB(subject, body);
}

//opt1=シート名重複
//opt2=共有登録なし
//opt3=新人作成したとき→youseimale★botにも
//opt4=新人削除しようとしたがシート名が見つからない
//opt5=新人削除したとき→youseimale★botにも
function mail_sinjin(address, sinjinN, opt) {

  var subject = "";
  var body = "";

  if (opt == 1) {//同氏名があったとき
    subject = '同じ氏名の新人教育表シートが既にがあるようです'; //件名
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
    body = `新人教育シートを作成or削除する前に、お使いのGoogleアカウントでの笠間店ファイルの共有登録が必要です。

共有登録は以下から。
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

※このメールは自動配信です。
`;

  } else if (opt == 3) {//新人作成したとき→youseimale★botにも
    subject = "新人表が作成されました。"; //件名
    body = address + "さんが新人表（" + sinjinN + "さん用）を作成。";
    address = "youseimale@gmail.com";

  } else if (opt == 4) {//新人削除しようとしたがシート名が見つからない
    subject = '入力された名称のシートは無いようです。'; //件名
    body = `シート名「` + sinjinN + `」の新人用シートは無いようです。
以下のファイルを見て、シート名が正しいか確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit

新人教育表削除フォーム
https://docs.google.com/forms/d/e/1FAIpQLSe04FTp2UNkWXTZdRXgjNb-BPGxMm6l35SfvYyiBFifqmyIzw/viewform

※このメールは自動配信です。
`;

  } else if (opt == 5) {//新人削除したとき→youseimale★botにも
    subject = "新人表が手動削除されました。"; //件名
    body = address + "さんが新人表（" + sinjinN + "）を手動削除。";
    address = "youseimale@gmail.com";

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  //通知
  MailApp.sendEmail(address, subject, body);
  if (opt == 3 || opt == 5) {
    botLib.pushSB(subject, body);
  }

}

