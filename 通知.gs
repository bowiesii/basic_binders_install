//期間統計作成者へメール
function mail_summary(email, filename, fileUrl) {
  var subject = "【笠間店】期間統計を作成しました。";
  var body = "期間統計出力フォルダ";
  body = body + "\n" + "https://drive.google.com/drive/folders/19t8-VEtn-LQ4pIP2Tdet4Cd7mco815YE";
  body = body + "\n" + filename;
  body = body + "\n" + fileUrl
  MailApp.sendEmail(email, subject, body);//本人にメール通知
}

//統計※管理者へは生存確認的目的でメール送る。
//ユーザーは統計コマンドで確認してもらう。
function mail_summaryDay(body) {
  var subject = "笠間店日報";
  MailApp.sendEmail("youseimale@gmail.com", subject, body);//【管理者】にメール通知のみ。
}

//発注警告★bot
function mail_order() {
  const subject = '笠間店発注未報告通知'; //件名
  let body = `本日朝締めの発注の一部または全部が現在未報告です。
確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=648587868
`;

  botLib.pushSB(subject, body);//bot通知
}

//opt1=シート名重複
//opt2=共有登録なし
//opt3=新人作成したとき→★bot
//opt4=シート名なし
//opt5=新人削除したとき→★bot
//opt6=新人作成したとき
//opt7=新人削除したとき
function mail_sinjin(address, sinjinN, opt) {

  var subject = "";
  var body = "";

  if (opt == 1) {//同氏名があったとき
    subject = '【笠間店】同じ氏名の新人教育表シートが既にあるようです'; //件名
    body = `氏名「` + sinjinN + `」さん用の新人用シートが既にあるようなので、新しいシートを作成しませんでした。

１，既にほかの誰かがシートを作成済み、もしくは
２，過去の新人とたまたま同じ氏名
の可能性があります。

以下のシートを確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit

２，だった場合は、別の区別可能な氏名で再作成して下さい。（性だけでなく名も含めるなど）
新人教育表作成フォーム
https://docs.google.com/forms/d/e/1FAIpQLSc0yBXDQc6dxrZxiMApc5tT0KgOCCHvvKeQuMmowoUGxQXPKw/viewform

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 2) {//共有未登録のとき
    subject = 'ファイルの共有登録を行って下さい。'; //件名
    body = `新人教育シートを作成or削除する前に、お使いのGoogleアカウントでの笠間店ファイルの共有登録が必要です。

共有登録は以下から。
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 3) {//新人作成したとき→★bot
    subject = "笠間店新人表作成通知"; //件名
    body = "新人表（" + sinjinN + "さん用）が作成されました。";
    address = "bot";

  } else if (opt == 4) {//新人削除しようとしたがシート名が見つからない
    subject = '【笠間店】入力された名称のシートは無いようです。'; //件名
    body = `シート名「` + sinjinN + `」の新人用シートが無いため、シートの削除を行いませんでした。
以下のファイルを見て、シート名が正しいか確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit

新人教育表削除フォーム
https://docs.google.com/forms/d/e/1FAIpQLSe04FTp2UNkWXTZdRXgjNb-BPGxMm6l35SfvYyiBFifqmyIzw/viewform

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 5) {//新人削除したとき→★bot
    subject = "笠間店新人表手動削除通知"; //件名
    body = "新人表（" + sinjinN + "）が手動削除されました。";
    address = "bot";

  } else if (opt == 6) {//新人作成
    subject = '【笠間店】新人教育シートを作成しました。'; //件名
    body = sinjinN + `さん用の新人教育シートを「基本バインダー類」ファイル内に作成しました。
以下から確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 7) {//新人削除
    subject = '【笠間店】新人教育シートを手動削除しました。'; //件名
    body = sinjinN + `シートを「基本バインダー類」ファイル内から削除しました。
以下のフォルダにバックアップされています。
https://drive.google.com/drive/folders/12QZoEbx8TU6LpHUnZEaykx4Y__MWEOMG

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  //メールor通知
  if (address == "bot") {
    botLib.pushSB(subject, body);
  } else {
    MailApp.sendEmail(address, subject, body);
  }

}

