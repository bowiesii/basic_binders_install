//プッシュリスト全員にメール通知
function pushToMailList(sub, body) {

  var sheetPushList = bbsLib.getSheetByIdGid(id_bbLog, gid_pushList);//プッシュ登録者リスト
  var listNum = sheetPushList.getLastRow() - 1;//人数

  if (listNum >= 1) {
    var emailAry = sheetPushList.getRange(2, 2, listNum, 1).getDisplayValues();//二次元
    for (let r = 0; r <= emailAry.length - 1; r++) {
      var email = emailAry[r][0];
      MailApp.sendEmail(email, sub, body);//プッシュメール通知
    }
  }

}

//1=プッシュ通知登録
//2=すでに登録済み
//3=登録解除
//4=もともと登録解除済み
function mail_push_substop(opt, email) {

  var sub = "";
  var body = "";

  if (opt == 1) {
    sub = "【笠間店】プッシュ通知登録を行いました。";
    body = "以後、このメールアドレスへプッシュ通知配信を行います。\n";
    body = body + `通知タイミング：
〇発注未報告のとき（毎朝４～５時チェック）
〇ファイル共有登録・解除・氏名変更時
〇新人教育表作成・手動削除時`;
    body = body + "\nプッシュ通知停止はこちら\n" + "https://docs.google.com/forms/d/e/1FAIpQLSeOOfCqlYouW8n1IakzqehkwvBTmpGmQ-fHish55kE_yR9mmg/viewform";

  } else if (opt == 2) {
    sub = "【笠間店】プッシュ通知登録済みです。"
    body = "既にプッシュ通知登録済みです。";

  } else if (opt == 3) {
    sub = "【笠間店】プッシュ通知を停止します。"
    body = "プッシュ通知配信を停止します。";

  } else if (opt == 4) {
    sub = "【笠間店】既にプッシュ通知配信は停止されています。"
    body = "もともとプッシュ通知登録されていません。";

  } else {
    Logger.log("opt指定エラー");
  }

  MailApp.sendEmail(email, sub, body);//本人にメール通知

}

//期間統計作成者へメール
function mail_makeSummary(opt, email, filename, fileUrl) {
  if (opt == 0) {
    let subject = "【笠間店】期間統計の作成に失敗しました。";
    let body = "期間の設定が誤っているか、もしくは期間内にデータが無かったと思われます";
    body = body + "\n\n問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。\n※このメールは自動配信です。";
    MailApp.sendEmail(email, subject, body);//本人にメール通知

  } else if (opt == 1) {
    let subject = "【笠間店】期間統計を作成しました。";
    let body = "期間統計出力フォルダ";
    body = body + "\n" + "https://drive.google.com/drive/folders/19t8-VEtn-LQ4pIP2Tdet4Cd7mco815YE";
    body = body + "\n\n" + "作成したファイル";
    body = body + "\n" + filename;
    body = body + "\n" + fileUrl
    body = body + "\n" + "あなたにこのファイルの編集権限を付与しました。";
    body = body + "\n\n問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。\n※このメールは自動配信です。";
    MailApp.sendEmail(email, subject, body);//本人にメール通知

  }

}

//統計※管理者へは生存確認的目的でメール送る。
//ユーザーは統計コマンドで確認してもらう。
function mail_summaryDay(body) {
  var subject = "（笠間店管理者向け）日報";
  MailApp.sendEmail("youseimale@gmail.com", subject, body);//【管理者】にメール通知のみ。
}

//週報→youseimale
function mail_summalyMonday(body) {
  var subject = "（笠間店管理者向け）週報";
  MailApp.sendEmail("youseimale@gmail.com", subject, body);//【管理者】にメール通知のみ。
}

//発注警告★プッシュ配信
function mail_order() {
  const subject = '（笠間店プッシュ通知）発注未報告'; //件名
  let body = `本日朝締めの発注の一部または全部が現在未報告です。
確認して下さい。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=648587868
`;
  body = body + "\nプッシュ通知停止はこちら\n" + "https://docs.google.com/forms/d/e/1FAIpQLSeOOfCqlYouW8n1IakzqehkwvBTmpGmQ-fHish55kE_yR9mmg/viewform";

  pushToMailList(subject, body);//プッシュ通知
}

//opt1=シート名重複
//opt2=共有登録なし
//opt3=新人作成したとき→★プッシュ配信
//opt4=シート名なし
//opt5=新人削除したとき→★プッシュ配信
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
    body = `お使いのGoogleアカウントでの笠間店ファイルの共有登録が必要です。

共有登録は以下から。
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 3) {//新人作成したとき→★プッシュ配信
    subject = "（笠間店プッシュ通知）新人表作成"; //件名
    body = "新人表（" + sinjinN + "さん用）が作成されました。";
    body = body + "\nプッシュ通知停止はこちら\n" + "https://docs.google.com/forms/d/e/1FAIpQLSeOOfCqlYouW8n1IakzqehkwvBTmpGmQ-fHish55kE_yR9mmg/viewform";
    address = "push";

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

  } else if (opt == 5) {//新人削除したとき→★プッシュ配信
    subject = "（笠間店プッシュ通知）新人表手動削除"; //件名
    body = "新人表（" + sinjinN + "）が手動削除されました。";
    body = body + "\nプッシュ通知停止はこちら\n" + "https://docs.google.com/forms/d/e/1FAIpQLSeOOfCqlYouW8n1IakzqehkwvBTmpGmQ-fHish55kE_yR9mmg/viewform";
    address = "push";

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

  //メールorプッシュ通知
  if (address == "push") {
    pushToMailList(subject, body);
  } else {
    MailApp.sendEmail(address, subject, body);
  }

}

