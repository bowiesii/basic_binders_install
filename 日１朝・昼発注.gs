//発注ログ（日１、昼）
function orderDailyNoon() {

  //発注確認表
  const sheet1 = bbsLib.getSheetByIdGid(id_bb, gid_order);
  const sheet2 = bbsLib.getSheetByIdGid(id_bb, gid_orderOld);
  const sheetlog = bbsLib.getSheetByIdGid(id_bbLog, gid_orderDay);

  var tasknum = sheet1.getLastColumn() - 1;//今は６

  var array1 = sheet1.getRange(3, 1, 14, tasknum + 1).getDisplayValues();
  var array2 = sheet2.getRange(3, 1, 14, tasknum + 1).getDisplayValues();
  var logary;//別の場所へ記録用（本日〆）

  //todayとデータの4行目が一致するかチェック
  var array130 = Utilities.parseDate(array1[3][0], 'JST', 'yyyy/MM/dd');//date型に変換
  array130 = Utilities.formatDate(array130, 'JST', 'yyyy/MM/dd');
  if (today_ymd == array130) {

    //正常な場合
    logary = array1[3];
    array2.push(array1.shift());//a1の先頭行を削除し、a2の末尾に移動
    array2.shift();//a2の先頭行を削除

    //a1の最終行に新たな日付を追加
    var today_11 = new Date(todayYear + '/' + todayMonth + '/' + todayDate);
    today_11.setDate(today.getDate() + 11);
    var today_11_ymd = Utilities.formatDate(today_11, 'JST', 'yyyy/MM/dd');
    var today_11_w_jpn = wary[today_11.getDay()];
    array1.push([today_11_ymd + " " + today_11_w_jpn]);
    array1[array1.length - 1][tasknum] = "";//ちゃんと書き込めるように列数を確保

  } else {
    //異常な場合
    Logger.log("日付が違う");
    throw new Error("発生させた例外：日１発注：日付が違う");
  }

  Logger.log(array1);
  Logger.log(array2);
  Logger.log(logary);

  //ログ２行目に挿入
  Logger.log("ログ段階");
  sheetlog.insertRowBefore(2);
  sheetlog.getRange(2, 1, 1, logary.length).setValues([logary]);
  if (sheetlog.getLastRow() >= 10001) {
    sheetlog.deleteRow(10001);//10000行に制限
  }

  //発注確認表に書き込む
  Logger.log("発注確認表書き込み段階");
  sheet1.getRange(3, 1, 14, tasknum + 1).clearContent();
  sheet1.getRange(3, 1, 14, tasknum + 1).setValues(array1);
  sheet2.getRange(3, 1, 14, tasknum + 1).clearContent();
  sheet2.getRange(3, 1, 14, tasknum + 1).setValues(array2);

}

//発注メール（日１、朝）
function orderDailyMorning() {

  //発注確認表
  const sheet = bbsLib.getSheetByIdGid(id_bb, gid_order);
  var tasknum = sheet.getLastColumn() - 1;//今は６
  var array = sheet.getRange(6, 1, 1, tasknum + 1).getDisplayValues();

  //空白・改行削除
  for (let col = 0; col <= tasknum; col++) {
    array[0][col] = array[0][col].replace(/\s/g, "");
  }
  Logger.log("本日朝〆行：" + array);

  var mail = 0;
  if (array[0][1] == "") { mail = 1; }
  if (array[0][2] == "" && array[0][3] == "") { mail = 1; }
  if (array[0][4] == "" && array[0][5] == "") { mail = 1; }

  if (mail == 1) {
    Logger.log("メールする");
    mail_order();
  } else {
    Logger.log("メールしない");
  }

  return mail;//メールしたなら１、しないなら０

}

