//週１回、日曜午前４時～５時に起動
function wtaskLogWeekly() {

  //本日は日曜日？
  if (today_wjpn != "日") {
    Logger.log("日曜日でない");
    return;
  }

  const sheet1 = bbsLib.getSheetByIdGid(id_bb, gid_wtask1);
  const sheet2 = bbsLib.getSheetByIdGid(id_bb, gid_wtask2);
  const sheet3 = bbsLib.getSheetByIdGid(id_bb, gid_wtask3);

  //日付をB3から取り出し（ファイル名は変更できるため保護セルから）
  var sheet1_n = sheet1.getRange(3, 2).getDisplayValue();
  var sheet2_n = sheet2.getRange(3, 2).getDisplayValue();
  var sheet3_n = sheet3.getRange(3, 2).getDisplayValue();
  var sheet1_d = new Date(sheet1_n);
  var sheet2_d = new Date(sheet2_n);
  var sheet3_d = new Date(sheet3_n);
  Logger.log(sheet1_n + " " + sheet1_d);
  Logger.log(sheet2_n + " " + sheet2_d);
  Logger.log(sheet3_n + " " + sheet3_d);

  //一番古いシート→一番新しいシートに、並べ替える//01(2)34→0134 ３→２　４→３
  var ary = [[1, sheet1, sheet1_n, sheet1_d], [2, sheet2, sheet2_n, sheet2_d], [3, sheet3, sheet3_n, sheet3_d]];
  ary.sort((a, b) => { return a[3] - b[3]; });//_dでソートする
  Logger.log(ary);

  //あしたの日付
  var today_1 = new Date(today);//この行はなんか必要らしい
  today_1.setDate(today.getDate() + 1);//明日にする
  today_1 = Utilities.formatDate(today_1, "JST", "yyyy/MM/dd")//時間部分をカット（00:00にする）
  today_1 = new Date(today_1);//デフォ形式に戻す
  //真ん中日付
  var ary12k = new Date(ary[1][2]);
  Logger.log(ary12k + " " + today_1);//真ん中日付＝明日だったら、変更済みということ

  if (Number(ary12k) == Number(today_1)) {//Date型はミリ秒（１９７０・１・１～）変換しないと等値比較不可。
    Logger.log("実行済のようです");
    throw new Error("発生させた例外：週１週タスク：実行済のようです");
  }

  //ary[0]いちばん古いファイル（１３日前週タスク）：情報を別ファイルに記録する→中身クリア→一番新しいファイル（８日後週タスク）にする
  //ary[1]二番目に古いファイル（６日前週タスク）：
  //ary[2]三番目に古いファイル（１日後タスク）：真ん中

  ary[0][3] = new Date(today);
  ary[0][3].setDate(today.getDate() + 8);//作業週
  ary[0][2] = Utilities.formatDate(ary[0][3], 'JST', 'yyyy/MM/dd');
  ary[0][4] = new Date(today);
  ary[0][4].setDate(today.getDate() + 15);//納品週
  ary[0][4] = Utilities.formatDate(ary[0][4], 'JST', 'yyyy/MM/dd');
  ary[0][5] = Utilities.formatDate(ary[0][3], 'JST', 'MM/dd') + "週バ" + ary[0][0];//シート名（新）
  Logger.log(ary[0][5] + " 納品は" + ary[0][4]);

  ary[1][3] = new Date(today);
  ary[1][3].setDate(today.getDate() - 6);//作業週
  ary[1][2] = Utilities.formatDate(ary[1][3], 'JST', 'yyyy/MM/dd');
  ary[1][4] = new Date(today);
  ary[1][4].setDate(today.getDate() + 1);//納品週
  ary[1][4] = Utilities.formatDate(ary[1][4], 'JST', 'yyyy/MM/dd');
  ary[1][5] = Utilities.formatDate(ary[1][3], 'JST', 'MM/dd') + "週バ" + ary[1][0];//シート名（古）
  Logger.log(ary[1][5] + " 納品は" + ary[1][4]);

  ary[2][3] = new Date(today);
  ary[2][3].setDate(today.getDate() + 1);//作業週
  ary[2][2] = Utilities.formatDate(ary[2][3], 'JST', 'yyyy/MM/dd');
  ary[2][4] = new Date(today);
  ary[2][4].setDate(today.getDate() + 8);//納品週
  ary[2][4] = Utilities.formatDate(ary[2][4], 'JST', 'yyyy/MM/dd');
  ary[2][5] = Utilities.formatDate(ary[2][3], 'JST', 'MM/dd') + "週バ" + ary[2][0];//シート名（真ん中）
  Logger.log(ary[2][5] + " 納品は" + ary[2][4]);


  Logger.log("ログ記録段階");

  /*
  //基本バインダー_ログ内シートにバックアップする場合。
  const sheetLog = bbsLib.getSheetByIdGid(id_bbLog, gid_wtaskWeek);
  //ログシートに[0]をログ記録する
  var logary = [today_ymddhm];//[0]は本日日付
  logary[1] = ary[0][1].getRange(2, 2).getValue();//バインダー
  logary[2] = ary[0][1].getRange(3, 2).getValue();//作業
  logary[3] = ary[0][1].getRange(4, 2).getValue();//納品
  var taskn = ary[0][1].getLastRow() - 6;//タスクの数
  var tasks = ary[0][1].getRange(7, 1, taskn, 3).getDisplayValues();//タスクの名称
  var taskslog = ary[0][1].getRange(7, 4, taskn, 1).getNotes();//ログ（隠し列）
  for (var row = 0; row <= tasks.length - 1; row++) {
    logary[row + 4] = tasks[row][0] + "\n" + tasks[row][1] + "\n" + taskslog[row][0] + "\n" + tasks[row][2];//改行で区切り
  }

  //ログシート２行目に挿入
  sheetLog.insertRowBefore(2);
  Logger.log(logary);
  sheetLog.getRange(2, 1, 1, logary.length).setValues([logary]);//二次元配列にして書き込み
  if (sheetLog.getLastRow() >= 10001) {  //10001行以上なら10001行目を削除
    sheetLog.deleteRow(10001);
  }
  */

  //週タスク(古)にシートコピー（削除はしない）※基本バインダー_ログではなく個別シートで保管する場合。
  var newfilename = ary[0][1].getRange(3, 2).getDisplayValue() + "週作業";
  copyToNewSpreadsheet(ary[0][1], "1NEsWHJLqXgrWANW82Tybox04B85WYYp9", newfilename);//移動

  Logger.log("週タスク表編集・名前変更段階");
  //バインダー番号（２，２）セルは永久に変更しない。
  ary[0][1].getRange(3, 2).setValue(ary[0][2]);//作業週（３，２）セル変更
  ary[0][1].getRange(4, 2).setValue(ary[0][4]);//納品週（４，２）セル変更
  ary[0][1].setName(ary[0][5]);//シート名変更
  ary[0][1].getRange(2, 4).setValue("新");//順序バックアップ

  //[0]シート→デフォルト状態
  for (let r = 7; r <= ary[0][1].getLastRow(); r++) {//７行以降
    if (ary[0][1].getRange(r, 1).getDisplayValue() == "") {//タスクが空白なら不要
      ary[0][1].getRange(r, 2).setValue("不要");
    } else {//空白でなければ未
      ary[0][1].getRange(r, 2).setValue("未");
    }
    ary[0][1].getRange(r, 3).setValue("");//備考を空に
  }

  //メモクリア
  ary[0][1].clearNotes();

  //[1][2]も便宜上やる
  ary[1][1].getRange(3, 2).setValue(ary[1][2]);//作業週
  ary[1][1].getRange(4, 2).setValue(ary[1][4]);//納品週
  ary[1][1].setName(ary[1][5]);//シート名
  ary[1][1].getRange(2, 4).setValue("古");//順序バックアップ

  ary[2][1].getRange(3, 2).setValue(ary[2][2]);//作業週
  ary[2][1].getRange(4, 2).setValue(ary[2][4]);//納品週
  ary[2][1].setName(ary[2][5]);//シート名
  ary[2][1].getRange(2, 4).setValue("中");//順序バックアップ


}