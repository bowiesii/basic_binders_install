//対象フォームへトリガー設定（１回走ればよし。）
//function setTrigger() {
//var file = FormApp.openById("1PcKmi_c2LQbtbpi1WELpmCgT-GyCOwm7L2c0KP6bG9c");//作成フォーム：済
//var functionName = "newSummaly"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

//フォームからトリガー
function newSummaly(e) {

  var email = e.response.getRespondentEmail();
  var answer = e.response.getItemResponses();
  var plusFileName = answer[0].getResponse();//追加文字列
  var startD = answer[1].getResponse();//開始日付
  var endD = answer[2].getResponse();//終了日付

  //共有登録されていなければスルー→メール
  if (shareingOrNot(email) == false) {
    Logger.log("共有登録されていない");
    mail_sinjin(email, "", 2);//本人
    return;
  }

  var date1 = new Date(startD);//→00;00にする
  var date2 = new Date(endD);//→23:59にする
  date1.setHours(0);
  date2.setHours(23);
  date2.setMinutes(59);

  var folderId = "19t8-VEtn-LQ4pIP2Tdet4Cd7mco815YE";//出力フォルダは固定

  var { editN, ptN, newSpS } = makeSummalyInPeriod(date1, date2, folderId, plusFileName);

  if (newSpS == null) {//作成されなかった。
    mail_makeSummary(0, email, "", "");//本人へメール
  }

  const file = DriveApp.getFileById(newSpS.getId());

  //編集権限を付与
  file.addEditor(email);
  file.setShareableByEditors(false);//★編集者による権限操作を禁止する

  var filename = newSpS.getName();
  var fileUrl = bbsLib.toUrl(newSpS.getId(), "0");

  mail_makeSummary(1, email, filename, fileUrl);//本人へメール

}


//特定の期間の統計を作成（date1からdate2まで）date1date2はDateフォーマットで。
//統合ログの実行者別回数をまとめる。
//フォルダに出力する。
function makeSummalyInPeriod(date1, date2, folderId, plusFileName) {

  //テスト用
  //var date1 = new Date("2023/8/13 00:00");
  //var date2 = new Date("2023/8/14 23:59");
  //var folderId = "19t8-VEtn-LQ4pIP2Tdet4Cd7mco815YE";
  //var plusFileName = "test";

  Logger.log("date1=" + date1);
  Logger.log("date2=" + date2);

  //行数を調べる
  var row_1 = dateToRow(date1);
  var row_2 = dateToRow(date2);
  var row_N = row_1 - row_2;
  row_2 = row_2 + 1;

  Logger.log("row_2=" + row_2 + ",row_N=" + row_N);

  if (row_N == 0) {//データが無いので作らない
    Logger.log("no data in that period");
    let editN = null;
    let ptN = null;
    let newSpS = null;
    return { editN, ptN, newSpS };
  }

  const sheetRaw = bbsLib.getSheetByIdGid(id_bbLog, gid_intLog);//統合ログ（ソースデータはこれだけとする。混乱回避のため日ごとデータは使わない）
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));//行列入れ替える関数

  //生データを作る
  var rawAry_r1 = sheetRaw.getRange(1, 1, 1, 15).getDisplayValues();//生データ１行目（２次元）
  var rawAry = sheetRaw.getRange(row_2, 1, row_N, 15).getDisplayValues();//生データ
  rawAry.unshift(rawAry_r1[0]);//先頭に目次を追加
  Logger.log("rawAry");
  for (let r = 0; r <= rawAry.length - 1; r++) {
    Logger.log(rawAry[r]);
  }

  ////////ユーザー別統計
  var uAry = [];//ユーザーごとのまとめ
  uAry[0] = ["実行者（simei）(0)", "実行者番号（simeiN）(1)", "編集数合計(2)", "発注編集数(3)", "週バ編集数(4)", "鮮度編集数(5)", "清掃編集数(6)", "新人編集数(7)", "ポイント合計(8)", "発注pt(9)", "週バpt(10)", "鮮度pt(11)", "清掃pt(12)", "新人pt(13)"];

  for (let r = 1; r <= rawAry.length - 1; r++) {//１行目はスルー
    let user = rawAry[r][4];//氏名※無ければ空文字
    let userN = rawAry[r][5];//氏名ナンバー※無ければ空文字
    let sN = rawAry[r][6];//シート名
    let pt = Number(rawAry[r][11]);//ポイント

    if (userN == "" || userN == null || userN == undefined) {
      userN = -1;
    }
    if (user == "" || user == null || user == undefined) {
      user = "？";
    }


    if (userN == 3) {//店機器のとき、userから探す

      for (let rr = 0; rr <= uAry.length - 1;) {//探す

        if (user != uAry[rr][0]) {//実行者名違ってたらスルーして次
          if (rr == uAry.length - 1) {//でも最後じゃんこれ
            rr++;
            uAry[rr] = [[user], userN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
          } else {
            rr++;
            continue;
          }
        }

        //氏名が一致
        uAry[rr][2]++;//合計＋１
        uAry[rr][8] = uAry[rr][8] + pt;//ポイント加算
        if (sN == "発注") {
          uAry[rr][3]++;
          uAry[rr][9] = uAry[rr][9] + pt;
        } else if (sN.includes("週バ")) {
          uAry[rr][4]++;
          uAry[rr][10] = uAry[rr][10] + pt;
        } else if (sN == "鮮度") {
          uAry[rr][5]++;
          uAry[rr][11] = uAry[rr][11] + pt;
        } else if (sN == "清掃") {
          uAry[rr][6]++;
          uAry[rr][12] = uAry[rr][12] + pt;
        } else if (sN.includes("【新】")) {
          uAry[rr][7]++;
          uAry[rr][13] = uAry[rr][13] + pt;
        }

        break;//見つかったので

      }

    } else {//店機器でないとき、userNから探す

      for (let rr = 0; rr <= uAry.length - 1;) {//探す

        if (userN != uAry[rr][1]) {//実行者番号違ったらスルーして次
          if (rr == uAry.length - 1) {//でも最後じゃんこれ
            rr++;
            uAry[rr] = [[user], userN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
          } else {
            rr++;
            continue;
          }
        }

        //氏名番号が一致
        uAry[rr][2]++;//合計＋１
        uAry[rr][8] = uAry[rr][8] + pt;//ポイント加算
        if (sN == "発注") {
          uAry[rr][3]++;
          uAry[rr][9] = uAry[rr][9] + pt;
        } else if (sN.includes("週バ")) {
          uAry[rr][4]++;
          uAry[rr][10] = uAry[rr][10] + pt;
        } else if (sN == "鮮度") {
          uAry[rr][5]++;
          uAry[rr][11] = uAry[rr][11] + pt;
        } else if (sN == "清掃") {
          uAry[rr][6]++;
          uAry[rr][12] = uAry[rr][12] + pt;
        } else if (sN.includes("【新】")) {
          uAry[rr][7]++;
          uAry[rr][13] = uAry[rr][13] + pt;
        }

        //uAry[rr][0]（[user]配列）にuserが無ければ追加する
        for (let rrr = 0; rrr <= uAry[rr][0].length - 1; rrr++) {
          if (uAry[rr][0][rrr] == user) {
            break;
          } else if (rrr == uAry[rr][0].length - 1) {//最後じゃんこれ→配列に追加
            uAry[rr][0].push(user);
          }
        }

        break;//見つかったので

      }

    }//店機器でないとき

  }

  //uAry[rr][0]を文字列に直す
  for (let rr = 1; rr <= uAry.length - 1; rr++) {
    var st = "";
    for (let rrr = 0; rrr <= uAry[rr][0].length - 1; rrr++) {
      st = st + "#" + uAry[rr][0][rrr];
    }
    uAry[rr][0] = st;
  }

  //uAryに全行合計を追加
  uAry = transpose(uAry);//行列入れ替え
  var rowU = uAry[0].length;
  uAry[0][rowU] = "合計";
  uAry[1][rowU] = "#";
  for (let r = 2; r <= uAry.length - 1; r++) {
    uAry[r][rowU] = rowSum(uAry[r]);
  }
  uAry = transpose(uAry);//行列入れ替え
  //これでuAryが完成

  Logger.log("uAry");
  for (let r = 0; r <= uAry.length - 1; r++) {
    Logger.log(uAry[r]);
  }
  ////////ユーザー別統計おわり


  ////////日・シフト種別統計
  var dhAry = [];//日・シフト種ごとのまとめ
  dhAry[0] = ["日(0)", "シフト種類(1)", "編集数合計(2)", "発注編集数(3)", "週バ編集数(4)", "鮮度編集数(5)", "清掃編集数(6)", "新人編集数(7)", "ポイント合計(8)", "発注pt(9)", "週バpt(10)", "鮮度pt(11)", "清掃pt(12)", "新人pt(13)"];

  for (let r = 1; r <= rawAry.length - 1; r++) {//１行目はスルー
    let shiftD = rawAry[r][1];//シフト日
    let shiftN = rawAry[r][2];//シフト時間帯
    let sN = rawAry[r][6];//シート名
    let pt = Number(rawAry[r][11]);//ポイント

    for (let rr = 0; rr <= dhAry.length - 1;) {//探す

      if (shiftD != dhAry[rr][0] || shiftN != dhAry[rr][1]) {//日もしくは時間帯が違ってたら次
        if (rr == dhAry.length - 1) {//でも最後じゃんこれ
          rr++;
          dhAry[rr] = [shiftD, shiftN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
        } else {
          rr++;
          continue;
        }
      }

      //日、時間帯が一致
      dhAry[rr][2]++;//合計＋１
      dhAry[rr][8] = dhAry[rr][8] + pt;//ポイント加算
      if (sN == "発注") {
        dhAry[rr][3]++;
        dhAry[rr][9] = dhAry[rr][9] + pt;
      } else if (sN.includes("週バ")) {
        dhAry[rr][4]++;
        dhAry[rr][10] = dhAry[rr][10] + pt;
      } else if (sN == "鮮度") {
        dhAry[rr][5]++;
        dhAry[rr][11] = dhAry[rr][11] + pt;
      } else if (sN == "清掃") {
        dhAry[rr][6]++;
        dhAry[rr][12] = dhAry[rr][12] + pt;
      } else if (sN.includes("【新】")) {
        dhAry[rr][7]++;
        dhAry[rr][13] = dhAry[rr][13] + pt;
      }

      break;//見つかったので

    }

  }

  //dhAryに全行合計を追加
  dhAry = transpose(dhAry);//行列入れ替え
  var rowU = dhAry[0].length;
  dhAry[0][rowU] = "合計";
  dhAry[1][rowU] = "#";
  for (let r = 2; r <= dhAry.length - 1; r++) {
    dhAry[r][rowU] = rowSum(dhAry[r]);
  }
  dhAry = transpose(dhAry);//行列入れ替え
  //これでdhAryが完成

  Logger.log("dhAry");
  for (let r = 0; r <= dhAry.length - 1; r++) {
    Logger.log(dhAry[r]);
  }
  ////////日・時間帯別統計おわり


  ////////日別統計
  var dAry = [];//日ごとのまとめ
  dAry[0] = ["日(0)", "#(1)", "編集数合計(2)", "発注編集数(3)", "週バ編集数(4)", "鮮度編集数(5)", "清掃編集数(6)", "新人編集数(7)", "ポイント合計(8)", "発注pt(9)", "週バpt(10)", "鮮度pt(11)", "清掃pt(12)", "新人pt(13)"];

  for (let r = 1; r <= rawAry.length - 1; r++) {//１行目はスルー
    let shiftD = rawAry[r][1];//シフト日
    let sN = rawAry[r][6];//シート名
    let pt = Number(rawAry[r][11]);//ポイント

    for (let rr = 0; rr <= dAry.length - 1;) {//探す

      if (shiftD != dAry[rr][0]) {//日が違ってたら次
        if (rr == dAry.length - 1) {//でも最後じゃんこれ
          rr++;
          dAry[rr] = [shiftD, "#", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
        } else {
          rr++;
          continue;
        }
      }

      //日が一致
      dAry[rr][2]++;//合計＋１
      dAry[rr][8] = dAry[rr][8] + pt;//ポイント加算
      if (sN == "発注") {
        dAry[rr][3]++;
        dAry[rr][9] = dAry[rr][9] + pt;
      } else if (sN.includes("週バ")) {
        dAry[rr][4]++;
        dAry[rr][10] = dAry[rr][10] + pt;
      } else if (sN == "鮮度") {
        dAry[rr][5]++;
        dAry[rr][11] = dAry[rr][11] + pt;
      } else if (sN == "清掃") {
        dAry[rr][6]++;
        dAry[rr][12] = dAry[rr][12] + pt;
      } else if (sN.includes("【新】")) {
        dAry[rr][7]++;
        dAry[rr][13] = dAry[rr][13] + pt;
      }

      break;//見つかったので

    }

  }

  //dAryに全行合計を追加
  dAry = transpose(dAry);//行列入れ替え
  var rowU = dAry[0].length;
  dAry[0][rowU] = "合計";
  dAry[1][rowU] = "#";
  for (let r = 2; r <= dAry.length - 1; r++) {
    dAry[r][rowU] = rowSum(dAry[r]);
  }
  dAry = transpose(dAry);//行列入れ替え
  //これでdAryが完成

  Logger.log("dAry");
  for (let r = 0; r <= dAry.length - 1; r++) {
    Logger.log(dAry[r]);
  }
  ////////日別統計おわり


  ////////シフト種別統計
  var hAry = [];//シフト種類ごとのまとめ
  hAry[0] = ["#(0)", "シフト種類(1)", "編集数合計(2)", "発注編集数(3)", "週バ編集数(4)", "鮮度編集数(5)", "清掃編集数(6)", "新人編集数(7)", "ポイント合計(8)", "発注pt(9)", "週バpt(10)", "鮮度pt(11)", "清掃pt(12)", "新人pt(13)"];

  for (let r = 1; r <= rawAry.length - 1; r++) {//１行目はスルー
    let shiftN = rawAry[r][2];//時間帯名
    let sN = rawAry[r][6];//シート名
    let pt = Number(rawAry[r][11]);//ポイント

    for (let rr = 0; rr <= hAry.length - 1;) {//探す

      if (shiftN != hAry[rr][1]) {//日が違ってたら次
        if (rr == hAry.length - 1) {//でも最後じゃんこれ
          rr++;
          hAry[rr] = ["#", shiftN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
        } else {
          rr++;
          continue;
        }
      }

      //時間帯が一致
      hAry[rr][2]++;//合計＋１
      hAry[rr][8] = hAry[rr][8] + pt;//ポイント加算
      if (sN == "発注") {
        hAry[rr][3]++;
        hAry[rr][9] = hAry[rr][9] + pt;
      } else if (sN.includes("週バ")) {
        hAry[rr][4]++;
        hAry[rr][10] = hAry[rr][10] + pt;
      } else if (sN == "鮮度") {
        hAry[rr][5]++;
        hAry[rr][11] = hAry[rr][11] + pt;
      } else if (sN == "清掃") {
        hAry[rr][6]++;
        hAry[rr][12] = hAry[rr][12] + pt;
      } else if (sN.includes("【新】")) {
        hAry[rr][7]++;
        hAry[rr][13] = hAry[rr][13] + pt;
      }

      break;//見つかったので

    }

  }

  //hAryに全行合計を追加
  hAry = transpose(hAry);//行列入れ替え
  var rowU = hAry[0].length;
  hAry[0][rowU] = "合計";
  hAry[1][rowU] = "#";
  for (let r = 2; r <= hAry.length - 1; r++) {
    hAry[r][rowU] = rowSum(hAry[r]);
  }
  hAry = transpose(hAry);//行列入れ替え
  //これでhAryが完成

  Logger.log("hAry");
  for (let r = 0; r <= hAry.length - 1; r++) {
    Logger.log(hAry[r]);
  }
  ////////時間帯別統計おわり


  //報告ファイルを作成(rawAry,uAry,dhAry,dAry,hAry)
  const date1_f = Utilities.formatDate(date1, "JST", "yyyy/MM/dd(E)HH:mm");
  const date2_f = Utilities.formatDate(date2, "JST", "yyyy/MM/dd(E)HH:mm");
  var newFileName = plusFileName + "_" + date1_f + "_から_" + date2_f + "_まで";

  var newSpS = copyAryToNewSpreadSheet(rawAry, folderId, newFileName, "生データ");
  copyAryToSpreadSheet(newSpS, uAry, "実行者別統計");
  copyAryToSpreadSheet(newSpS, dhAry, "日・シフト種類別統計");
  copyAryToSpreadSheet(newSpS, dAry, "日別統計");
  copyAryToSpreadSheet(newSpS, hAry, "シフト種類別統計");

  var editN = hAry[rowU][2];//編集数合計
  var ptN = hAry[rowU][8];//ポイント合計

  return { editN, ptN, newSpS };//作ったスプシも返す


}

