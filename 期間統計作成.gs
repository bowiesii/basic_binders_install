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

  var folderId = "19t8-VEtn-LQ4pIP2Tdet4Cd7mco815YE";//出力フォルダは固定
  var date1 = new Date(startD);
  var date2 = new Date(endD);
  date2.setDate(date2.getDate() + 1);//24:00なので日付的には次の日の00:00になる

  var { editN, ptN, newSpS } = makeSummalyInPeriod(date1, date2, folderId, plusFileName);

  var filename = newSpS.getName();
  var fileUrl = bbsLib.toUrl(newSpS.getId(), "0");

  mail_summary(email, filename, fileUrl);//本人へメール

}


//特定の期間の統計を作成（date1からdate2まで）date1date2はDateフォーマットで。
//統合ログの実行者別回数をまとめる。
//フォルダに出力する。
function makeSummalyInPeriod(date1, date2, folderId, plusFileName) {

  //行数を調べる
  var row_1 = dateToRow(date1);
  var row_2 = dateToRow(date2);
  var row_N = row_2 - row_1;

  if (row_N == 0) {//データが無いので週報は作らない
    Logger.log("no data in that period");
    return;
  }

  const sheetRaw = bbsLib.getSheetByIdGid(id_bbLog, gid_intLog);//統合ログ（ソースデータはこれだけとする。混乱回避のため日ごとデータは使わない）
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));//行列入れ替える関数

  //週報ファイルを作る
  var rawAry_r1 = sheetRaw.getRange(1, 1, 1, 12).getValues();//生データ１行目
  var rawAry = sheetRaw.getRange(row_1 + 1, 1, row_N, 12).getValues();//生データ
  rawAry.unshift(rawAry_r1);//先頭に目次を追加
  var uAry = [];//ユーザーごとのまとめ
  uAry[0] = ["実行者（simei）(0)", "実行者番号(1)", "編集数合計(2)", "発注編集数(3)", "週バ編集数(4)", "鮮度編集数(5)", "清掃編集数(6)", "新人編集数(7)", "ポイント合計(8)", "発注pt(9)", "週バpt(10)", "鮮度pt(11)", "清掃pt(12)", "新人pt(13)"];

  for (let r = 1; r <= rawAry.length - 1; r++) {//１行目はスルー
    let user = rawAry[r][1];//氏名※無ければ空文字
    let userN = rawAry[r][2];//氏名ナンバー※無ければ空文字
    let sN = rawAry[r][3];//シート名
    let pt = rawAry[r][8];//ポイント

    if (userN == "") {
      userN = -1;
      user = "#不明";
    }

    if (userN == 3) {//店機器のとき、userから探す

      for (let rr = 1; rr <= uAry.length - 1;) {//１行目はスルーして探す

        if (user != uAry[rr][0]) {//実行者名違ってたらスルーして次
          if (rr == uAry.length - 1) {//でも最後じゃんこれ
            uAry[rr] = [[user], userN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
            rr++;
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

      for (let rr = 1; rr <= uAry.length - 1;) {//１行目はスルーして探す

        if (userN != uAry[rr][1]) {//実行者番号違ったらスルーして次
          if (rr == uAry.length - 1) {//でも最後じゃんこれ
            uAry[rr] = [[user], userN, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];//あらたな配列→下へ続く
            rr++;
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
            uAry[rr][0], push(user);
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
  uAry[1][rowU] = "###";
  for (let r = 2; r <= uAry.length - 1; r++) {
    uAry[r][rowU] = rowSum(uAry[r]);
  }
  uAry = transpose(uAry);//行列入れ替え
  //これでuAryが完成

  for (let r = 0; r <= uAry.length - 1; r++) {
    Logger.log(uAry[r]);
  }

  //報告ファイルを作成(rawAry,uAry)
  const date1_f = Utilities.formatDate(date1, "JST", "yyyy/MM/dd(E)HH:mm");
  const date2_f = Utilities.formatDate(date2, "JST", "yyyy/MM/dd(E)HH:mm");
  var newFileName = plusFileName + "_" + date1_f + "_から_" + date2_f + "_まで";

  var newSpS = copyAryToNewSpreadSheet(uAry, folderId, newFileName, "ユーザーごとの統計");
  copyAryToSpreadSheet(newSpS, rawAry, "生データ");

  var editN = uAry[rowU][2];//編集数合計
  var ptN = uAry[rowU][8];//ポイント合計

  return { editN, ptN, newSpS };//作ったスプシも返す


}

