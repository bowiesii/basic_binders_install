//基本バインダー
const id_bb = "1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg";
const bbSpreadSheet = SpreadsheetApp.openById(id_bb);
const gid_order = "648587868";//発注
const gid_orderOld = "318850821";//発注（古）
const gid_wtask1 = "1616634719";//週タスク１のGID
const gid_wtask2 = "2097376321";//週タスク２
const gid_wtask3 = "1024816661";//週タスク３
const gid_fcheck = "1405667253";//鮮度
const gid_h_edit = "1020343545";//h_編集数一時
const gid_h_simei = "402406560";//h_氏名今日
const gid_h_wtask = "997759008";//h_週タスク今日
const gid_h_fcheckPlace = "1413518585";//h_鮮度箇所
const gid_h_fcheck = "2006815737";//h_鮮度今日
const gid_h_sinjin = "247760006";//h_新人

//新人原本
const id_sinjinOri = "10Lo84sEyGnuYzhYuJhlfodEMmJM-lpPJU_tTnj7AD9o"
const gid_sinjinOri = "865253827";


//基本バインダーログ
const id_bbLog = "17bZ83U_NeHXLT__NOV0zfHd2B8XZIBEgKfd_akNZDuY";
const gid_simei = "1508872407";//氏名日ごと
const gid_wtaskDay = "1990699055";//週タスク日ごと
const gid_wtaskWeek = "0";//週タスク週ごと
const gid_fcheckDay = "547488950";//鮮度日ごと
const gid_fcheckDaySum = "1585149555";//鮮度日ごと統計
const gid_orderDay = "838294828";//発注日ごと
const gid_editDay = "1153461632";//編集数日ごと
const gid_editDaySum = "588078675";//編集数日ごと統計
const gid_sinjinList = "689879716";//新人リスト
const gid_sinjinDay = "239524908";//新人日ごと

//本日日付定義
//const today = new Date("2023/8/7");
//やっぱり本当の日付
const today = new Date();
const todayYear = today.getFullYear();
const todayMonth = today.getMonth() + 1;
const todayDate = today.getDate();
const wary = ["日", "月", "火", "水", "木", "金", "土"];
const today_wjpn = wary[today.getDay()];
const today_ymd = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
const today_ymdd = today_ymd + " " + today_wjpn;
const today_hm = Utilities.formatDate(today, 'JST', 'HH:mm');
const today_ymddhm = today_ymdd + " " + today_hm;
