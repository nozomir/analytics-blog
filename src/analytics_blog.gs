function fromGAToSummarizeSheet() {

  //スプレッドシートを取得
  var fromUrl = 'https://docs.google.com/spreadsheets/d/1WP6M2Kei2v953Lr4tn1-aHhSaDlzbYNs3jGu1N6kGK4/edit#gid=463221618';
  var spreadSheet = SpreadsheetApp.openByUrl(fromUrl);
  // レポートシートを取得
  var weeklyReports = [];
  for (var i = 1, max = spreadSheet.getSheets().length; i < max; i++) {
    weeklyReports.push(spreadSheet.getSheets()[i]);
  }

  // 出力先のスプレッドシートを取得
  var toUrl = 'https://docs.google.com/spreadsheets/d/13deRTQFun1kiSuc0o7Vy74-aUPUL-5KTnK6ByvSZjQY/edit#gid=1784139815';
  var toSheet = SpreadsheetApp.openByUrl(toUrl).getSheets()[0];

  var setColumn = 8;

  var COLUMN = 159;

  while (toSheet.getRange(COLUMN, setColumn).getValue() !== "") {
    setColumn++;
  }

  // 日付の設定
  if (toSheet.getRange(22, setColumn).getValue() === "") {
    toSheet.getRange(22, setColumn).setValue(getYesterday());
  }

  // スプレッドシートに出力
  for (var i = 0, max = weeklyReports.length; i < max; i++) {
    var uu = weeklyReports[i].getRange(12, 3).getValue();
    var pv = weeklyReports[i].getRange(12, 4).getValue();
    var organicSearches = weeklyReports[i].getRange(12, 6).getValue();
    var bounceRate = weeklyReports[i].getRange(12, 7).getValue() * 100;
    var repeatRate = 0;
    if (uu > 0) {
      repeatRate = (uu - weeklyReports[i].getRange(12, 5).getValue()) / uu * 100;
    }
    var mobilePv = 0;
    var twitterPv = 0;
    var socialPv = 0;

    for (var j = 16, row = weeklyReports[i].getMaxRows(); j <= row; j++) {
      var social = weeklyReports[i].getRange(j, 1).getValue();
      var device = weeklyReports[i].getRange(j, 2).getValue();
      if (social !== '(not set)' && social !== 'Twitter') {
        socialPv += weeklyReports[i].getRange(j, 4).getValue();
      }
      if (social === 'Twitter') {
        twitterPv += weeklyReports[i].getRange(j, 4).getValue();
      }
      if (device === 'tablet' || device === 'mobile') {
        mobilePv += weeklyReports[i].getRange(j, 4).getValue();
      }
    }
    var mobilePvRate = 0.0
    if (pv > 0) {
      mobilePvRate = mobilePv / pv * 100;
    }

    // 全体PVUUシートに書き出し
    toSheet.getRange(COLUMN + i * 8, setColumn).setValue(pv);
    toSheet.getRange(COLUMN + i * 8 + 1, setColumn).setValue(uu);
    toSheet.getRange(COLUMN + i * 8 + 2, setColumn).setValue(organicSearches);
    toSheet.getRange(COLUMN + i * 8 + 3, setColumn).setValue(twitterPv);
    toSheet.getRange(COLUMN + i * 8 + 4, setColumn).setValue(socialPv);
    toSheet.getRange(COLUMN + i * 8 + 5, setColumn).setValue(bounceRate);
    toSheet.getRange(COLUMN + i * 8 + 6, setColumn).setValue(repeatRate);
    toSheet.getRange(COLUMN + i * 8 + 7, setColumn).setValue(mobilePvRate);
  }

}

function getYesterday() {
  // 本日日付を取得
  var date = new Date();
  // 昨日の日付を取得
  date.setDate(date.getDate() - 1);

  return date;
}

function chatworkFromGA() {
  //スプレッドシートを取得
  var url = 'https://docs.google.com/spreadsheets/d/1WP6M2Kei2v953Lr4tn1-aHhSaDlzbYNs3jGu1N6kGK4/edit#gid=463221618';
  var spreadSheet = SpreadsheetApp.openByUrl(url);
  //Weeklyレポートシートを取得
  var weeklyReports = [];
  for (var i = 1, max = spreadSheet.getSheets().length; i < max; i++) {
    weeklyReports.push(spreadSheet.getSheets()[i]);
  }
  // 最終実行日付を取得
  var yDate = weeklyReports[0].getRange(2,2).getValue();
  // 昨日の日付を取得
  yDate.setDate(yDate.getDate() - 1);

  // 先週の日付を取得
  var lastweekDate = new Date();
  lastweekDate.setDate(yDate.getDate() - 6);

  // タイトル
  var title = "[title]" + Utilities.formatDate(lastweekDate, 'JST', 'yyyy/MM/dd') + "〜" + Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + "のレポート[/title]";

  var strBody = "[info]" + title;

  for (var i = 0, max = weeklyReports.length; i < max; i++) {
    var pv = weeklyReports[i].getRange(12, 4).getValue();
    var uu = weeklyReports[i].getRange(12, 3).getValue();
    var twitterPv = 0;
    var twitterUu = 0;
    strBody += (weeklyReports[i].getName() + " → " + // シート名
                pv + "PV, " + // ga:pageViews
                uu + "UU\n"); // ga:users
    for (var j = 16, maxLow = weeklyReports[i].getLastRow(); j <= maxLow; j++) {
      if (weeklyReports[i].getRange(j, 1).getValue() !== "Twitter") continue;
      twitterPv += weeklyReports[i].getRange(j, 4).getValue();
      twitterUu += weeklyReports[i].getRange(j, 3).getValue();
    }
    var twitterPvRate = 0;
    var twitterUuRate = 0;
    if (pv > 0) {
      twitterPvRate = myRound(twitterPv / pv * 100, 2);
    }
    if (uu > 0) {
      twitterUuRate = myRound(twitterUu / uu * 100, 2);
    }
    strBody += (weeklyReports[i].getName() + "(Twitter) → " + // シート名
               twitterPv + "PV(" + twitterPvRate + "％), " + // ga:pageViews
               twitterUu + "UU(" + twitterUuRate + "％)\n"); // ga:users
  }

  strBody += "[/info]";

  sendMessage(strBody);

}

function sendMessage(message) {
  var client = ChatWorkClient.factory({token: 'c2291484e09e9947c51622e2de8021b9'});　//チャットワークAPI
  client.sendMessage({
    room_id: 74752951, //ルームID(めぐんたさん)
    body: message});

  client.sendMessage({
    room_id: 94961753, //ルームID(やっくんさん)
    body: message});

  client.sendMessage({
    room_id: 97667125, //ルームID(ひとしさん)
    body: message});

  //client.sendMessage({
  //  room_id: 65174489, //ルームID(のんたん)
  //  body: message});
}

function myRound(val, precision){
     //小数点を移動させる為の数を10のべき乗で求める
     //例) 小数点以下2桁の場合は 100 をかける必要がある
     digit = Math.pow(10, precision);

     //四捨五入したい数字に digit を掛けて小数点を移動
     val = val * digit;

     //roundを使って四捨五入
     val = Math.round(val);

     //移動させた小数点を digit で割ることでもとに戻す
     val = val / digit;

     return val;
}
