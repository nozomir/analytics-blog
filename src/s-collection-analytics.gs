var FROM_URL = 'https://docs.google.com/spreadsheets/d/1WP6M2Kei2v953Lr4tn1-aHhSaDlzbYNs3jGu1N6kGK4/edit#gid=463221618'

function fromGAToSummarizeSheetForSCollection() {
  fromGAToSummarizeSheet(FROM_URL, 159);
}

function chatworkFromGA() {
  //スプレッドシートを取得
  var url = FROM_URL;
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
