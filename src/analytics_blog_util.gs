function fromGAToSummarizeSheet(fromUrl, row) {

  //スプレッドシートを取得
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

  while (toSheet.getRange(row, setColumn).getValue() !== "") {
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
    toSheet.getRange(row + i * 8, setColumn).setValue(pv);
    toSheet.getRange(row + i * 8 + 1, setColumn).setValue(uu);
    toSheet.getRange(row + i * 8 + 2, setColumn).setValue(organicSearches);
    toSheet.getRange(row + i * 8 + 3, setColumn).setValue(twitterPv);
    toSheet.getRange(row + i * 8 + 4, setColumn).setValue(socialPv);
    toSheet.getRange(row + i * 8 + 5, setColumn).setValue(bounceRate);
    toSheet.getRange(row + i * 8 + 6, setColumn).setValue(repeatRate);
    toSheet.getRange(row + i * 8 + 7, setColumn).setValue(mobilePvRate);
  }

}

function getYesterday() {
  // 本日日付を取得
  var date = new Date();
  // 昨日の日付を取得
  date.setDate(date.getDate() - 1);

  return date;
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
