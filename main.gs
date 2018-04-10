var webhook_url = "https://hooks.slack.com/services/xxxxxxxxxxxx/xxxxxxxxxxxx/xxxxxxxxxxxx";

var sheet_url   = "https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxxx";
var spreadSheet = SpreadsheetApp.openByUrl(sheet_url);
var sheet       = spreadSheet.getSheetByName("朝会当番");

var payload = {
  'url'         : 'https://slack.com/api/chat.postMessage',
  'username'    : "せんとくん",
  'icon_emoji'  : ":sentokun_stamp:",
  'channel'     : "#hogehoge",
  'link_names'  : 1,
};

function sentokun() {
  var date = new Date();
  
  if( is_holiday(date) ) {
    return;
  }
  
  var today = Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd");

  var startRow = 2;
  var lastRow  = sheet.getLastRow();
  var lastCol  = sheet.getLastColumn();

  var youbiArr     = sheet.getSheetValues(startRow, 1, lastRow, 1);
  var dateArr      = sheet.getSheetValues(startRow, 2, lastRow, 2);  
  var nameArr      = sheet.getSheetValues(startRow, 3, lastRow, 3);
  var slackIdArr   = sheet.getSheetValues(startRow, 4, lastRow, 4);
  var shikaiArr    = sheet.getSheetValues(startRow, 6, lastRow, 6);
  var mainThemeArr = sheet.getSheetValues(startRow, 7, lastRow, 7);
  var subThemeArr  = sheet.getSheetValues(startRow, 8, lastRow, 8);
  var koshienArr   = sheet.getSheetValues(startRow, 9, lastRow, 9);

  for( var i = 0; i < dateArr.length; i++ ) {

    var selectDate = new Date(dateArr[i][0]);
    var selectToday = Utilities.formatDate(selectDate, "Asia/Tokyo", "yyyy/MM/dd");

    if( selectToday == today　) {
      var youbi     = youbiArr[i+1][0];
      var name      = nameArr[i+1][0];
      var slackID   = slackIdArr[i+1][0];
      var mainTheme = mainThemeArr[i+1][0];
      var subTheme  = subThemeArr[i+1][0];
      var shikai    = shikaiArr[i+1][0];

      var text = "";

      for( var j = 1; j<= 7; j++ ){
        var t = i+j+1;
        if (nameArr[t][0] == ""){
          text = text + (new Date(dateArr[t][0])).toLocaleDateString() +'(' + youbiArr[t][0] + ') :\t :kotori: スピーチお休み :kotori: \n';
        } else {
          text = text + (new Date(dateArr[t][0])).toLocaleDateString() +'(' + youbiArr[t][0] + ') :\t' + nameArr[t][0] + 'さん\n';
        }
      }

      if( name == "" || slackID == "" ) {
        var message = "今日は朝会スピーチお休みです :frog:"
      } else {
        var message = "明日の朝会スピーチは" + name + "さん `" + slackID + "` だよ。\nスピーチするのが無理な場合などは司会の `" + shikai + "` に相談してね！";
      }

      var attachments =  [
        {
          title  : "今後の朝会当番",
          title_link: sheet_url,
          text   : text,
          fields : [
            {
              title : "テーマ①",
              value : mainTheme,
              short :true,
            },
            {
              title : "テーマ②",
              value : subTheme,
              short : true,
            }],
        }
      ];
      post_to_slack(message, attachments);
    }
    
  }
  
}

function is_holiday(date) {
  var tommorow = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
  // 翌日土日判定
  if (tommorow.getDay() == 0 || tommorow.getDay() == 6) {
    return true; 
  }
  // 翌日祝日判定
  var calendar   = CalendarApp.getCalendarById("en.japanese#holiday@group.v.calendar.google.com");
  var holiday    = calendar.getEventsForDay(tommorow);
  if (holiday.length > 0 ) {
    return true;
  }
  return false;
}

function post_to_slack(message, attachments) {

  payload.text        = message;
  payload.attachments = attachments;

  var params = {
    'method' : "post",
    'payload': JSON.stringify(payload),
    };
 
  var res = UrlFetchApp.fetch(webhook_url, params);
}
