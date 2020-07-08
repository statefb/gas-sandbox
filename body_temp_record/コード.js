/*
 * LINE botで受け取った情報をスプレッドシートに書き込む
 */

// line channel access token
var ACCESS_TOKEN = "xxxxxxxxxxxxxxxxxxxx"
// spread sheet id
var ID = "xxxxxxxxxxxxxxxxxxxx"

function getSheet(sheet_name) {
  // スプレッドシートオブジェクトをシート名から取得する
  var sheet = SpreadsheetApp.openById(ID).getSheetByName(sheet_name)
  return sheet
}

function getSourceUser(e) {
  // 送信元ユーザー情報（webhook ID、ユーザ名）を取得する
  var sourceId = JSON.parse(e.postData.contents).events[0].source.userId;
  var sourceUser = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + sourceId, {
    "method": "get",
    "headers": {
      'Authorization': 'Bearer ' + ACCESS_TOKEN
    }
  });
  var sourceName = JSON.parse(sourceUser).displayName

  return {
    "id": sourceId,
    "name": sourceName
  }
}

function getMessage(e) {
  // ユーザーが送信したメッセージを取得する
  var userMessage = JSON.parse(e.postData.contents).events[0].message.text;
  return userMessage;
}

function isMessageValid(text) {
  // 入力内容が適切か否かをチェックする
  var num = parseInt(text, 10);
  if (isNaN(num)) {
    // 数値に変換できないものはNG
    return false
    //　体温として現実的な数値でないものはNG
  } else if (num < 35.0 || num > 40.0) {
    return false
  } else {
    return true
  }
}


function sendMessage(e, text) {
  // メッセージを送信する
  // 参考：https://developers.line.biz/ja/reference/messaging-api/#webhook-event-objects
  var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  var userId = JSON.parse(e.postData.contents).events[0].source.userId;

  var url = "https://api.line.me/v2/bot/message/push"
  var headers = {
    "Content-Type": "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + ACCESS_TOKEN,
  };

  var postData = {
    "to": userId,
    "messages": [{
      'type': 'text',
      'text': text
    }]
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };

  return UrlFetchApp.fetch(url, options);
}

function insertOrUpdate(inputData) {
  // データを入力または更新する
  // 参考：https://qiita.com/rf_p/items/f423ed2b23b789b2ebe9
  var sheet = getSheet("シート1")
  var column = findColumn(sheet, inputData.userName)

  if (!column.found) {
    sheet.getRange(1, column.index).setValue(inputData.userName);
  }

  var row = findRow(sheet, inputData.date)

  if (row.found) {
    sheet.getRange(row.index, column.index).setValue(inputData.temp)
  } else {
    sheet.getRange(row.index, 1).setValue(inputData.date)
    sheet.getRange(row.index, column.index).setValue(inputData.temp)
  }

}

function findRow(sheet, date) {
  // 行番号を返す
  var searchDate = Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy/MM/dd');
  var values = sheet.getDataRange().getValues();

  for (var i = values.length - 1; i > 0; i--) {
    var dataDate = Utilities.formatDate(new Date(values[i][0]), 'Asia/Tokyo', 'yyyy/MM/dd');
    if (dataDate == searchDate) {
      return {
        "found": true,
        "index": i + 1
      }
    }
  }
  return {
    "found": false,
    "index": values.length + 1
  }
}

function findColumn(sheet, userName) {
  // 列番号を返す
  var values = sheet.getDataRange().getValues();

  for (var i = 0; i < values[0].length; i++) {
    var name = values[0][i];
    if (name === userName) {
      return {
        "found": true,
        "index": i + 1
      }
    }
  }
  return {
    "found": false,
    "index": values[0].length + 1
  }
}

function doPost(e) {
  //ユーザーがLINEにメッセージ送信した時の処理

  //送信元ユーザー情報
  var sourceUser = getSourceUser(e);
  //ユーザーのメッセージを取得
  var message = getMessage(e);
  // メッセージ内容のバリデーション
  if (!isMessageValid(message)) {
    sendMessage(e, '入力内容が不正です。現実的な体温（数字のみ）を入力してください。例："36.5"など')
    return
  }
  // 本日の日付
  var todayDate = new Date();
  var todayString = Utilities.formatDate(todayDate, 'Asia/Tokyo', 'yyyy/MM/dd')
  // 入力するデータ
  var inputData = {
    "date": todayString,
    "userName": sourceUser.name,
    "temp": message
  }

  // データを入力or更新する
  insertOrUpdate(inputData)

  // ユーザーに登録した旨を通知する
  sendMessage(e, '登録完了しました');
};