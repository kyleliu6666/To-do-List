// 推播每日Todo
function todoPush() {
  let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss = ss.getSheets()[0];
  let lastRow = ss.getLastRow(); //目前最後一行
  let alertDay = 10; //倒數提醒日
  let value = ss.getRange(lastDayCol+'2:'+lastDayCol + lastRow).getValues();
  let valueLength = value.length; //資料長度
  let message = '';
  
  let dataArray = []
  for(let i=0;i<valueLength;i++){
    dataArray.push(value[i][0]);
  }

  dataArray.forEach(function(data,index) {
    if(data <= alertDay & data >= 0){
      dayend = ss.getRange(dateCol + (index + 2)).getValues()[0][0];
      dayend = Utilities.formatDate(dayend, 'Asia/Taipei', 'yyyy/MM/dd');
      thisweekday = ss.getRange(weekday + (index + 2)).getValues()[0][0];
      thistime = ss.getRange(beginCol + (index + 2)).getValues()[0][0];
      detail = ss.getRange(mission + (index + 2)).getValues()[0][0];
      message += ` \n${dayend}(${thisweekday}) ${thistime}\n${detail}\n倒數 ${data}天\n`;
      console.log(dayend,thisweekday,thistime,detail,data);
    }
  });
  message = message.trimEnd();

  //有資料才POST到LINE Notify,Telegram
  if(message != ''){
    //sendLineNotify(message,diaryToken);
    sendToTelegram('【每日提醒】'+ message,notifybotToken);
    sendLinePushMessage('【每日提醒】'+ message);
  }
}

function sendToTelegram(text,botToken) {

  //https://api.telegram.org/bot<YourBOTToken>/sendMessage?chat_id=<YourChatID>&text=Hello

  // Telegram API URL
  let url = "https://api.telegram.org/bot" + botToken + "/sendMessage";
  
  // 準備發送的資料
  let data = {
    "chat_id": chatId,
    "text": text,
    "parse_mode": "HTML"  // 可選：支援 HTML 格式
  };
  
  // 設定 HTTP 請求選項
  let options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(data)
  };
  
  // 發送請求
  try {
    let response = UrlFetchApp.fetch(url, options);
    Logger.log("訊息發送成功：" + response.getContentText());
  } catch(e) {
    Logger.log("發送失敗：" + e.toString());
  }
}


// 發送LINE Notify通知
function sendLineNotify(message, token){
  let options =
  {
    "method"  : "post",
    "payload" : {"message" : message},
    "headers" : {"Authorization" : "Bearer " + token}
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

// 發送LINE LineMessage通知
function sendLinePushMessage(text) {

  // 要發送訊息的使用者 ID
  const USER_ID = 'U63f02493f5f166a0ec5d6ca404e63833';
  
  // LINE Push Message API 端點
  const LINE_ENDPOINT = 'https://api.line.me/v2/bot/message/push';
  
  // 要發送的訊息內容
  const message = {
    to: USER_ID, //User ID, Group ID, Room ID
    messages: [{
      type: 'text',
      text: text
    }]
  };
  
  // 設定 HTTP 請求選項
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    'payload': JSON.stringify(message)
  };
  
  try {
    // 發送請求
    const response = UrlFetchApp.fetch(LINE_ENDPOINT, options);
    Logger.log('訊息發送成功：' + response.getContentText());
  } catch(error) {
    Logger.log('發生錯誤：' + error);
  }
}





// 處理LINE訊息API的POST請求
function doPost(e) {
  let replyToken= JSON.parse(e.postData.contents).events[0].replyToken;
  let userMessage= JSON.parse(e.postData.contents).events[0].message.text;
  let userId =  JSON.parse(e.postData.contents).events[0].source.userId;

  if (typeof replyToken === 'undefined') {
    return;
  }

  let replyMessage = '';
  
  if (userMessage === '你好') {
    replyMessage = '您好，請問有什麼我可以幫助你的？' + userId;
  } else if (userMessage === '今天天氣如何') {
    replyMessage = '我無法提供天氣資訊，請查詢最近的天氣預報。';
  } else if (userMessage === 'todo') {
    replyMessage = 'https://script.google.com/macros/s/AKfycbzjeFRa2Y6Y--hFZ04U8QgC8b-tSWhF5HGmSptfdGkyT7Lg0jLyAzX81DOAL8SCj6gX/exec';
  }else {
    replyMessage = '對不起，我無法理解你的問題。';
  }

  let message = {
    'replyToken': replyToken,
    'messages': [{
      'type': 'text',
      'text': replyMessage
    }]
  };

  
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify(message),
  });
  return ContentService.createTextOutput(JSON.stringify({'content': 'message sent'})).setMimeType(ContentService.MimeType.JSON);
}

// 處理LINE訊息API的GET請求
function doGet() {
  //重新設定觸發器
  deleteTriggers();
  setUpTrigger();

  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('To-do List');
  let range = sheet.getDataRange();
  let values = range.getValues();
  let ignoreColumns = ["StartTime", "EndTime","Location","StartTimeFlag","EndTimeFlag","SetCalendar"];
  let headers = values[0].filter(function(header) {
    return !ignoreColumns.includes(header);
  });
  
  let todoList = [];
  for (let i = 1; i < values.length; i++) {
    let row = values[i];
    let todoItem = {};
    for (let j = 0; j < row.length; j++) {
      if (!ignoreColumns.includes(values[0][j])) {
        let cell = row[j];
        if (cell instanceof Date) {
          cell = Utilities.formatDate(cell, "GMT+8", "yyyy/MM/dd");
        }
        todoItem[values[0][j]] = cell;
      }
    }
    todoList.push(todoItem);
  }
  
  let template = HtmlService.createTemplateFromFile('index');
  template.headers = headers;
  template.todoList = todoList;
  return template.evaluate();
}