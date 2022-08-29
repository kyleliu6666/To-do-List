//自動排序
function sort(){
  let ss = SpreadsheetApp.openById(sheetID);
  let sourceSheet = ss.getSheetByName("To-do List");
  let rangeSS = sourceSheet.getRange(2,1,sourceSheet.getLastRow()-1,sourceSheet.getLastColumn());
  rangeSS.sort([{column:1,ascending:true}]);
}

//推播每日Todo
function todoPush() {
  let ss = SpreadsheetApp.openById(sheetID);
  ss = ss.getSheets()[0];
  let lastRow = ss.getLastRow(); //目前最後一行
  let alertDay = 10; //倒數提醒日
  let value = ss.getRange("D2:D" + lastRow).getValues();
  let valueLength = value.length; //資料長度
  let message = '';
  
  let dataArray = []
    for(let i=0;i<valueLength;i++){
      dataArray.push(value[i][0]);
  }

  dataArray.forEach(function(data,index) {
    if(data <= alertDay & data >= 0){
      dayend = ss.getRange(date + (index + 2)).getValues()[0][0];
      dayend = Utilities.formatDate(dayend, 'Asia/Taipei', 'MM/dd/yyyy');
      detail = ss.getRange(mission + (index + 2)).getValues()[0][0];
      message += ` \n${dayend}\n${detail}\n倒數 ${data}天\n`;
      console.log(dayend,detail,data);
    }

      
  });
  message = message.replace(/([\s]*$)/g, ""); //移除最後一個空白字元

  //有資料才POST到LINE
  if(message != ''){
    sendLineNotify(message, testToken);
  }
  
}

//過期案件移至complete List sheet
function moveList() {
  let ss = SpreadsheetApp.openById(sheetID);
  let sourceSheet = ss.getSheetByName("To-do List");
  let sourceRange = sourceSheet.getDataRange();
  let sourceValues = sourceRange.getValues();
  sourceValues.splice(0,1);
  let alertDay = 10; //倒數提醒日

  if (sourceValues[0] != null){
     let rowCount = sourceValues.length;
     let columnCount = sourceValues[0].length;
     
     //過期件
     let updateList = [];
     for(let i=0; i < rowCount; i++){
       if(sourceValues[i][3] < 0){
         updateList.push([sourceValues[i][0],'',sourceValues[i][2],sourceValues[i][3],'',''])
       }
    }

    //再排序
    updateList.sort(function(a, b) {
        // boolean false == 0; true == 1
        return a[3] > b[3] ? 1 : -1;
    });
 
    let targetSheet = ss.getSheetByName("complete List");
    let lastRow = targetSheet.getLastRow();
    let targetRange = targetSheet.getRange(lastRow+1, 1, updateList.length, columnCount);

    targetRange.setValues(updateList);
  }
  keepFit();
}

//清除過期案件
function keepFit(){
  let ss = SpreadsheetApp.openById(sheetID);
  let targetSheet = ss.getSheetByName("To-do List");
  let values = targetSheet.getRange("D2:D").getValues();
  let row_del = new Array();

  for(let i=0;i<values.length;i++)
  {
    if(values[i]!= '' && values[i] < 0){
      row_del.push(i+2);
    }
  }

  for (let i = row_del.length - 1; i >= 0; i--) {
  targetSheet.deleteRow(row_del[i]); 
  }
}



function sendLineNotify(message, token){
  let options =
   {
     "method"  : "post",
     "payload" : {"message" : message},
     "headers" : {"Authorization" : "Bearer " + token}
   };
   UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}
