function editForm(e){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let thisRow = e.range.getRow();
  let thisCol = e.range.getColumn();
  let cal = CalendarApp.getCalendarById(calendar_id);
  let title = ss.getRange(mission + thisRow).getValues()[0][0];
  let startTime = ss.getRange(startTimeCol + thisRow).getValues()[0][0];
  let endTimeFlag = ss.getRange(time2 + thisRow).getValues()[0][0];
  let desc = ss.getRange(descriptionCol + thisRow).getValues()[0][0];
  let loc = ss.getRange(locationCol + thisRow).getValues()[0][0];
  let calendar_flag = ss.getRange(calendarCol + thisRow).getValues()[0][0];
  
  if(thisRow != 1){ //編輯第一列的時候不觸發

    if(endTimeFlag != ''){
      ss.getRange(endTimeCol + thisRow).setFormula(`=${date}${thisRow}+${time2}${thisRow}`);
      let endTime = ss.getRange(endTimeCol + thisRow).getValues()[0][0];

      if(calendar_flag ==='Y'){
      //插入日曆
        cal.createEvent(title, startTime, endTime, {description : desc, location : loc});
      }
    }

  
    //自動填入倒數日期
    //第一欄有填&&D欄=Last Day
    if (ss.getRange(date + thisRow).getValues() != '' && ss.getRange(lastDay+1).getValues()[0][0] === 'Last Day'){
      ss.getRange(lastDay + thisRow).setFormula(`=${date}${thisRow}-today()`);
      ss.getRange(weekday + thisRow).setFormula(`=CHOOSE(WEEKDAY(${date}${thisRow},2),"一","二","三","四","五","六","日")`);
      sort();  //再排序
    }
  }

}

function doGet() {
  //刪除觸發器
  deleteTriggers();
  //安裝觸發器
  setUpTrigger();
  return ContentService.createTextOutput('Triggers Setup is Complete');
}

//安裝觸發器
function setUpTrigger() {
  let ss = SpreadsheetApp.openById(sheetID);
  ScriptApp.newTrigger('editForm')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}

//刪除觸發器
function deleteTriggers() {
    let triggers = ScriptApp.getProjectTriggers();
     for (let i = 0; i < triggers.length; i++) {
         if (triggers[i].getHandlerFunction() == "editForm") {
          ScriptApp.deleteTrigger(triggers[i]);
      }
     }
}