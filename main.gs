// 安裝觸發器
function setUpTrigger() {
  let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ScriptApp.newTrigger('editForm')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
}

// 刪除觸發器
function deleteTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "editForm") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// 自動排序
function sort(){
  let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sourceSheet = ss.getSheetByName("To-do List");
  let rangeSS = sourceSheet.getRange(2,1,sourceSheet.getLastRow()-1,sourceSheet.getLastColumn());
  rangeSS.sort([{column:1,ascending:true}]);
}

// 觸發表單編輯事件
function editForm(e){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let sheetName = ss.getName();
  
  // 檢查是否編輯的是 "To-do List" 表單
  if(sheetName !== "To-do List") {
    return;
  }

  let thisRow = e.range.getRow();
  let thisCol = e.range.getColumn();
  let cal = CalendarApp.getCalendarById(calendar_id);
  let title = ss.getRange(mission + thisRow).getValues()[0][0];
  let dateFlag = ss.getRange(dateCol + thisRow).getValues()[0][0];
  let startTime = ss.getRange(startTimeCol + thisRow).getValues()[0][0];
  let startTimeFlag = ss.getRange(beginCol + thisRow).getValues()[0][0];
  let endTimeFlag = ss.getRange(endCol + thisRow).getValues()[0][0];
  let desc = ss.getRange(descriptionCol + thisRow).getValues()[0][0];
  let loc = ss.getRange(locationCol + thisRow).getValues()[0][0];
  let calendar_flag = ss.getRange(calendarCol + thisRow).getValues()[0][0];
  
  // 非編輯第一列時觸發
  if(thisRow != 1){ 

    // 自動填入倒數日期
    if (ss.getRange(dateCol + thisRow).getValues() != ''){
      ss.getRange(lastDayCol + thisRow).setFormula(`=${dateCol}${thisRow}-today()`);
      // 預設填入N，不加入行事曆
      if(ss.getRange(calendarCol + thisRow).getValues() == ''){
        ss.getRange(calendarCol + thisRow).setValue('N');
      }
      const caseStatusList = ['N','Y'];
      // dropdown init
      let defaultOption = SpreadsheetApp.newDataValidation().requireValueInList(caseStatusList).build();
      ss.getRange(calendarCol + thisRow).setDataValidation(defaultOption);  
      ss.getRange(weekday + thisRow).setFormula(`=CHOOSE(WEEKDAY(${dateCol}${thisRow},2),"一","二","三","四","五","六","日")`);
      sort();  // 再排序
    }

    // 有填入C欄StartTime，插入I欄startTimeCol日期時間固定格式
    if(startTimeFlag != ''){
      ss.getRange(startTimeCol + thisRow).setFormula(`=${dateCol}${thisRow}+${beginCol}${thisRow}`);
    }

    // 有填入D欄EndTime，插入J欄endTimeCol日期時間固定格式
    if(endTimeFlag != ''){
      ss.getRange(endTimeCol + thisRow).setFormula(`=${dateCol}${thisRow}+${endCol}${thisRow}`);
    }

    // 建立日曆事件
    if(calendar_flag === 'Y'){
      // 結束時間
      let endTime = ss.getRange(endTimeCol + thisRow).getValues()[0][0];
      
      if(startTimeFlag == '' && endTimeFlag == ''){
        console.log(dateFlag);
        // 建立全日事件
        cal.createAllDayEvent(title,dateFlag);
      }
      // 建立有時間的事件
      else{
        cal.createEvent(title, startTime, endTime, {description : desc, location : loc});
      }
    }

  }
}