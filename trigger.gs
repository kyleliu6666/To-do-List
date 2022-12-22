function editForm(e){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let thisRow = e.range.getRow();
  let thisCol = e.range.getColumn();

  if(thisRow != 1){ //編輯第一列的時候不觸發
    //自動填入倒數日期
    //第一欄有填&&D欄=Last Day
    if (ss.getRange(date + thisRow).getValues() != '' && ss.getRange(lastDay+1).getValues()[0][0] === 'Last Day'){
      ss.getRange(lastDay + thisRow).setFormula(`=${date}${thisRow}-today()`);
      ss.getRange(weekday + thisRow).setFormula(`=CHOOSE(WEEKDAY(${date}${thisRow},2),"一","二","三","四","五","六","日")`);
      sort();
    }
    else{
      ss.getRange(lastDay + thisRow).setValue('');
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