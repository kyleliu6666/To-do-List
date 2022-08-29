function editForm(e){
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let thisRow = e.range.getRow();


  // 自動填入倒數日期
  if (ss.getRange(date + thisRow).getValues() !='' && ss.getRange(lastDay+1).getValues()[0][0] === 'Last Day'){
    ss.getRange(lastDay + thisRow).setFormula(`=${date}${thisRow}-today()`);
    sort();
  }
  else{
    ss.getRange(lastDay + thisRow).setValue('');
  }
  console.log(range);
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