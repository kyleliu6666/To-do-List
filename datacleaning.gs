// 過期案件移至complete List sheet
function moveList() {
    let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
        console.log(sourceValues[i][4]);
        if(sourceValues[i][4] < 0){ //lastDayCol
          updateList.push(
            [sourceValues[i][0],sourceValues[i][1],sourceValues[i][2],sourceValues[i][3],sourceValues[i][4],sourceValues[i][5],sourceValues[i][6],sourceValues[i][7],sourceValues[i][8],
            sourceValues[i][9],sourceValues[i][10]])
        }
      }
  
      //再排序
      updateList.sort(function(a, b) {
          // boolean false == 0; true == 1
          return a[5] > b[5] ? 1 : -1;
      });
  
      let targetSheet = ss.getSheetByName("complete List");
      let lastRow = targetSheet.getLastRow();
      let targetRange = targetSheet.getRange(lastRow+1, 1, updateList.length, columnCount);
  
      targetRange.setValues(updateList);
    }
    keepFit();
  }
  
  // 清除過期案件
  function keepFit(){
    let ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let targetSheet = ss.getSheetByName("To-do List");
    let values = targetSheet.getRange(lastDayCol+'2:'+lastDayCol).getValues();
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
  