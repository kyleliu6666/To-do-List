function getRow99(){
   console.log('test1');
}

function data_num(){
  let data = readData();
  console.log(data.length);
}




function test(){
  let message = "聽話!!!讓我看看";
  sendLineNotify(message, testToken);
}


function test1(){
  var members = [
      {name: 'Mike', age: 20},
      {name: 'Jimmy', age: 40},
      {name: 'Judy', age: 30}
  ];

  members.sort(function(a, b) {
      // boolean false == 0; true == 1
      return a.age > b.age ? 1 : -1;
  });

  // 順序依序會是 Mike -> Jimmy -> Judy
  console.log(members);
}


function test2(){
 let ss = SpreadsheetApp.getActiveSheet();
 console.log(ss.getRange(lastDay+1).getValues()[0][0]);
}



