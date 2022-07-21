function getIsDuplicate(arr1, arr2) {
  return [...arr1, ...arr2].filter(item => arr1.includes(item) && arr2.includes(item)).length > 0
}

function getjuuhuku(array){
  var repetitiveArray = array.filter((value, index) => {return array.indexOf(value) !== index});
  Logger.log(repetitiveArray)
}


function myFunctio() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = spreadsheet.getSheetByName('フォームの回答').getDataRange().getValues();
  var sheet2 = spreadsheet.getSheetByName('スクリプト処理用コピー').getDataRange().getValues();
  Logger.log(sheet1.length)
  Logger.log(sheet2.length)
  // if(sheet1.length != sheet2.length){
  //   Logger.log('数値が一致していません')
  //   return 0;
  // }
  ans = getIsDuplicate(sheet1, sheet2)
  if(ans){
    Logger.log(ans)
  }
  name = []
  for(i=0;i<sheet2.length;i++){
    name.push(sheet2[i][3])
  }
  getjuuhuku(name)
  

}