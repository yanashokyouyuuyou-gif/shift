function doPost(e){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Responses');
  const data=JSON.parse(e.postData.contents);

  const values=sheet.getDataRange().getValues();

  for(let i=1;i<values.length;i++){
    if(values[i][0]===data.month && values[i][1]===data.name){
      sheet.getRange(i+1,1,1,4).setValues([[data.month,data.name,new Date(),JSON.stringify(data.slots)]]);
      return ContentService.createTextOutput('UPDATED');
    }
  }

  sheet.appendRow([data.month,data.name,new Date(),JSON.stringify(data.slots)]);
  return ContentService.createTextOutput('OK');
}
