function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menubuttons = [ {name: "ย้ายข้อมูลไปยังระบบและส่งอีเมลแจ้ง", functionName: "moveToSystem"}, {name: "ลบข้อมูลเสื้อผ้าทั้งหมด", functionName: "clearRange1"},
                       {name: "ลบเฉพาะราคาทั้งหมด", functionName: "clearRange2"}, {name: "ลบข้อมูลทั้งหมด", functionName: "clearRange3"}];
    sheet.addMenu("@@เมนูจัดการสำหรับ Admin", menubuttons);
} 

//1. ย้ายข้อมูลไประบบพร้อมส่งอีเมล

function moveToSystem() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('โปรดยืนยัน?', 'ต้องการย้ายข้อมูลไป Report และส่งอีเมลหรือไม่?', ui.ButtonSet.YES_NO);
  var sheet = SpreadsheetApp.getActive().getSheetByName('home');
  var srep = SpreadsheetApp.getActive().getSheetByName('srepaccounting');
  var email = SpreadsheetApp.getActive().getSheetByName('emailservices');
  
  if (result == ui.Button.YES) {
//workspace   
  var date = sheet.getRange('B3');
  var source = sheet.getRange('K19');
  var profit = sheet.getRange('L19');

  var destRowColumnRange = srep.getLastColumn() && srep.getLastRow();

  var destDate = srep.getRange(destRowColumnRange+1, 1);
  var destSource = srep.getRange(destRowColumnRange+1, 2);
  var destProfit = srep.getRange(destRowColumnRange+1, 3);

  date.copyTo (destDate, {contentsOnly: true});  
  source.copyTo (destSource, {contentsOnly: true});
  profit.copyTo (destProfit, {contentsOnly: true});

//email services
  var emailInfoRange = email.getRange("A2:G2");
  var dateEmail = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
  var lastRow = srep.getLastRow();
  var colcost = 2;
  var colprof = 3;
  var dataEmail = emailInfoRange.getValues();
  var lastRowCost = srep.getRange(lastRow, colcost).getValue();
  var lastRowProf = srep.getRange(lastRow, colprof).getValue();
//  var getRangeFunct = srep.getValue(destSource);
    
  for (i in dataEmail) {
    var rowData = dataEmail[i];
    var emailAddress = rowData[1];
    var recipient = rowData[0];
    var message1 = rowData[2];
    var message2 = rowData[3];
    var parameter2 = lastRowProf;
    var message3 = rowData[5];
    var message4 = rowData[6];
    var message = "สวัสดีคุณ " + recipient + ',\n\n' + message1 + ' ' + message2 + ' ' + parameter2 + ' ' + message3 + '\n\n' + message4;
    var subject = "(REP) ส่งยอดล่าสุด " + dateEmail;
    MailApp.sendEmail(emailAddress, subject, message);   
    ui.alert("ย้ายข้อมูลไปแล้วและส่งอีเมลแล้ว");
  }
 }
    else{  
    ui.alert("ไม่ได้ทำการย้ายข้อมูล");
  }
}

//2. ลบข้อมูลเสื้อผ้าทั้งหมดออก

function clearRange1() { //replace 'Sheet1' with your actual sheet name
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('โปรดยืนยัน?', 'ต้องการลบข้อมูลทั้งหมดหรือไม่', ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {

  var sheet = SpreadsheetApp.getActive().getSheetByName('home');
  sheet.getRange('C8:C17').setValue("_");
  sheet.getRange('D8:D17').setValue("_");
  sheet.getRange('E8:E17').setValue("_");
  sheet.getRange('F8:F17').setValue("_");
    
    ui.alert('ลบข้อมูลออกแล้ว.');
  } else {

    ui.alert('ไม่ได้ทำการลบข้อมูลออก.');
  }
}

//3. ลบข้อมูลเงินทั้งหมดออก

function clearRange2() { //replace 'Sheet1' with your actual sheet name
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('โปรดยืนยัน?', 'ต้องการลบข้อมูลจำนวนเงินทั้งหมดหรือไม่', ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    
  var sheet = SpreadsheetApp.getActive().getSheetByName('home');
  sheet.getRange('G8:G17').setValue(0);
    
    ui.alert('ลบข้อมูลออกแล้ว.');
  } else {

    ui.alert('ไม่ได้ทำการลบข้อมูลออก.');
  }
}

//4. ลบข้อมูลทั้งหมดออก

function clearRange3() { //replace 'Sheet1' with your actual sheet name
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('โปรดยืนยัน?', 'ต้องการลบข้อมูลทั้งหมดหรือไม่', ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('home');
  sheet.getRange('C8:C17').setValue("_");
  sheet.getRange('D8:D17').setValue("_");
  sheet.getRange('E8:E17').setValue("_");
  sheet.getRange('F8:F17').setValue("_");
  sheet.getRange('G8:G17').setValue(0);
    
    ui.alert('ลบข้อมูลออกแล้ว.');
  } else {

    ui.alert('ไม่ได้ทำการลบข้อมูลออก.');
  }
}

//5. Function change font color

function changeFontColor() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('home');
  var cell_target = sheet.getRange('K21')
  var cell_values = cell_target.getValues();
  var targetvalue = "ไม่มีข้อผิดพลาด";
  var targetvalue2 = "ข้อผิดพลาดที่พบ (40 / 20)";
  var red = 'red';
  var blue = 'blue';
  
  if(cell_values == targetvalue)
  {
    cell_target.setFontColor(blue);    
  }
  else
  {
    cell_target.setFontColor(red);    
  }  
  
}
