function doGet(e) {

  let page = e.parameter.page || "Dashboard";
  let html = HtmlService.createTemplateFromFile(page).evaluate();
  let htmlOutput = HtmlService.createHtmlOutput(html);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1').setTitle('บันทึกข้อมูลคอมพิวเตอร์').setFaviconUrl('https://img2.pic.in.th/pic/272978196_302253305270068_6802695808492940439_n.th.png')
  htmlOutput.setContent(htmlOutput.getContent().replace("<Navbar/>",getNavbar(page)));
  return htmlOutput;

}

function globalVariables(){ 
  var varArray = {
    sheetNameCOM3   : 'COM3',
    sheetNameCOM7   : 'COM7',
    sheetNameCOMLIB : 'COMLIB',
    sheetNameNUMROOM : 'NUMROOM',
    sheetNameADMIN : 'ADMIN',
    firstColumn : '!A',
    lastColumn : 'M',
    dataRange : "!A2:M",
    spreadsheetId   : '',
  };
  return varArray;
}

function getDropdownOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameADMIN);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();  
  var result =  data.map(function(item) { return item[0]; });
  return result;
}

function getDropdownOptions3() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameNUMROOM);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();  
  var result =  data.map(function(item) { return item[0]; });
  return result;
}
function getDropdownOptions7() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameNUMROOM);
  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1).getValues();  
  var result =  data.map(function(item) { return item[0]; });
  return result;
}
function getDropdownOptionsLib() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameNUMROOM);
  var data = sheet.getRange(2, 3, sheet.getLastRow() - 1).getValues();  
  var result =  data.map(function(item) { return item[0]; });
  return result;
}

function getNavbar(activePage) {
  var scriptURLHome = getScriptURL("?page=Dashboard");
  var scriptURLPageCom3 = getScriptURL("?page=PageCom3");
  var scriptURLPageCom7 = getScriptURL("?page=PageCom7");
  var scriptURLPageComLib = getScriptURL("?page=PageComLib");

  var navbar = `
  <div id="layoutSidenav_nav">
    <nav class="sb-sidenav accordion sb-sidenav-dark" id="sidenavAccordion">
        <div class="sb-sidenav-menu">
          <div class="nav">
              <div class="sb-sidenav-menu-heading">ระบบบันทึกข้อมูลคอมพิวเตอร์</div>
              <a class="nav-link collapsed" style="font-size:14px" href="#" data-bs-toggle="collapse" data-bs-target="#collapseLayouts" aria-expanded="false" aria-controls="collapseLayouts">
                  <div class="sb-nav-link-icon"><i class="fa-solid fa-computer"></i></div>
                  บันทึกข้อมูลคอมพิวเตอร์
                  <div class="sb-sidenav-collapse-arrow"><i class="fas fa-angle-down"></i></div>
              </a>
              <div class="collapse" id="collapseLayouts" aria-labelledby="headingOne" data-bs-parent="#sidenavAccordion">
                  <nav class="sb-sidenav-menu-nested nav" style="font-size: 16px;">
                    <a class="nav-link ${activePage === 'Dashboard' ? 'active' : ''}" href="${scriptURLHome}"><i class="fa-solid fa-chart-line"></i>&nbsp;หน้าแดชบอร์ด</a>
                    <a class="nav-link ${activePage === 'PageCom3' ? 'active' : ''}" href="${scriptURLPageCom3}"><i class="fa-solid fa-desktop"></i>&nbsp;คอม ตึก3</a>
                    <a class="nav-link ${activePage === 'PageCom7' ? 'active' : ''}" href="${scriptURLPageCom7}"><i class="fa-solid fa-desktop"></i>&nbsp;คอม ตึก7</a>
                    <a class="nav-link ${activePage === 'PageComLib' ? 'active' : ''}" href="${scriptURLPageComLib}"><i class="fa-solid fa-desktop"></i>&nbsp;คอม ห้องสมุด</a>
                  </nav>
              </div>
              <div class="sb-sidenav-menu-heading">ระบบใจดีให้ยืม</div>
              <a class="nav-link collapsed" style="font-size:14px" href="#" data-bs-toggle="collapse" data-bs-target="#collapseLayouts2" aria-expanded="false" aria-controls="collapseLayouts">
                <div class="sb-nav-link-icon"><i class="fa-solid fa-book-open"></i></div>
                ใจดีให้ยืม
                <div class="sb-sidenav-collapse-arrow"><i class="fas fa-angle-down"></i></div>
              </a>
              <div class="collapse"  style="font-size: 16px;" id="collapseLayouts2" aria-labelledby="headingOne" data-bs-parent="#sidenavAccordion">
                <nav class="sb-sidenav-menu-nested nav">
                    <a class="nav-link"><i class="fa-solid fa-file-pen"></i>&nbsp;ยืม-คืนพัสดุ</a>
                </nav>
              </div>
              <div class="sb-sidenav-menu-heading">ระบบลงชื่อเข้าใช้งานห้องสมุด</div>
              <a class="nav-link collapsed" style="font-size:14px" href="#" data-bs-toggle="collapse" data-bs-target="#collapseLayouts3" aria-expanded="false" aria-controls="collapseLayouts">
                <div class="sb-nav-link-icon"><i class="fa-solid fa-book"></i></div>
                ลงชื่อเข้าใช้งานห้องสมุด
                <div class="sb-sidenav-collapse-arrow"><i class="fas fa-angle-down"></i></div>
              </a>
              <div class="collapse"  style="font-size: 16px;" id="collapseLayouts3" aria-labelledby="headingOne" data-bs-parent="#sidenavAccordion">
                <nav class="sb-sidenav-menu-nested nav">
                    <a href="#" class="nav-link"><i class="fa-solid fa-file-pen"></i>&nbsp;ลงชื่อเข้าใช้ห้องสมุด</a>
                </nav>
              </div>
          </div>
        </div>
    </nav>
  </div>`;

  return navbar;
}

function getScriptURL(qs) {
  var url = ScriptApp.getService().getUrl();
  if(qs){
    if(qs.indexOf("?") === -1) {
      qs = "?" + qs;
    }
    url = url + qs;
  }
  return url;
}

function getData(sheetName) {
  var spreadSheetId = globalVariables().spreadsheetId; 
  var dataRange = sheetName + globalVariables().dataRange; 
  var range = Sheets.Spreadsheets.Values.get(spreadSheetId, dataRange);
  var values = range.values;
  return values;
}

function generateUniqueId() {
  let id = Utilities.getUuid();
  return id;
}

function checkID(ID,sheetNameEdit){
  var idList = readData(globalVariables().spreadsheetId,sheetNameEdit + '!A2:A').reduce(function(a,b){return a.concat(b);});
  return idList.includes(ID);
}

function processForm(formObject) {
  let timestamp = new Date();
  var sheet = SpreadsheetApp.openById(globalVariables().spreadsheetId);
  var worksheet = sheet.getSheetByName(formObject.sheetName.toString());

  worksheet.appendRow([
    generateUniqueId(),
    formObject.comcodeInsert,
    formObject.cpuInsert,
    formObject.ramInsert,
    formObject.screenInsert,
    formObject.comdetailInsert,
    formObject.bdInsert,
    formObject.roomInsert,
    formObject.programsInsert.toString(),
    formObject.programsDetailInsert,
    formObject.statusInsert,
    formObject.adminInsert,
    timestamp,
  ])
}

function processEditForm(formObject){
  updateData(getFormValues(formObject),globalVariables().spreadsheetId,getRangeByID(formObject.recIdEdit,formObject.sheetNameEdit));
}

function updateData(values,spreadsheetId,range){
  var valueRange = Sheets.newValueRange();
  valueRange.values = values;
  Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, { valueInputOption: "RAW"});
}

function getFormValues(formObject){
  if(formObject.recIdEdit && checkID(formObject.recIdEdit,formObject.sheetNameEdit)){
    let timestamp = new Date();
      var values = [[
        formObject.recIdEdit.toString(),
        formObject.comcodeEdit,
        formObject.cpuEdit,
        formObject.ramEdit,
        formObject.screenEdit,
        formObject.comdetailEdit,
        formObject.bdEdit,
        formObject.roomEdit,
        formObject.programsEdit.toString(),
        formObject.programsDetailEdit,
        formObject.statusEdit,
        formObject.adminEdit,
        timestamp,
      ]]
    return values;
  }
}

function getRangeByID(id,sheetNameEdit){
  if(id){
    var idList = readData(globalVariables().spreadsheetId,sheetNameEdit + '!A2:A');
    for(var i=0;i<idList.length;i++){
      if(id==idList[i][0]){
        return sheetNameEdit + globalVariables().firstColumn + (i+2) +':'+ globalVariables().lastColumn +(i+2);
      }
    }
  }
}

function getRecordById(id,sheetNameEdit){
  if(id && checkID(id,sheetNameEdit)){
    var result = readData(globalVariables().spreadsheetId,getRangeByID(id,sheetNameEdit));
    return result;
  }
}

function readData(spreadsheetId,range){
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

function deleteRecord(id,sheetNameDelete) {

  const rowToDelete = getRowIndexById(id,sheetNameDelete);
  switch(sheetNameDelete){
    case 'COM3':
      var indexSheet = '0';
      break;
    case 'COM7':
      var indexSheet = '1691346315';
      break;
    case 'COMLIB':
      var indexSheet = '1480557122';
      break;
  } 

  // ข้อมูล	Sheet Name: COM3 | ID: 0
  // ข้อมูล	Sheet Name: COM7 | ID: 1691346315
  // ข้อมูล	Sheet Name: COMLIB | ID: 1480557122
  // ข้อมูล	Sheet Name: ADMIN | ID: 1653254103
  // ข้อมูล	Sheet Name: NUMROOM | ID: 1777573168

  let deleteRequest = {
    "deleteDimension": {
      "range": {
        "sheetId": indexSheet,
        "dimension": "ROWS",
        "startIndex": rowToDelete,
        "endIndex": rowToDelete + 1
      }
    }
  };
  
  Sheets.Spreadsheets.batchUpdate({'requests': [deleteRequest]}, globalVariables().spreadsheetId);
  return getLastTenRecords(sheetNameDelete);
}

function getRowIndexById(id,sheetName) {
  if (!id) {
    throw new Error('Invalid ID');
  }

  const idList = readRecord(sheetName + "!A2:A");
  for (var i = 0; i < idList.length; i++) {
    if (id == idList[i][0]) {
      var rowIndex = parseInt(i + 1);
      return rowIndex;
    }
  }
}

function getLastTenRecords(sheetName) {
  let lastRow = readRecord(sheetName + globalVariables().dataRange).length + 1;
  let startRow = lastRow - 9;
  if (startRow < 2) {
    startRow = 2;
  }
  let range = sheetName + globalVariables().firstColumn + startRow + ":" + globalVariables().lastColumn + lastRow;
  let lastTenRecords = readRecord(range);
  return lastTenRecords;
}

function readRecord(range) {
  try {
    let result = Sheets.Spreadsheets.Values.get(globalVariables().spreadsheetId, range);
    return result.values;
  } catch (err) {
    console.log('Failed with error %s', err.message);
  }
}

function getDataById(id) {

  var sheet = SpreadsheetApp.openById(globalVariables().spreadsheetId).getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == id) { 
      return data[i]; 
    }
  }
  return null; 
}

function countTotalRow() {

  var counterRow = [];
  var counterRoom = [];
  var sheetCOM3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOM3);
  var sheetNUMROOM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameNUMROOM);
  
  if (sheetCOM3 && sheetNUMROOM) {
    var lastRow = sheetCOM3.getLastRow();
    var rowCountCOM3 = lastRow - 1; 
    var range = sheetNUMROOM.getRange("A2:A");
    var values = range.getValues();
  
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] !== "") {
        count++;
      }
    }
    counterRoom.push(count);
    counterRow.push(rowCountCOM3)
  } 

  var sheetCOM7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOM7);
  if (sheetCOM7 && sheetNUMROOM) {
    var lastRow = sheetCOM7.getLastRow();
    var rowCountCOM7 = lastRow - 1;
    var range = sheetNUMROOM.getRange("B2:B");
    var values = range.getValues();
  
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] !== "") {
        count++;
      }
    }
    counterRoom.push(count);
    counterRow.push(rowCountCOM7)
  } 

  var sheetCOMLIB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOMLIB);
  if (sheetCOMLIB && sheetNUMROOM) {
    var lastRow = sheetCOMLIB.getLastRow();
    var rowCountCOMLIB = lastRow - 1; 
    var range = sheetNUMROOM.getRange("C2:C");
    var values = range.getValues();
  
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] !== "") {
        count++;
      }
    }
    counterRoom.push(count);
    counterRow.push(rowCountCOMLIB)
  } 
  var result = [...counterRow,...counterRoom]
  return result;
}

function getRoomCOM3(){
  let roomValues = readRecord(globalVariables().sheetNameNUMROOM + '!A2:A');
  let comValues = readRecord(globalVariables().sheetNameCOM3 + '!H2:H');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOM3);
  const lastRow = sheet.getLastRow();

  var roomList = []

  roomValues.forEach(e => {
    roomList.push(e.toString())
  })
  let sheetCOM3Data = {
    room: [],
    comCounter: [],
    comStatusCounter: [],
  }

  roomValues.forEach((e) => {
    sheetCOM3Data.room.push(e.toString());
  })

  var targetComValues = roomList;  
  var comCounterValue = []; 

  for (let i = 0; i < comValues.length; i++) {
    for (let j = 0; j < comValues[i].length; j++) {
      if (targetComValues.includes(comValues[i][j].toString())) {
        comCounterValue.push(comValues[i][j].toString());
      }
    }
  }
  var targetValuesCount = roomList;

  let counts = targetValuesCount.map(value => {
    return comCounterValue.flat().filter(item => item === value).length;
  });

  counts.forEach((item) => {
    sheetCOM3Data.comCounter.push(item.toString())
  })
  // เลือกข้อมูลจากคอลัมน์ H (คอลัมน์ที่ 8) และ K (คอลัมน์ที่ 11)
  const rangeH = sheet.getRange(2, 8, lastRow - 1, 1); // เลือกคอลัมน์ H (8)
  const rangeK = sheet.getRange(2, 11, lastRow - 1, 1); // เลือกคอลัมน์ K (11)
  const valuesH = rangeH.getValues(); // ข้อมูลจากคอลัมน์ H
  const valuesK = rangeK.getValues(); // ข้อมูลจากคอลัมน์ K
  
  // อาเรย์ของค่าที่ต้องการนับในคอลัมน์ H
  const valuesToCount = roomList // ระบุค่าที่ต้องการนับ
  
  const countMap = {};
  
  // ตั้งค่าเริ่มต้นสำหรับแต่ละค่าใน valuesToCount เป็น 0
  valuesToCount.forEach(value => {
    countMap[value] = 0;
  });
  
  // วนลูปข้อมูลเพื่อตรวจสอบค่า "พร้อมใช้งาน" ในคอลัมน์ K และนับค่าในคอลัมน์ H
  for (let i = 0; i < valuesH.length; i++) {
    const valueInH = valuesH[i][0]; // ข้อมูลในคอลัมน์ H
    const valueInK = valuesK[i][0]; // ข้อมูลในคอลัมน์ K
    
    if (valueInK === "พร้อมใช้งาน" && valuesToCount.includes(valueInH.toString())) {
      countMap[valueInH]++; // เพิ่มค่าเมื่อพบค่าใน valuesToCount
    }
  }
  
  // แปลง countMap เป็นอาเรย์ของจำนวน (count) เท่านั้น
  const countArray = Object.values(countMap);

  countArray.forEach(item => {
    sheetCOM3Data.comStatusCounter.push(item.toString())
  })
  
  return sheetCOM3Data;
  
}
function getRoomCOM7(){
  let roomValues = readRecord(globalVariables().sheetNameNUMROOM + '!B2:B');
  let comValues = readRecord(globalVariables().sheetNameCOM7 + '!H2:H');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOM7);
  const lastRow = sheet.getLastRow();

  var roomList = []

  roomValues.forEach(e => {
    roomList.push(e.toString())
  })
  let sheetCOM7Data = {
    room: [],
    comCounter: [],
    comStatusCounter: [],
  }

  roomValues.forEach((e) => {
    sheetCOM7Data.room.push(e.toString());
  })

  var targetComValues = roomList;  
  var comCounterValue = []; 

  for (let i = 0; i < comValues.length; i++) {
    for (let j = 0; j < comValues[i].length; j++) {
      if (targetComValues.includes(comValues[i][j].toString())) {
        comCounterValue.push(comValues[i][j].toString());
      }
    }
  }
  var targetValuesCount = roomList;

  let counts = targetValuesCount.map(value => {
    return comCounterValue.flat().filter(item => item === value).length;
  });

  counts.forEach((item) => {
    sheetCOM7Data.comCounter.push(item.toString())
  })
  // เลือกข้อมูลจากคอลัมน์ H (คอลัมน์ที่ 8) และ K (คอลัมน์ที่ 11)
  const rangeH = sheet.getRange(2, 8, lastRow - 1, 1); // เลือกคอลัมน์ H (8)
  const rangeK = sheet.getRange(2, 11, lastRow - 1, 1); // เลือกคอลัมน์ K (11)
  const valuesH = rangeH.getValues(); // ข้อมูลจากคอลัมน์ H
  const valuesK = rangeK.getValues(); // ข้อมูลจากคอลัมน์ K
  
  // อาเรย์ของค่าที่ต้องการนับในคอลัมน์ H
  const valuesToCount = roomList // ระบุค่าที่ต้องการนับ
  
  const countMap = {};
  
  // ตั้งค่าเริ่มต้นสำหรับแต่ละค่าใน valuesToCount เป็น 0
  valuesToCount.forEach(value => {
    countMap[value] = 0;
  });
  
  // วนลูปข้อมูลเพื่อตรวจสอบค่า "พร้อมใช้งาน" ในคอลัมน์ K และนับค่าในคอลัมน์ H
  for (let i = 0; i < valuesH.length; i++) {
    const valueInH = valuesH[i][0]; // ข้อมูลในคอลัมน์ H
    const valueInK = valuesK[i][0]; // ข้อมูลในคอลัมน์ K
    
    if (valueInK === "พร้อมใช้งาน" && valuesToCount.includes(valueInH.toString())) {
      countMap[valueInH]++; // เพิ่มค่าเมื่อพบค่าใน valuesToCount
    }
  }
  
  // แปลง countMap เป็นอาเรย์ของจำนวน (count) เท่านั้น
  const countArray = Object.values(countMap);

  countArray.forEach(item => {
    sheetCOM7Data.comStatusCounter.push(item.toString())
  })

  return sheetCOM7Data;
  
}
function getRoomCOMLIB(){
  let roomValues = readRecord(globalVariables().sheetNameNUMROOM + '!C2:C');
  let comValues = readRecord(globalVariables().sheetNameCOMLIB + '!H2:H');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(globalVariables().sheetNameCOMLIB);
  const lastRow = sheet.getLastRow();

  var roomList = []

  roomValues.forEach(e => {
    roomList.push(e.toString())
  })
  let sheetCOMLIBData = {
    room: [],
    comCounter: [],
    comStatusCounter: [],
  }

  roomValues.forEach((e) => {
    sheetCOMLIBData.room.push(e.toString());
  })

  var targetComValues = roomList;  
  var comCounterValue = []; 

  for (let i = 0; i < comValues.length; i++) {
    for (let j = 0; j < comValues[i].length; j++) {
      if (targetComValues.includes(comValues[i][j].toString())) {
        comCounterValue.push(comValues[i][j].toString());
      }
    }
  }
  var targetValuesCount = roomList;

  let counts = targetValuesCount.map(value => {
    return comCounterValue.flat().filter(item => item === value).length;
  });

  counts.forEach((item) => {
    sheetCOMLIBData.comCounter.push(item.toString())
  })
  // เลือกข้อมูลจากคอลัมน์ H (คอลัมน์ที่ 8) และ K (คอลัมน์ที่ 11)
  const rangeH = sheet.getRange(2, 8, lastRow - 1, 1); // เลือกคอลัมน์ H (8)
  const rangeK = sheet.getRange(2, 11, lastRow - 1, 1); // เลือกคอลัมน์ K (11)
  const valuesH = rangeH.getValues(); // ข้อมูลจากคอลัมน์ H
  const valuesK = rangeK.getValues(); // ข้อมูลจากคอลัมน์ K
  
  // อาเรย์ของค่าที่ต้องการนับในคอลัมน์ H
  const valuesToCount = roomList // ระบุค่าที่ต้องการนับ
  
  const countMap = {};
  
  // ตั้งค่าเริ่มต้นสำหรับแต่ละค่าใน valuesToCount เป็น 0
  valuesToCount.forEach(value => {
    countMap[value] = 0;
  });
  
  // วนลูปข้อมูลเพื่อตรวจสอบค่า "พร้อมใช้งาน" ในคอลัมน์ K และนับค่าในคอลัมน์ H
  for (let i = 0; i < valuesH.length; i++) {
    const valueInH = valuesH[i][0]; // ข้อมูลในคอลัมน์ H
    const valueInK = valuesK[i][0]; // ข้อมูลในคอลัมน์ K
    
    if (valueInK === "พร้อมใช้งาน" && valuesToCount.includes(valueInH.toString())) {
      countMap[valueInH]++; // เพิ่มค่าเมื่อพบค่าใน valuesToCount
    }
  }
  
  // แปลง countMap เป็นอาเรย์ของจำนวน (count) เท่านั้น
  const countArray = Object.values(countMap);

  countArray.forEach(item => {
    sheetCOMLIBData.comStatusCounter.push(item.toString())
  })

  return sheetCOMLIBData;
  
}

// เรียกใช้ไฟล์ html ตัวอย่าง: <?!= include('CSS.html'); ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}