function doGet(){
  return HtmlService.createHtmlOutputFromFile('index').setTitle('ลงชื่อเข้าใช้งานห้องสมุด').setFaviconUrl('https://img2.pic.in.th/pic/272978196_302253305270068_6802695808492940439_n.th.png');
}

// เรียกใช้ไฟล์ CSS.html ตัวอย่าง: <?!= include('CSS'); ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// เช็ตข้อมูลชีตเริ่มต้น
function initialSpreadsheets(){
  const data = {
    spreadSheetID: '',
    sheetNameSTUDENT_ID: 'STUDENT_ID',
    sheetNameSTUIDRange: 'STUDENT_ID!A2:A',
    sheetNameLOG: 'LOG',
    sheetNameLOGRange: 'LOG!A2:C',
  }
  return data;
}

// ดึงข้อมูลจากชีต รูปแบบ 'sheetname!A2:B'
function readData(spreadsheetId,range){
  let result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

// เช็ตรูปแบบเวลา
function formatNewDate(){
  let _formatNewDate = Utilities.formatDate(new Date(),Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
  return _formatNewDate;
}

// เช็ตรูปแบบเวลาที่รับอินพุตเข้ามา
function formatInputDate(time){
  let _formatInputDate = Utilities.formatDate(new Date(time),Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
  return _formatInputDate;
}

// เช็ครหัสนักศึกษาที่มีในชีต STUDENT_ID
function getDataFilterStudentIdForCheck(spreadSheetID,sheetName,studentId){
  let stdentIdSheet = readData(spreadSheetID, sheetName);
  var getstuId = stdentIdSheet.filter(function(row) {
    return row[0] == studentId;
  });
  return getstuId;
}

// ดึงข้อมูลการเข้า - ออกห้องสมุด
function getStudentSigninHistory(studentId){
  let logSheetcheckStudentId = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameLOG);
  let lastRow = logSheetcheckStudentId.getLastRow(); // ดึงหมายเลขแถวสุดท้ายของชีต
  let getDataSheetLOG = logSheetcheckStudentId.getRange(2, 1, lastRow, logSheetcheckStudentId.getLastColumn()).getValues();

  // ใช้ filter เพื่อค้นหารหัสนักศึกษาที่ลงชื่อเข้าใช้งานห้องสมุด
  var result = getDataSheetLOG.filter(function(row) {
    return row[0] == studentId; 
  });

  var allEntries = result.map(function(row) {
    let dateSignin = formatInputDate(row[1]);
    let dateLogout = formatInputDate(row[2]);
    let hours = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "HH:mm:ss");
    if(row[2] == ''){
      dateLogout = "ไม่มีข้อมูล";
    }
    if(row[3] == ''){
      hours = "ไม่มีข้อมูล";
    }
    return row[0] + "," + dateSignin + "," + dateLogout + "," + hours;
  });
  return allEntries;
}

// คำนวนชั่วโมงจาก เวลาเข้า เวลาออก
function calculateLibraryUsage(studentId) {

  var sheet = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameLOG);
  var data = sheet.getDataRange().getValues();

  // Filter data by student ID
  var studentData = data.filter(function(row) {
    return row[0] == studentId; // Assuming student ID is in column A (index 0)
  });

  // คำนวณเวลาการใช้งานสำหรับแต่ละรอบ
  var results = studentData.map(function(row) {
    var checkInTime = new Date(row[1]); // Assuming check-in time is in column D (index 3)
    var checkOutTime = new Date(row[2]); // Assuming check-out time is in column E (index 4)

    // คำนวณความแตกต่างระหว่างเวลาเข้าและออก
    var differenceInMillis = checkOutTime - checkInTime;

    if (differenceInMillis < 0) {
      return {
        round: row[2], // Assuming round info is in column C (index 2)
        error: 'Check-out time is earlier than check-in time'
      };
    }
    var seconds = Math.floor((differenceInMillis / 1000) % 60);
    var minutes = Math.floor((differenceInMillis / (1000 * 60)) % 60);
    var hours = Math.floor(differenceInMillis / (1000 * 60 * 60));

    return `${hours}:${minutes}:${seconds}`;
  });
  return results;
}

// เช็คค่าว่างเวลาออกจากห้องสมุดของรหัสนักศึกษาแถวล่าสุด
function checkLastRowValueStudentLogoutEmpty(studentId) {
  var studentSignin = getDataFilterStudentIdForCheck(initialSpreadsheets().spreadSheetID,initialSpreadsheets().sheetNameLOG, studentId)
  try{
    if (studentSignin.length > 0) {
      var lastRowIndex = studentSignin.map(function(row) { return row[0]; }).lastIndexOf(studentId);
      var column3Value = studentSignin[lastRowIndex];
      if(column3Value[2] == null){
        return true;
      }else{
        return false;
      }
    }
  }catch{
    return false;
  }
}

// เพิ่มรหัสนักษาที่เคยใช้งานครั้งแรกลงในชีต STUDENT_ID
function addStudentIfNotExists(studentId) {
  var sheet = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameSTUDENT_ID);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    sheet.getRange(2, 1).setValue(studentId); // วางรหัสนักศึกษาใน A2
    sheet.getRange(2, 2).setValue(formatNewDate()); // วางวันเวลาที่ลงชื่อเข้าใน B2
    return 'Student added successfully';
  } else {
    // ดึงข้อมูลตั้งแต่ A2 ลงไป (ข้ามหัวตาราง)
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    // ตรวจสอบว่ารหัสนักศึกษามีอยู่แล้วหรือไม่
    var studentExists = data.some(function(row) {
      return row[0] == studentId; // Assuming student ID is in column A
    });

    // ถ้ายังไม่มี ให้วางข้อมูลใหม่
    if (!studentExists) {
      var newRow = lastRow + 1; // แถวถัดไปหลังจากแถวสุดท้าย
      sheet.getRange(newRow, 1).setValue(studentId); // วางรหัสนักศึกษาในคอลัมน์ A
      sheet.getRange(newRow, 2).setValue(formatNewDate()); // วางวันเวลาที่ลงชื่อเข้าในคอลัมน์ B
      return 'Student added successfully';
    } else {
      return 'Student already exists';
    }
  }
}

// ฟังก์ชันหลักที่ใช้กำหนดการทำงาน
function processForm(studentId,status){
  try{
    switch(status){
    case "signin":
      var dataReturnSignIn = appendDataSignIn(studentId)
      return dataReturnSignIn;
    case "logout":
      var dataReturnLogout = updateDataLogout(studentId)
      return dataReturnLogout;
  }
  }catch{
    return {
      studentData: studentId,
      studentLog: [],
      status: false,
    }
  }
}

// อัปเดตเวลาเข้าใช้งานห้องสมุด
function appendDataSignIn(studentId){
  try{
    let sheetStudentId = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameSTUDENT_ID);
    let sheetLogSignIn = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameLOG);
    if(checkLastRowValueStudentLogoutEmpty(studentId) != true){
      sheetLogSignIn.appendRow([studentId, formatNewDate()]);
      addStudentIfNotExists(studentId)
      var studentLog = getStudentSigninHistory(studentId)
      var getTime = studentLog[0].split(",")[1]
      var timeSignin = getTime.split(" ")[1]
      return {
        studentData: studentId,
        studentLog: studentLog,
        timeSignin: timeSignin,
        status:"signin",
      }
    }else{
      return {
        studentData: [studentId],
        studentLog: [],
        status: null,
      }
    }
  }catch{
    return {
      studentData: [studentId],
      studentLog: [],
      status: false,
    }
  }
}

// อัปเดตเวลาออกจากห้องสมุด
function updateDataLogout(studentId){
  try{
    let logSheetcheckStudentIdLogout = SpreadsheetApp.openById(initialSpreadsheets().spreadSheetID).getSheetByName(initialSpreadsheets().sheetNameLOG);
    let lastRow = logSheetcheckStudentIdLogout.getLastRow(); // ดึงหมายเลขแถวสุดท้ายของชีต
    let getStudentIdFromSheetLOG = logSheetcheckStudentIdLogout.getRange(2, 1, lastRow - 1, logSheetcheckStudentIdLogout.getLastColumn()).getValues();

    let latestRow = -1; // เริ่มต้นเป็น -1 เพื่อเช็คว่าพบรหัสนักศึกษาไหม
    for (let i = 0; i < getStudentIdFromSheetLOG.length; i++) {
      if (getStudentIdFromSheetLOG[i][0].toString() == studentId) {
        latestRow = i + 2; // เก็บตำแหน่งแถวล่าสุดที่พบรหัสนักศึกษา (บวก 2 เพื่อให้ตรงกับแถวในชีต)
      }
    }
    // ตรวจสอบว่าพบรหัสนักศึกษาหรือไม่
    if (latestRow !== -1) {
      if(checkLastRowValueStudentLogoutEmpty(studentId) == true){ // อัปเดตเวลาลงในคอลัมน์ C ถ้าค่าในคอลัมเท่ากับค่าว่าง
        logSheetcheckStudentIdLogout.getRange(latestRow, 3).setValue(formatNewDate());
        logSheetcheckStudentIdLogout.getRange(latestRow, 4).setValue(calculateLibraryUsage(studentId));
        var studentLog = getStudentSigninHistory(studentId)
        var getTime = studentLog[0].split(",")[2]
        var timeLogout = getTime.split(" ")[1]
        return {
          studentData: studentId,
          studentLog: studentLog,
          timeLogout: timeLogout,
          status:"logout",
        }
      }else{
        return {
          studentData: studentId,
          studentLog: [],
          status: false,
        }
      }
    }else{
      return {
        studentData: studentId,
        studentLog: [],
        status: false,
      }
    }
  }catch{
    return {
      studentData: studentId,
      studentLog: [],
      status: false,
    }
  }
}



