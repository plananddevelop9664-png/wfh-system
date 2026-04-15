// กำหนด ID ของ Google Sheet
var SHEET_ID = '1QLigeftApNV9HBhsibeaET024Mlxwe4VIAl1bfgCJ-E'; // << อย่าลืมเปลี่ยนเป็น ID Sheet ของตัวเอง

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var result = { status: "error", message: "Invalid action" };

  try {
    if (action === 'login') {
      result = doLogin(e.parameter.username, e.parameter.password);
    }
    else if (action === 'submit') {
      var payload = JSON.parse(e.parameter.payload);
      result = submitRequest(payload);
    }
    else if (action === 'getData') {
      result = getCalendarData(e.parameter.employee_id);
    }
    else if (action === 'getHistory') {
      result = getHistory(e.parameter.employee_id);
    }
    else if (action === 'getApprovals') {
      result = getApprovals(e.parameter.manager_email, e.parameter.role);
    }
    else if (action === 'processApproval') {
      result = processApproval(e.parameter.request_id, e.parameter.status, e.parameter.manager_email, e.parameter.comment);
    }
    else if (action === 'editRequest') {
      var editPayload = JSON.parse(e.parameter.payload);
      result = editRequest(editPayload.request_id, editPayload.new_reason);
    }
    else if (action === 'deleteRequest') {
      result = deleteRequest(e.parameter.request_id, e.parameter.manager_email);
    }
    else if (action === 'getReport') {
      result = getReportData(e.parameter.month); // month มาในรูปแบบ "YYYY-MM"
    }

  } catch (err) {
    result = { status: "error", message: err.message };
  }

  // ห่อ Response ด้วย JSONP Callback เพื่อส่งกลับไปให้หน้าเว็บ HTML ได้อย่างปลอดภัย
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// -------------------------------------------------------------
// ฟังก์ชันผู้ช่วย: ดักจับและป้องกัน Error วันที่จาก Google Sheets
function parseDates(cellValue) {
  if (!cellValue) return [];
  // ถ้า Google Sheets ดันแปลงช่องนี้เป็น Date อัตโนมัติ ให้ดึงกลับมาเป็น Text แบบ YYYY-MM-DD
  if (Object.prototype.toString.call(cellValue) === '[object Date]') {
    var y = cellValue.getFullYear();
    var m = ("0" + (cellValue.getMonth() + 1)).slice(-2);
    var d = ("0" + cellValue.getDate()).slice(-2);
    return [y + "-" + m + "-" + d];
  }
  // ถ้าเป็นข้อความปกติ ให้ตัดด้วยลูกน้ำ
  return String(cellValue).split(',');
}

// -------------------------------------------------------------
// 1. ระบบเข้าสู่ระบบ (Login)
function doLogin(user, pass) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Employees');
  if (!sheet) return { status: "error", message: "ไม่พบแท็บชีท 'Employees'" };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var sheetId = String(data[i][0]).trim();
    var sheetUser = String(data[i][1]).trim();
    var sheetEmail = String(data[i][2]).trim();

    // เช็คว่า Username(Col B) หรือ Email(Col C) ตรงไหม และ Password คือ EmployeeID(Col A)
    if ((sheetUser == user || sheetEmail == user) && sheetId == pass) {
      return {
        status: "success",
        data: {
          employee_id: data[i][0],   // Col A
          name: data[i][3],          // Col D
          manager_email: data[i][4], // Col E
          department: data[i][5],    // Col F
          role: data[i][6],          // Col G
          email: data[i][2]          // Col C
        }
      };
    }
  }
  return { status: "error", message: "ชื่อผู้ใช้ หรือ รหัสผ่าน ไม่ถูกต้อง" };
}

// -------------------------------------------------------------
// 2. ส่งคำขอ WFH ใหม่
function submitRequest(payload) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  var reqId = "REQ-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);

  // ใส่ ' ไว้ข้างหน้า dates ป้องกัน Google sheets ตีความเป็น Date
  var datesToSave = "'" + payload.dates.join(',');

  sheet.appendRow([
    reqId,
    payload.employee_id,
    payload.name,
    payload.department,
    payload.email,
    payload.manager_email,
    datesToSave, // บันทึกเป็นข้อความ Text ป้องกัน Error 100%
    payload.reason,
    "Pending",
    "",
    new Date()
  ]);

  return { status: "success", message: "ยื่นเรื่องเรียบร้อย รอหัวหน้าอนุมัติ" };
}

// -------------------------------------------------------------
// 3. ดึงข้อมูลปฏิทินของพนักงาน
function getCalendarData(empId) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == empId) {
      var datesArray = parseDates(data[i][6]); // ใช้ฟังก์ชันผู้ช่วยแปลง
      for (var d = 0; d < datesArray.length; d++) {
        if (datesArray[d].trim() !== "") {
          result.push({
            date: datesArray[d].trim(),
            status: data[i][8],
            reason: data[i][7]
          });
        }
      }
    }
  }
  return { status: "success", data: result };
}

// -------------------------------------------------------------
// 4. ประวัติการลางานส่วนตัว
function getHistory(empId) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == empId) {
      result.push({
        request_id: data[i][0],
        employee_name: data[i][2],
        dates: parseDates(data[i][6]), // ใช้ฟังก์ชันผู้ช่วยแปลง
        reason: data[i][7],
        status: data[i][8],
        manager_comment: data[i][9]
      });
    }
  }
  result.reverse();
  return { status: "success", data: result };
}

// -------------------------------------------------------------
// 5. ดึงข้อมูลให้หัวหน้าอนุมัติ
function getApprovals(managerEmail, role) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  if (!sheet) return { status: "error", message: "ไม่พบแท็บชีทชื่อ 'Requests'" };

  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (role === 'Admin' || data[i][5] == managerEmail) {
      result.push({
        request_id: data[i][0],
        employee_id: data[i][1],
        name: data[i][2],
        dates: parseDates(data[i][6]), // ใช้ฟังก์ชันผู้ช่วยแปลง
        reason: data[i][7],
        status: data[i][8]
      });
    }
  }
  result.reverse();
  return { status: "success", data: result };
}

// -------------------------------------------------------------
// 6. เปลี่ยนสถานะคำขอ
function processApproval(reqId, status, managerEmail, comment) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  if (!sheet) return { status: "error", message: "ไม่พบแท็บชีทชื่อ 'Requests'" };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == reqId) {
      sheet.getRange(i + 1, 9).setValue(status);
      sheet.getRange(i + 1, 10).setValue(comment || "");
      return { status: "success", message: "อัปเดตสถานะเป็น " + status + " เรียบร้อย" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลคำขอนี้" };
}

// -------------------------------------------------------------
// 7. พนักงานแก้ไขเหตุผลและส่งใหม่
function editRequest(reqId, newReason) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == reqId) {
      sheet.getRange(i + 1, 8).setValue(newReason);
      sheet.getRange(i + 1, 9).setValue("Pending");
      sheet.getRange(i + 1, 10).setValue("");
      return { status: "success", message: "บันทึกและส่งใหม่เรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบคำขอ" };
}

// -------------------------------------------------------------
// 8. แอดมินลบคำขอ
function deleteRequest(reqId, managerEmail) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == reqId) {
      sheet.deleteRow(i + 1);
      return { status: "success", message: "ลบข้อมูลเรียบร้อยแล้ว" };
    }
  }
  return { status: "error", message: "ไม่พบข้อมูลที่ต้องการลบ" };
}

// -------------------------------------------------------------
// 9. ทำหน้ารายงาน
function getReportData(yyyy_mm) {
  var empSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Employees');
  var reqSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Requests');

  if (!empSheet || !reqSheet) return { status: "error", message: "ไม่พบฐานข้อมูลชีทที่ระบุ" };

  var empData = empSheet.getDataRange().getValues();
  var reqData = reqSheet.getDataRange().getValues();

  var resultData = [];

  for (var i = 1; i < empData.length; i++) {
    resultData.push({
      emp_id: empData[i][0],
      name: empData[i][3],
      position: empData[i][5] || "-",
      wfh_dates: []
    });
  }

  for (var r = 1; r < reqData.length; r++) {
    var status = reqData[r][8];
    if (status === "Approved") {
      var reqEmpId = reqData[r][1];
      var datesArray = parseDates(reqData[r][6]); // ใช้ฟังก์ชันผู้ช่วยแปลง

      var empIndex = resultData.findIndex(function (e) { return e.emp_id == reqEmpId; });
      if (empIndex !== -1) {
        for (var d = 0; d < datesArray.length; d++) {
          var dateItem = datesArray[d].trim();
          if (dateItem.indexOf(yyyy_mm) === 0) {
            var dayOnly = dateItem.substring(8, 10);
            resultData[empIndex].wfh_dates.push(dayOnly);
          }
        }
      }
    }
  }

  return { status: "success", data: resultData };
}

// -------------------------------------------------------------
// 10. Setup เริ่มต้น
function setupInitialSheets() {
  if (SHEET_ID === 'ใส่_ID_SHEET_ของคุณที่นี่' || SHEET_ID === '') {
    throw new Error("กรุณาใส่ SHEET_ID ของคุณที่บรรทัดที่ 2 ก่อนรันฟังก์ชันนี้นะครับ");
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);

  var empSheet = ss.getSheetByName('Employees');
  if (!empSheet) {
    empSheet = ss.insertSheet('Employees');
    var empHeaders = ["employee_id", "username", "email", "name", "manager_email", "department", "role"];
    empSheet.appendRow(empHeaders);

    var mockEmployees = [
      ["admin123", "admin", "admin@wfh.com", "ผู้ดูแลระบบ", "", "IT", "Admin"],
      ["boss123", "manager1", "manager1@wfh.com", "หัวหน้า สมทรง", "", "อก.คน.", "Manager"],
      ["user123", "user1", "user1@wfh.com", "พนักงาน สมใจ", "manager1@wfh.com", "IT", "User"]
    ];

    empSheet.getRange(2, 1, mockEmployees.length, mockEmployees[0].length).setValues(mockEmployees);
    empSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#ff758c").setFontColor("#ffffff");
    empSheet.setFrozenRows(1);
  }

  var reqSheet = ss.getSheetByName('Requests');
  if (!reqSheet) {
    reqSheet = ss.insertSheet('Requests');
    var reqHeaders = ["RequestID", "EmployeeID", "Name", "Department", "Email", "ManagerEmail", "Dates", "Reason", "Status", "ManagerComment", "Timestamp"];
    reqSheet.appendRow(reqHeaders);
    reqSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#ff758c").setFontColor("#ffffff");
    reqSheet.setFrozenRows(1);
  }
}