/**
 * ระบบ HR Dashboard + สรุปพนักงาน + รายงานเดือน (V16.0)
 * ใช้งานได้ 100% – หน้าแรกไม่ว่าง
 */

const SPREADSHEET_ID = "1iXO1kluiWECp702MLPzzRejRedVUw78E2PoAhlK0dpU"; 
const EMPLOYEE_SHEET_NAME = "Employees";
const LOG_SHEET_PREFIX = "TimeLog_"; 
const LOG_SHEET_HEADERS = [ 
  "Timestamp", "EmployeeID", "Name", "Department",
  "Action", "Latitude", "Longitude", "Remarks"
];

const CACHE_EXPIRATION_SECONDS = 60 * 15;
const OFFICE_LATITUDE = 17.659836585286648;
const OFFICE_LONGITUDE = 100.1068322094819;
const ALLOWED_RADIUS_METERS = 200;
const SHIFT_MORNING_START = "09:05:00"; 
const SHIFT_AFTERNOON_START = "12:05:00";

// -------------------------------------------------------------------
// ฟังก์ชันช่วย
// -------------------------------------------------------------------
function getLocalDateString(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getMonthlyLogSheetName(date) {
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  return `${LOG_SHEET_PREFIX}${year}-${month}`;
}

function getOrCreateLogSheet(ss, date) {
  const sheetName = getMonthlyLogSheetName(date);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(LOG_SHEET_HEADERS);
    sheet.getRange(1, 1, 1, LOG_SHEET_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function clearCacheForDate(dateString) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(`summary_${dateString}`);
  } catch (e) {}
}

// -------------------------------------------------------------------
// doGet – ทุกหน้า + หน้าแรกไม่ว่าง
// -------------------------------------------------------------------
function doGet(e) {
  const page = e.parameter.page;
  const empId = e.parameter.empId;

  // หน้า HR Dashboard
  if (page === 'hr' || !page) {
    return HtmlService.createTemplateFromFile('hr_dashboard')
      .evaluate()
      .setTitle('HR Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // หน้าสรุปพนักงาน
  if (page === 'emp_summary') {
    return HtmlService.createTemplateFromFile('employee_summary')
      .evaluate()
      .setTitle('สรุปพนักงาน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // หน้ารายงานเดือน
  if (page === 'monthly_report') {
    return HtmlService.createTemplateFromFile('monthly_report')
      .evaluate()
      .setTitle('รายงานประจำเดือน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // หน้าเข้างานพนักงาน
  if (empId) {
    const employee = lookupEmployee(empId);
    if (!employee) {
      return HtmlService.createHtmlOutput(`<h1 style="color:red;">ไม่พบพนักงานรหัส '${empId}'</h1>`);
    }
    const statusData = getEmployeeStatus(empId);
    const template = HtmlService.createTemplateFromFile('index');
    template.employee = employee;
    template.empId = empId;
    template.currentStatus = statusData.status;
    template.officeLat = OFFICE_LATITUDE;
    template.officeLng = OFFICE_LONGITUDE;
    template.officeRadius = ALLOWED_RADIUS_METERS;
    return template.evaluate()
      .setTitle('ระบบลงเวลาเข้างาน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // ไม่พบหน้า
  return HtmlService.createHtmlOutput("<h1>ไม่พบหน้า</h1>");
}

// -------------------------------------------------------------------
// พนักงาน
// -------------------------------------------------------------------
function lookupEmployee(empId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EMPLOYEE_SHEET_NAME);
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === empId.toString().trim()) {
        return {
          id: data[i][0],
          name: data[i][1],
          nickname: data[i][2] || "",
          department: data[i][3] || "",
          email: data[i][4] || "",
          phone: data[i][5] || ""
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

function getEmployeeStatus(empId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const today = new Date(); today.setHours(0,0,0,0);
    const logSheet = getOrCreateLogSheet(ss, today);
    const data = logSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const logTime = new Date(data[i][0]);
      const logEmpId = data[i][1].toString().trim();
      const logDate = new Date(logTime); logDate.setHours(0,0,0,0);
      if (logEmpId === empId && logDate.getTime() === today.getTime()) {
        const action = data[i][4];
        if (action.includes("เข้างาน") || action.includes("กลับจากพัก")) return { status: "ClockedIn" };
        if (action.includes("เริ่มพัก")) return { status: "OnBreak" };
        if (action === "เลิกงาน") return { status: "ClockedOut" };
      }
    }
    return { status: "ClockedOut" };
  } catch (e) { return { status: "ClockedOut" }; }
}

function recordTime(empId, action, userLat, userLng, distance) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const timestamp = new Date();
    const logSheet = getOrCreateLogSheet(ss, timestamp); 
    const employee = lookupEmployee(empId);
    if (!employee) return { status: "error", message: `ไม่พบพนักงาน ${empId}` };
    const remarks = (distance <= ALLOWED_RADIUS_METERS) ? "ในพื้นที่" : "นอกพื้นที่"; 
    logSheet.appendRow([timestamp, employee.id, employee.name, employee.department, action, userLat, userLng, remarks]);
    clearCacheForDate(getLocalDateString(timestamp));
    return { status: "success", message: `บันทึกสำเร็จ: ${action}`, newState: action.includes("เข้างาน") || action.includes("กลับจากพัก") ? "ClockedIn" : action.includes("เริ่มพัก") ? "OnBreak" : "ClockedOut" };
  } catch (e) { return { status: "error", message: e.message }; }
}

// -------------------------------------------------------------------
// HR Dashboard
// -------------------------------------------------------------------
function getDailySummary(dateString) {
  const cache = CacheService.getScriptCache();
  const key = `summary_${dateString}`;
  const cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const empSheet = ss.getSheetByName(EMPLOYEE_SHEET_NAME);
    const allEmp = empSheet.getDataRange().getValues().slice(1).filter(r => r[0]);
    const targetDate = new Date(dateString + "T00:00:00");
    const logSheet = ss.getSheetByName(getMonthlyLogSheetName(targetDate));
    const logs = logSheet ? logSheet.getDataRange().getValues().slice(1) : [];
    const today = new Date(); today.setHours(0,0,0,0);
    const isToday = getLocalDateString(targetDate) === getLocalDateString(today);

    const counts = { total: allEmp.length, present: 0, late: 0, sick: 0, personal: 0, vacation: 0, absent: 0, pending: 0 };
    const summary = allEmp.map(emp => {
      const empId = emp[0].toString().trim();
      const empLogs = logs.filter(l => l[1].toString().trim() === empId);
      let status = isToday ? "ยังไม่ลงเวลา (Pending)" : "ขาดงาน (Absent)";
      let clockIn = null, clockOut = null, breakStart = null, breakEnd = null, remarks = [];

      empLogs.forEach(log => {
        const action = log[4].toString();
        const time = new Date(log[0]);
        if (action.includes("เข้างาน")) { clockIn = time; status = "มาทำงาน (Present)"; }
        if (action === "เลิกงาน") clockOut = time;
        if (action.includes("เริ่มพัก")) breakStart = time;
        if (action.includes("กลับจากพัก")) breakEnd = time;
        if (action.includes("ลา") || action.includes("ขาด")) { status = action; }
        if (log[7]) remarks.push(log[7]);
      });

      if (clockIn && status === "มาทำงาน (Present)") {
        const start = new Date(targetDate);
        start.setHours(...SHIFT_MORNING_START.split(':').map(Number));
        if (clockIn > start) status = "มาสาย (Late)";
      }

      if (status.includes("Present")) counts.present++;
      else if (status.includes("Late")) counts.late++;
      else if (status.includes("ลาป่วย")) counts.sick++;
      else if (status.includes("ลากิจ")) counts.personal++;
      else if (status.includes("ลาพักร้อน")) counts.vacation++;
      else if (status.includes("ขาดงาน")) counts.absent++;
      else if (status.includes("Pending")) counts.pending++;

      return {
        empId, name: emp[1], department: emp[3],
        status,
        clockIn: clockIn ? clockIn.toLocaleTimeString('th-TH') : "N/A",
        clockOut: clockOut ? clockOut.toLocaleTimeString('th-TH') : "N/A",
        breakStart: breakStart ? breakStart.toLocaleTimeString('th-TH') : "N/A",
        breakEnd: breakEnd ? breakEnd.toLocaleTimeString('th-TH') : "N/A",
        remarks: remarks.join(', ')
      };
    });

    counts.present += counts.late;
    const result = { summary, counts };
    cache.put(key, JSON.stringify(result), CACHE_EXPIRATION_SECONDS);
    return result;
  } catch (e) { return { error: e.message }; }
}

function manualUpdateLog(data) {
  try {
    const { empId, date, action, remarks } = data;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = getOrCreateLogSheet(ss, new Date(date));
    const employee = lookupEmployee(empId);
    if (!employee) return { status: "error", message: "ไม่พบพนักงาน" };
    logSheet.appendRow([new Date(date), employee.id, employee.name, employee.department, action, null, null, `[HR] ${remarks || ""}`]);
    clearCacheForDate(date);
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; }
}

// -------------------------------------------------------------------
// สรุปพนักงาน + รายงานเดือน
// -------------------------------------------------------------------
function getAllEmployees() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EMPLOYEE_SHEET_NAME);
    const data = sheet.getDataRange().getValues().slice(1);
    return data.map(row => ({ id: row[0], name: row[1], department: row[3] || "" }));
  } catch (e) { return []; }
}

function getEmployeeMonthlySummary(empId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const sheetName = `${LOG_SHEET_PREFIX}${year}-${month}`;
    const logSheet = ss.getSheetByName(sheetName);
    if (!logSheet) return { summary: { present: 0, late: 0, absent: 0 }, logs: [] };

    const logs = logSheet.getDataRange().getValues().slice(1);
    const empLogs = logs.filter(row => row[1].toString().trim() === empId);
    const summary = { present: 0, late: 0, absent: 0 };
    const daily = {};

    empLogs.forEach(log => {
      const date = Utilities.formatDate(new Date(log[0]), "GMT+7", "yyyy-MM-dd");
      if (!daily[date]) daily[date] = { clockIn: "N/A", clockOut: "N/A", breakStart: "N/A", breakEnd: "N/A", status: "ขาดงาน", remarks: "" };
      const action = log[4].toString();
      const time = Utilities.formatDate(new Date(log[0]), "GMT+7", "HH:mm:ss");
      if (action.includes("เข้างาน")) { daily[date].clockIn = time; daily[date].status = "มาทำงาน"; }
      if (action.includes("เลิกงาน")) daily[date].clockOut = time;
      if (action.includes("เริ่มพัก")) daily[date].breakStart = time;
      if (action.includes("กลับจากพัก")) daily[date].breakEnd = time;
      if (action.includes("ลา") || action.includes("ขาด")) { daily[date].status = action; daily[date].remarks = log[7] || ""; }
    });

    Object.values(daily).forEach(d => {
      if (d.status.includes("มาทำงาน")) summary.present++;
      else if (d.status.includes("สาย")) summary.late++;
      else summary.absent++;
    });

    const logList = Object.keys(daily).map(date => ({
      date, clockIn: daily[date].clockIn, clockOut: daily[date].clockOut,
      breakStart: daily[date].breakStart, breakEnd: daily[date].breakEnd,
      status: daily[date].status, remarks: daily[date].remarks
    })).reverse();

    return { summary, logs: logList };
  } catch (e) { throw new Error(e.message); }
}

function getMonthlyReport(year, month) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = `${LOG_SHEET_PREFIX}${year}-${month}`;
    const logSheet = ss.getSheetByName(sheetName);
    if (!logSheet) return { total: { workingDays: 0, present: 0, late: 0, absent: 0 }, byEmployee: [] };

    const logs = logSheet.getDataRange().getValues().slice(1);
    const empData = ss.getSheetByName(EMPLOYEE_SHEET_NAME).getDataRange().getValues().slice(1);
    const total = { workingDays: 0, present: 0, late: 0, absent: 0 };
    const byEmp = {};
    
    // สร้าง object พนักงาน
    empData.forEach(row => {
      const id = row[0].toString().trim();
      byEmp[id] = { id, name: row[1], dept: row[3] || "", present: 0, late: 0, absent: 0 };
    });

    const dates = new Set();
    const MORNING_START = new Date(`1970-01-01T${SHIFT_MORNING_START}Z`); // 09:00:00

    logs.forEach(log => {
      const timestamp = new Date(log[0]);
      const date = Utilities.formatDate(timestamp, "GMT+7", "yyyy-MM-dd");
      dates.add(date);

      const empId = log[1].toString().trim();
      if (!byEmp[empId]) return;

      const action = log[4].toString();
      const timeStr = Utilities.formatDate(timestamp, "GMT+7", "HH:mm:ss");
      const logTime = new Date(`1970-01-01T${timeStr}Z`);

      // ตรวจสอบ "เข้างาน"
      if (action.includes("เข้างาน")) {
        if (logTime > MORNING_START) {
          byEmp[empId].late++;     // มาสาย
          total.late++;
        } else {
          byEmp[empId].present++;  // มาทำงานตรงเวลา
          total.present++;
        }
      }

      // ตรวจสอบ "ลา" หรือ "ขาด"
      if (action.includes("ลา") || action.includes("ขาด")) {
        byEmp[empId].absent++;
        total.absent++;
      }
    });

    total.workingDays = dates.size;

    return {
      total,
      byEmployee: Object.values(byEmp).filter(e => e.present + e.late + e.absent > 0)
    };
  } catch (e) {
    throw new Error(e.message);
  }
}
