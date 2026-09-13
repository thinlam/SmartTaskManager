/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 5: Setup (tao workbook structure)
 *
 * setupSmartTaskManager() la entry point duy nhat nguoi dung can chay.
 * PHAI idempotent: chay lai nhieu lan khong duoc:
 * - tao trung sheet
 * - ghi de du lieu Task da co
 * - tao trung data validation / conditional formatting
 *
 * Phase nay CHI build day du cho 02_Tasks, 10_Activity_Log, 12_Settings
 * (dung pham vi yeu cau hien tai). 11 sheet con lai duoc tao voi ghi chu
 * toi thieu de dat dung cho trong cau truc 14 sheet - noi dung/logic cua
 * chung thuoc Phase 7+ (Dashboard, Kanban, Calendar, Gantt, Reports...).
 *
 * Phu thuoc global: types.gs, constants.gs, config.gs, utils.gs
 */

/**
 * Entry point chinh - chay 1 lan de khoi tao toan bo workbook.
 * An toan khi chay lai nhieu lan (idempotent).
 */
function setupSmartTaskManager() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createAllSheets_(ss);
  setupTasksSheet_(ss);
  setupActivityLogSheet_(ss);
  setupSettingsSheet_(ss);
  removeDefaultBlankSheetIfUnused_(ss);

  SpreadsheetApp.flush();
  Logger.log("setupSmartTaskManager() hoan tat.");
}

// ==================================================
// SHEET CREATION (idempotent)
// ==================================================

/**
 * Tao du 14 sheet theo dung thu tu SHEET_NAMES neu chua ton tai.
 * Sheet da ton tai thi giu nguyen, khong ghi de.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createAllSheets_(ss) {
  var orderedNames = [
    SHEET_NAMES.DASHBOARD,
    SHEET_NAMES.TASKS,
    SHEET_NAMES.CALENDAR,
    SHEET_NAMES.KANBAN,
    SHEET_NAMES.GANTT,
    SHEET_NAMES.PROJECTS,
    SHEET_NAMES.MEMBERS,
    SHEET_NAMES.CATEGORIES,
    SHEET_NAMES.REPORTS,
    SHEET_NAMES.ACTIVITY_LOG,
    SHEET_NAMES.NOTIFICATIONS,
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.LISTS,
    SHEET_NAMES.INSTRUCTIONS
  ];

  for (var i = 0; i < orderedNames.length; i++) {
    var name = orderedNames[i];
    var sheet = getOrCreateSheet_(ss, name);

    var reservedForLaterPhase =
      name !== SHEET_NAMES.TASKS &&
      name !== SHEET_NAMES.ACTIVITY_LOG &&
      name !== SHEET_NAMES.SETTINGS;

    if (reservedForLaterPhase && sheet.getRange("A1").getValue() === "") {
      sheet.getRange("A1").setValue(
        "[" + name + "] Se duoc xay dung o phase sau (Dashboard/Kanban/Calendar/Gantt/Reports...)."
      );
      sheet.getRange("A1").setFontStyle("italic").setFontColor(DESIGN_COLORS.TEXT_SECONDARY);
    }
  }

  for (var pos = 0; pos < orderedNames.length; pos++) {
    var s = ss.getSheetByName(orderedNames[pos]);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(pos + 1);
    }
  }
}

/**
 * Lay sheet theo ten, tao moi neu chua co. Khong bao gio tao trung.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  return ss.insertSheet(name);
}

/**
 * Sheet "Sheet1" mac dinh cua Google khi tao Spreadsheet moi - xoa di neu
 * no dang trong va khong phai 1 trong 14 sheet chinh thuc.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function removeDefaultBlankSheetIfUnused_(ss) {
  var officialNames = [
    SHEET_NAMES.DASHBOARD, SHEET_NAMES.TASKS, SHEET_NAMES.CALENDAR,
    SHEET_NAMES.KANBAN, SHEET_NAMES.GANTT, SHEET_NAMES.PROJECTS,
    SHEET_NAMES.MEMBERS, SHEET_NAMES.CATEGORIES, SHEET_NAMES.REPORTS,
    SHEET_NAMES.ACTIVITY_LOG, SHEET_NAMES.NOTIFICATIONS, SHEET_NAMES.SETTINGS,
    SHEET_NAMES.LISTS, SHEET_NAMES.INSTRUCTIONS
  ];
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var isOfficial = officialNames.indexOf(name) !== -1;
    var looksLikeDefault = /^Sheet\d*$/.test(name);
    var isEmpty = sheet.getLastRow() === 0 && sheet.getLastColumn() === 0;

    if (!isOfficial && looksLikeDefault && isEmpty && sheets.length > 1) {
      ss.deleteSheet(sheet);
    }
  }
}

// ==================================================
// TASKS SHEET SETUP
// ==================================================

/**
 * Tao header + dinh dang + data validation cho sheet 02_Tasks.
 * Idempotent: neu header da dung thi khong ghi de, khong xoa du lieu Task
 * da nhap ben duoi.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupTasksSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.TASKS);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.TASKS + " chua duoc tao.");

  var headerRange = sheet.getRange(1, 1, 1, TASK_HEADERS.length);
  var currentHeader = headerRange.getValues()[0];
  var headerMatches = arraysEqual_(currentHeader, TASK_HEADERS);

  if (!headerMatches) {
    headerRange.setValues([TASK_HEADERS]);
  }

  headerRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(DESIGN_COLORS.PRIMARY)
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  sheet.setColumnWidth(TASK_COL.taskName + 1, 220);
  sheet.setColumnWidth(TASK_COL.description + 1, 260);
  sheet.setColumnWidth(TASK_COL.notes + 1, 220);

  setupTasksDataValidation_(sheet);
  setupTasksConditionalFormatting_(sheet);
}

/**
 * Dropdown cho Status/Priority/RecurringType + gioi han Progress 0-100.
 * setDataValidation() luon ghi de bang dung 1 rule moi nhat -> idempotent.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupTasksDataValidation_(sheet) {
  var maxRows = Math.max(sheet.getMaxRows(), 1000);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TASK_COL.status + 1, maxRows - 1, 1).setDataValidation(statusRule);

  var priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PRIORITY_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TASK_COL.priority + 1, maxRows - 1, 1).setDataValidation(priorityRule);

  var recurringValues = [
    RECURRING_TYPE.NONE, RECURRING_TYPE.DAILY, RECURRING_TYPE.WEEKLY,
    RECURRING_TYPE.MONTHLY, RECURRING_TYPE.QUARTERLY
  ];
  var recurringRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(recurringValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TASK_COL.recurringType + 1, maxRows - 1, 1).setDataValidation(recurringRule);

  var progressRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 100)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, TASK_COL.progress + 1, maxRows - 1, 1).setDataValidation(progressRule);
}

/**
 * Conditional Formatting cho cot Status/Priority/RiskLevel - loai bo rule
 * cu cua dung cac range nay truoc khi them lai, tranh cong don qua nhieu
 * lan chay setup.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function setupTasksConditionalFormatting_(sheet) {
  var maxRows = Math.max(sheet.getMaxRows(), 1000);
  var allRules = sheet.getConditionalFormatRules();

  var statusRange = sheet.getRange(2, TASK_COL.status + 1, maxRows - 1, 1);
  var priorityRange = sheet.getRange(2, TASK_COL.priority + 1, maxRows - 1, 1);
  var riskRange = sheet.getRange(2, TASK_COL.riskLevel + 1, maxRows - 1, 1);

  var statusA1 = statusRange.getA1Notation();
  var priorityA1 = priorityRange.getA1Notation();
  var riskA1 = riskRange.getA1Notation();

  var keptRules = allRules.filter(function (rule) {
    var ranges = rule.getRanges().map(function (r) { return r.getA1Notation(); });
    var touchesManagedRange =
      ranges.indexOf(statusA1) !== -1 ||
      ranges.indexOf(priorityA1) !== -1 ||
      ranges.indexOf(riskA1) !== -1;
    return !touchesManagedRange;
  });

  var newRules = [];

  STATUS_LIST.forEach(function (statusValue) {
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(statusValue)
        .setBackground(STATUS_COLORS[statusValue])
        .setFontColor("#FFFFFF")
        .setRanges([statusRange])
        .build()
    );
  });

  PRIORITY_LIST.forEach(function (priorityValue) {
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(priorityValue)
        .setBackground(PRIORITY_COLORS[priorityValue])
        .setFontColor("#FFFFFF")
        .setRanges([priorityRange])
        .build()
    );
  });

  [RISK_LEVEL.LOW, RISK_LEVEL.MEDIUM, RISK_LEVEL.HIGH, RISK_LEVEL.CRITICAL].forEach(function (riskValue) {
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(riskValue)
        .setBackground(RISK_COLORS[riskValue])
        .setFontColor("#FFFFFF")
        .setRanges([riskRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(keptRules.concat(newRules));
}

// ==================================================
// ACTIVITY LOG SHEET SETUP
// ==================================================

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupActivityLogSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.ACTIVITY_LOG);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.ACTIVITY_LOG + " chua duoc tao.");

  var headerRange = sheet.getRange(1, 1, 1, ACTIVITY_LOG_HEADERS.length);
  var currentHeader = headerRange.getValues()[0];

  if (!arraysEqual_(currentHeader, ACTIVITY_LOG_HEADERS)) {
    headerRange.setValues([ACTIVITY_LOG_HEADERS]);
  }

  headerRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(DESIGN_COLORS.PRIMARY_DARK)
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
}

// ==================================================
// SETTINGS SHEET SETUP
// ==================================================

/**
 * Sheet Settings luu dang 3 cot Key/Value/Description, doc bang
 * getSetting_(key). Chi ghi gia tri mac dinh cho key CHUA ton tai - khong
 * ghi de key ma nguoi dung da tu chinh truoc do.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.SETTINGS + " chua duoc tao.");

  var expectedHeader = ["Key", "Value", "Description"];
  var headerRange = sheet.getRange(1, 1, 1, 3);
  var currentHeader = headerRange.getValues()[0];

  if (!arraysEqual_(currentHeader, expectedHeader)) {
    headerRange.setValues([expectedHeader]);
    headerRange
      .setFontWeight("bold")
      .setFontColor("#FFFFFF")
      .setBackground(DESIGN_COLORS.PRIMARY)
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }

  var existingKeys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existingData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(function (row) {
      if (row[0]) existingKeys[row[0]] = true;
    });
  }

  var defaultRows = [
    ["systemName", DEFAULT_APP_SETTINGS.systemName, "Ten he thong hien thi tren Dashboard"],
    ["defaultStatus", DEFAULT_APP_SETTINGS.defaultStatus, "Status mac dinh khi tao Task moi"],
    ["defaultPriority", DEFAULT_APP_SETTINGS.defaultPriority, "Priority mac dinh khi tao Task moi"],
    ["dueSoonDays", DEFAULT_APP_SETTINGS.dueSoonDays, "So ngay tinh la 'Due Soon' tren KPI"],
    ["workingDays", DEFAULT_APP_SETTINGS.workingDays.join(","), "Cac ngay lam viec trong tuan (1=Mon..7=Sun)"],
    ["workingHoursPerDay", DEFAULT_APP_SETTINGS.workingHoursPerDay, "So gio lam viec/ngay"],
    ["weekendDays", DEFAULT_APP_SETTINGS.weekendDays.join(","), "Cac ngay cuoi tuan"],
    ["notificationsEnabled", DEFAULT_APP_SETTINGS.notificationsEnabled, "Bat/tat toan bo notification"],
    ["notificationEmailOnOverdue", DEFAULT_APP_SETTINGS.notificationEmailOnOverdue, "Gui email khi Task Overdue"],
    ["notificationEmailOnDueToday", DEFAULT_APP_SETTINGS.notificationEmailOnDueToday, "Gui email khi Task Due Today"],
    ["notificationEmailOnDueTomorrow", DEFAULT_APP_SETTINGS.notificationEmailOnDueTomorrow, "Gui email khi Task Due Tomorrow"],
    ["notificationEmailOnCritical", DEFAULT_APP_SETTINGS.notificationEmailOnCritical, "Gui email khi Task Critical"],
    ["theme", DEFAULT_APP_SETTINGS.theme, "Light hoac Dark"],
    ["dateFormat", DEFAULT_APP_SETTINGS.dateFormat, "Dinh dang ngay hien thi"],
    ["currency", DEFAULT_APP_SETTINGS.currency, "Don vi tien te"],
    ["language", DEFAULT_APP_SETTINGS.language, "Ngon ngu he thong"]
  ];

  var rowsToAppend = defaultRows.filter(function (row) {
    return !existingKeys[row[0]];
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 320);
}

/**
 * Doc 1 gia tri Setting theo key. Tra ve defaultValue neu key khong ton tai.
 * @param {string} key
 * @param {*} [defaultValue]
 * @returns {*}
 */
function getSetting_(key, defaultValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return defaultValue;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return defaultValue;

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return defaultValue;
}

// ==================================================
// SHARED HELPER
// ==================================================

/**
 * So sanh 2 mang gia tri (dung de kiem tra header da dung chua).
 * @param {Array} a
 * @param {Array} b
 * @returns {boolean}
 */
function arraysEqual_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (String(a[i]) !== String(b[i])) return false;
  }
  return true;
}