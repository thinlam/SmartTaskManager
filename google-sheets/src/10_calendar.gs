/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 8b: Calendar
 *
 * refreshCalendar() la view phai sinh tu 02_Tasks (Single Source of
 * Truth) - hien thi Task theo DueDate tren luoi thang. Moi lan goi doc
 * lai toan bo Task va ghi de TOAN BO vung Calendar.
 *
 * Mac dinh hien thi thang/nam HIEN TAI. Co the goi
 * refreshCalendar(year, month) de xem thang khac (month: 1-12).
 *
 * Phu thuoc global: types.gs, constants.gs, tasks.gs
 */

var CALENDAR_LAYOUT = Object.freeze({
  TITLE_ROW: 1,
  WEEKDAY_HEADER_ROW: 2,
  GRID_START_ROW: 3,
  MAX_WEEKS: 6,
  DAY_CELL_HEIGHT: 70,
  MAX_TASKS_SHOWN_PER_DAY: 3
});

var WEEKDAY_LABELS = Object.freeze(["Thu 2", "Thu 3", "Thu 4", "Thu 5", "Thu 6", "Thu 7", "CN"]);

/**
 * @param {number} [year] - mac dinh nam hien tai
 * @param {number} [month] - 1-12, mac dinh thang hien tai
 */
function refreshCalendar(year, month) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CALENDAR);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.CALENDAR + " chua duoc tao. Chay setupSmartTaskManager() truoc.");

  var now = new Date();
  var targetYear = year || now.getFullYear();
  var targetMonth = month || (now.getMonth() + 1); // 1-12

  var tasks = getAllTasks();
  var tasksByDate = groupTasksByDueDate_(tasks);

  clearCalendarManagedRange_(sheet);
  writeCalendarHeader_(sheet, targetYear, targetMonth);
  writeCalendarGrid_(sheet, targetYear, targetMonth, tasksByDate, now);

  SpreadsheetApp.flush();
  Logger.log("refreshCalendar() hoan tat cho thang " + targetMonth + "/" + targetYear);
}

function clearCalendarManagedRange_(sheet) {
  var maxCol = 7;
  var maxRow = CALENDAR_LAYOUT.GRID_START_ROW + CALENDAR_LAYOUT.MAX_WEEKS * 1 + 2;
  sheet.getRange(1, 1, maxRow, maxCol).clearContent();
  sheet.getRange(1, 1, maxRow, maxCol).clearFormat();
}

/**
 * @param {Array<Object>} tasks
 * @returns {Object} map "yyyy-MM-dd" -> Array<Task>
 */
function groupTasksByDueDate_(tasks) {
  var map = {};
  tasks.forEach(function (t) {
    if (!t.dueDate) return;
    if (t.status === STATUS.CANCELLED) return;
    var key = String(t.dueDate).substring(0, 10);
    if (!map[key]) map[key] = [];
    map[key].push(t);
  });
  // Trong tung ngay, uu tien task SmartScore cao len truoc.
  Object.keys(map).forEach(function (key) {
    map[key].sort(function (a, b) { return Number(b.smartScore) - Number(a.smartScore); });
  });
  return map;
}

function writeCalendarHeader_(sheet, year, month) {
  var monthNames = [
    "Thang 1", "Thang 2", "Thang 3", "Thang 4", "Thang 5", "Thang 6",
    "Thang 7", "Thang 8", "Thang 9", "Thang 10", "Thang 11", "Thang 12"
  ];

  var titleRange = sheet.getRange(CALENDAR_LAYOUT.TITLE_ROW, 1, 1, 7);
  titleRange.merge();
  titleRange.setValue(monthNames[month - 1] + " " + year);
  titleRange
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontColor(DESIGN_COLORS.TEXT_PRIMARY);

  var weekdayRange = sheet.getRange(CALENDAR_LAYOUT.WEEKDAY_HEADER_ROW, 1, 1, 7);
  weekdayRange.setValues([WEEKDAY_LABELS]);
  weekdayRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(DESIGN_COLORS.PRIMARY)
    .setHorizontalAlignment("center");

  for (var c = 1; c <= 7; c++) sheet.setColumnWidth(c, 150);
}

function writeCalendarGrid_(sheet, year, month, tasksByDate, now) {
  var firstDayOfMonth = new Date(year, month - 1, 1);
  var daysInMonth = new Date(year, month, 0).getDate();

  // Chuyen Sunday=0 ve he Thu2=0..CN=6 de khop WEEKDAY_LABELS.
  var firstWeekday = (firstDayOfMonth.getDay() + 6) % 7;

  var todayKey = formatDateOnly(getTodayDateOnly(now));
  var dayCounter = 1;

  for (var week = 0; week < CALENDAR_LAYOUT.MAX_WEEKS; week++) {
    var rowIndex = CALENDAR_LAYOUT.GRID_START_ROW + week;
    if (dayCounter > daysInMonth) break;

    for (var col = 0; col < 7; col++) {
      var cell = sheet.getRange(rowIndex, col + 1);
      var isValidDay = !(week === 0 && col < firstWeekday) && dayCounter <= daysInMonth;

      if (!isValidDay) {
        cell.setBackground("#F1F5F9");
        continue;
      }

      var currentDate = new Date(year, month - 1, dayCounter);
      var dateKey = formatDateOnly(currentDate);
      var dayTasks = tasksByDate[dateKey] || [];
      var isToday = dateKey === todayKey;

      var lines = [String(dayCounter)];
      var hasOverdue = false;
      var hasDueToday = false;

      dayTasks.slice(0, CALENDAR_LAYOUT.MAX_TASKS_SHOWN_PER_DAY).forEach(function (t) {
        var marker = t.status === STATUS.COMPLETED ? "[Done] " : "";
        lines.push(marker + t.taskId + ": " + truncateText_(t.taskName, 22));
        if (t.deadlineStatus === DEADLINE_STATUS.OVERDUE) hasOverdue = true;
        if (t.deadlineStatus === DEADLINE_STATUS.DUE_TODAY) hasDueToday = true;
      });

      if (dayTasks.length > CALENDAR_LAYOUT.MAX_TASKS_SHOWN_PER_DAY) {
        lines.push("+" + (dayTasks.length - CALENDAR_LAYOUT.MAX_TASKS_SHOWN_PER_DAY) + " task khac");
      }

      cell.setValue(lines.join("\n"));
      cell.setWrap(true);
      cell.setVerticalAlignment("top");
      cell.setFontSize(9);

      var bg = DESIGN_COLORS.SURFACE;
      if (hasOverdue) bg = "#FEE2E2";
      else if (hasDueToday) bg = "#FFEDD5";
      cell.setBackground(bg);

      if (isToday) {
        cell.setBorder(true, true, true, true, false, false, DESIGN_COLORS.PRIMARY, SpreadsheetApp.BorderStyle.SOLID_THICK);
      } else {
        cell.setBorder(true, true, true, true, false, false, DESIGN_COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
      }

      dayCounter++;
    }

    sheet.setRowHeight(rowIndex, CALENDAR_LAYOUT.DAY_CELL_HEIGHT);
  }
}

function truncateText_(text, maxLen) {
  var str = String(text || "");
  return str.length > maxLen ? str.substring(0, maxLen - 1) + "…" : str;
}