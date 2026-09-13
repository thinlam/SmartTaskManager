/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 7: Dashboard
 *
 * refreshDashboard() la entry point chinh - doc toan bo Task (qua
 * getAllTasks() da co o Phase 5), tinh KPI/Top Tasks/Attention trong
 * memory, roi ghi lai TOAN BO vung Dashboard bang 1-2 lan setValues()
 * batch. Khong doc/ghi Tasks o day - Dashboard luon la view phai sinh
 * (derived view) tu 02_Tasks, dung nguyen tac "Tasks la Single Source
 * of Truth" (muc XXVII/XXVIII).
 *
 * Idempotent: goi lai nhieu lan chi ghi de dung vung Dashboard, khong
 * tao them sheet/table, khong dung lai du lieu cu con sot.
 *
 * Phu thuoc global: types.gs, constants.gs, config.gs, utils.gs,
 * smart-engine.gs, setup.gs, tasks.gs
 */

// ==================================================
// LAYOUT CONFIG (vi tri cac vung tren sheet 01_Dashboard)
// ==================================================

var DASHBOARD_LAYOUT = Object.freeze({
  TITLE_ROW: 1,
  SUBTITLE_ROW: 2,
  KPI_LABEL_ROW: 4,
  KPI_VALUE_ROW: 5,
  KPI_COL_WIDTH: 2, // moi KPI card chiem 2 cot (gop 1 the trong report, khong dung merge de don gian)
  KPI_COUNT: 8,

  TOP_TASKS_TITLE_ROW: 8,
  TOP_TASKS_HEADER_ROW: 9,
  TOP_TASKS_MAX_ROWS: 10,

  ATTENTION_TITLE_ROW: 21,
  ATTENTION_HEADER_ROW: 22,
  ATTENTION_MAX_ROWS: 10
});

// ==================================================
// ENTRY POINT
// ==================================================

/**
 * Tinh lai toan bo Dashboard tu du lieu Tasks hien tai va ghi len sheet
 * 01_Dashboard. An toan khi goi nhieu lan (idempotent) - luon ghi de dung
 * vung quan ly, khong cong don noi dung cu.
 */
function refreshDashboard() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DASHBOARD);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.DASHBOARD + " chua duoc tao. Chay setupSmartTaskManager() truoc.");

  var tasks = getAllTasks();
  var now = new Date();

  clearDashboardManagedRanges_(sheet);

  writeDashboardHeader_(sheet, now);

  var kpis = computeDashboardKpis_(tasks, now);
  writeKpiCards_(sheet, kpis);

  var topTasks = getTopTasksToDoNow_(tasks, DASHBOARD_LAYOUT.TOP_TASKS_MAX_ROWS);
  writeTopTasksSection_(sheet, topTasks);

  var attentionTasks = getTasksNeedAttention_(tasks, DASHBOARD_LAYOUT.ATTENTION_MAX_ROWS);
  writeAttentionSection_(sheet, attentionTasks);

  SpreadsheetApp.flush();
  Logger.log("refreshDashboard() hoan tat. Tong so Task: " + tasks.length);
}

// ==================================================
// CLEAR (idempotent - xoa dung vung se ghi lai, khong dung merge phuc tap)
// ==================================================

function clearDashboardManagedRanges_(sheet) {
  var maxCol = 12;
  var maxRow = DASHBOARD_LAYOUT.ATTENTION_HEADER_ROW + DASHBOARD_LAYOUT.ATTENTION_MAX_ROWS + 2;
  sheet.getRange(1, 1, maxRow, maxCol).clearContent();
  sheet.getRange(1, 1, maxRow, maxCol).clearFormat();
}

// ==================================================
// HEADER
// ==================================================

function writeDashboardHeader_(sheet, now) {
  var systemName = getSetting_("systemName", DEFAULT_APP_SETTINGS.systemName);

  var titleRange = sheet.getRange(DASHBOARD_LAYOUT.TITLE_ROW, 1, 1, 6);
  titleRange.setValue(String(systemName).toUpperCase());
  titleRange.merge();
  titleRange
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontColor(DESIGN_COLORS.TEXT_PRIMARY);

  var subtitleRange = sheet.getRange(DASHBOARD_LAYOUT.SUBTITLE_ROW, 1, 1, 6);
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "EEEE, dd/MM/yyyy HH:mm");
  subtitleRange.setValue("Cap nhat luc: " + dateStr);
  subtitleRange.merge();
  subtitleRange.setFontColor(DESIGN_COLORS.TEXT_SECONDARY).setFontStyle("italic");
}

// ==================================================
// KPI COMPUTATION
// ==================================================

/**
 * @param {Array<Object>} tasks
 * @param {Date} now
 * @returns {Array<{label:string, value:string, color:string}>}
 */
function computeDashboardKpis_(tasks, now) {
  var total = tasks.length;
  var completed = 0;
  var inProgress = 0;
  var overdue = 0;
  var dueToday = 0;
  var dueSoon = 0;
  var blocked = 0;

  var dueSoonDays = Number(getSetting_("dueSoonDays", DEADLINE_CONFIG.dueSoonDays));

  tasks.forEach(function (t) {
    if (t.status === STATUS.COMPLETED) completed++;
    if (t.status === STATUS.IN_PROGRESS) inProgress++;
    if (t.isBlocked) blocked++;
    if (t.deadlineStatus === DEADLINE_STATUS.OVERDUE) overdue++;
    if (t.deadlineStatus === DEADLINE_STATUS.DUE_TODAY) dueToday++;
    if (
      t.daysRemaining !== null &&
      t.daysRemaining !== "" &&
      Number(t.daysRemaining) >= 0 &&
      Number(t.daysRemaining) <= dueSoonDays &&
      t.status !== STATUS.COMPLETED
    ) {
      dueSoon++;
    }
  });

  var completionRate = total > 0 ? Math.round((completed / total) * 100) : 0;

  return [
    { label: "TOTAL TASKS", value: String(total), color: DESIGN_COLORS.PRIMARY },
    { label: "COMPLETED", value: String(completed), color: STATUS_COLORS[STATUS.COMPLETED] },
    { label: "IN PROGRESS", value: String(inProgress), color: STATUS_COLORS[STATUS.IN_PROGRESS] },
    { label: "OVERDUE", value: String(overdue), color: DEADLINE_COLORS[DEADLINE_STATUS.OVERDUE] },
    { label: "DUE TODAY", value: String(dueToday), color: DEADLINE_COLORS[DEADLINE_STATUS.DUE_TODAY] },
    { label: "DUE SOON", value: String(dueSoon), color: DEADLINE_COLORS[DEADLINE_STATUS.DUE_IN_3_DAYS] },
    { label: "BLOCKED", value: String(blocked), color: STATUS_COLORS[STATUS.BLOCKED] },
    { label: "COMPLETION RATE", value: completionRate + "%", color: DESIGN_COLORS.SUCCESS }
  ];
}

function writeKpiCards_(sheet, kpis) {
  var labelRow = DASHBOARD_LAYOUT.KPI_LABEL_ROW;
  var valueRow = DASHBOARD_LAYOUT.KPI_VALUE_ROW;

  for (var i = 0; i < kpis.length; i++) {
    var col = i + 1;
    var kpi = kpis[i];

    var labelCell = sheet.getRange(labelRow, col);
    labelCell.setValue(kpi.label);
    labelCell.setFontSize(9).setFontColor(DESIGN_COLORS.TEXT_SECONDARY).setFontWeight("bold");

    var valueCell = sheet.getRange(valueRow, col);
    valueCell.setValue(kpi.value);
    valueCell
      .setFontSize(18)
      .setFontWeight("bold")
      .setFontColor(kpi.color)
      .setBackground(DESIGN_COLORS.SURFACE);

    sheet.setColumnWidth(col, 110);
  }
}

// ==================================================
// TOP TASKS TO DO NOW
// ==================================================

/**
 * Lay N Task can chu y nhat, sap xep theo SmartScore giam dan.
 * Bo qua Task da Completed/Cancelled (khong can "lam ngay" nua).
 * @param {Array<Object>} tasks
 * @param {number} limit
 * @returns {Array<Object>}
 */
function getTopTasksToDoNow_(tasks, limit) {
  return tasks
    .filter(function (t) {
      return t.status !== STATUS.COMPLETED && t.status !== STATUS.CANCELLED;
    })
    .sort(function (a, b) {
      return Number(b.smartScore) - Number(a.smartScore);
    })
    .slice(0, limit);
}

function writeTopTasksSection_(sheet, topTasks) {
  var titleRow = DASHBOARD_LAYOUT.TOP_TASKS_TITLE_ROW;
  var headerRow = DASHBOARD_LAYOUT.TOP_TASKS_HEADER_ROW;

  var titleRange = sheet.getRange(titleRow, 1, 1, 6);
  titleRange.setValue("TOP TASKS TO DO NOW");
  titleRange.merge();
  titleRange.setFontSize(13).setFontWeight("bold").setFontColor(DESIGN_COLORS.TEXT_PRIMARY);

  var headers = ["Task ID", "Task Name", "Priority", "Smart Score", "Deadline Status", "Recommended Action"];
  var headerRange = sheet.getRange(headerRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(DESIGN_COLORS.PRIMARY_DARK);

  if (topTasks.length === 0) {
    sheet.getRange(headerRow + 1, 1).setValue("Khong co task nao can chu y.");
    sheet.getRange(headerRow + 1, 1).setFontColor(DESIGN_COLORS.TEXT_SECONDARY).setFontStyle("italic");
    return;
  }

  var rows = topTasks.map(function (t) {
    return [t.taskId, t.taskName, t.priority, t.smartScore, t.deadlineStatus, t.recommendedAction];
  });

  var dataRange = sheet.getRange(headerRow + 1, 1, rows.length, headers.length);
  dataRange.setValues(rows);

  // To mau cot Priority va Deadline Status theo dung design token, khong
  // dung Conditional Formatting o day vi vung nay duoc ghi lai toan bo
  // moi lan refresh - to mau truc tiep don gian hon.
  for (var r = 0; r < rows.length; r++) {
    var priorityCell = sheet.getRange(headerRow + 1 + r, 3);
    priorityCell.setFontColor(PRIORITY_COLORS[rows[r][2]]).setFontWeight("bold");

    var deadlineCell = sheet.getRange(headerRow + 1 + r, 5);
    var deadlineColor = DEADLINE_COLORS[rows[r][4]];
    if (deadlineColor) deadlineCell.setFontColor(deadlineColor).setFontWeight("bold");
  }
}

// ==================================================
// TASKS NEED ATTENTION (Overdue + Blocked + Critical Risk)
// ==================================================

/**
 * @param {Array<Object>} tasks
 * @param {number} limit
 * @returns {Array<Object>}
 */
function getTasksNeedAttention_(tasks, limit) {
  return tasks
    .filter(function (t) {
      if (t.status === STATUS.COMPLETED || t.status === STATUS.CANCELLED) return false;
      return (
        t.deadlineStatus === DEADLINE_STATUS.OVERDUE ||
        t.isBlocked === true ||
        t.riskLevel === RISK_LEVEL.CRITICAL
      );
    })
    .sort(function (a, b) {
      return Number(b.smartScore) - Number(a.smartScore);
    })
    .slice(0, limit);
}

function writeAttentionSection_(sheet, attentionTasks) {
  var titleRow = DASHBOARD_LAYOUT.ATTENTION_TITLE_ROW;
  var headerRow = DASHBOARD_LAYOUT.ATTENTION_HEADER_ROW;

  var titleRange = sheet.getRange(titleRow, 1, 1, 6);
  titleRange.setValue("TASKS NEED ATTENTION");
  titleRange.merge();
  titleRange.setFontSize(13).setFontWeight("bold").setFontColor(DESIGN_COLORS.DANGER);

  var headers = ["Task ID", "Task Name", "Priority", "Deadline Status", "Risk Level", "Health Status"];
  var headerRange = sheet.getRange(headerRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(DESIGN_COLORS.DANGER);

  if (attentionTasks.length === 0) {
    sheet.getRange(headerRow + 1, 1).setValue("Khong co task nao can canh bao. Tot lam!");
    sheet.getRange(headerRow + 1, 1).setFontColor(DESIGN_COLORS.SUCCESS).setFontStyle("italic");
    return;
  }

  var rows = attentionTasks.map(function (t) {
    return [t.taskId, t.taskName, t.priority, t.deadlineStatus, t.riskLevel, t.healthStatus];
  });

  var dataRange = sheet.getRange(headerRow + 1, 1, rows.length, headers.length);
  dataRange.setValues(rows);

  for (var r = 0; r < rows.length; r++) {
    var riskCell = sheet.getRange(headerRow + 1 + r, 5);
    riskCell.setFontColor(RISK_COLORS[rows[r][4]]).setFontWeight("bold");
  }
}