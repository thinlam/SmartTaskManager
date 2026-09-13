/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 9: Notifications & Automation
 *
 * checkDeadlines() quet toan bo Task dang hoat dong, xac dinh Task nao can
 * nhac nho (Overdue / Due Today / Due Tomorrow / Critical) va gui email
 * qua sendTaskReminder_() NEU OwnerEmail ton tai va Setting tuong ung
 * duoc bat.
 *
 * Chong spam (muc LXV): moi (TaskId, NotificationType) chi gui TOI DA
 * 1 lan / ngay - kiem tra qua log trong sheet 11_Notifications truoc khi
 * gui, khong dua vao bien nho tam (vi Apps Script khong giu state giua
 * cac lan chay trigger).
 *
 * createTriggers() tao Time-driven Trigger chay checkDeadlines() moi ngay
 * luc 07:00 - idempotent, kiem tra trigger da ton tai truoc khi tao.
 *
 * Phu thuoc global: types.gs, constants.gs, config.gs, utils.gs, tasks.gs,
 * setup.gs (dung getSetting_, arraysEqual_)
 */

var NOTIFICATION_TYPE = Object.freeze({
  OVERDUE: "Overdue",
  DUE_TODAY: "DueToday",
  DUE_TOMORROW: "DueTomorrow",
  CRITICAL: "Critical"
});

var NOTIFICATIONS_HEADERS = Object.freeze([
  "Timestamp", "Task ID", "Notification Type", "Recipient", "Status", "Message"
]);

var NOTIFICATION_TRIGGER_HOUR = 7;

// ==================================================
// ENTRY POINT
// ==================================================

function checkDeadlines() {
  var notificationsEnabled = getSetting_("notificationsEnabled", DEFAULT_APP_SETTINGS.notificationsEnabled);
  if (notificationsEnabled === false || notificationsEnabled === "FALSE") {
    Logger.log("checkDeadlines() bo qua - notificationsEnabled = false.");
    return;
  }

  ensureNotificationsSheet_();

  var tasks = getAllTasks();
  var sentToday = loadSentNotificationsToday_();
  var sentCount = 0;
  var skippedCount = 0;

  tasks.forEach(function (task) {
    if (task.status === STATUS.COMPLETED || task.status === STATUS.CANCELLED) return;

    var candidates = getApplicableNotificationTypes_(task);
    candidates.forEach(function (type) {
      var dedupeKey = task.taskId + "|" + type;
      if (sentToday[dedupeKey]) {
        return;
      }

      if (!task.ownerEmail) {
        logNotification_(task.taskId, type, "", "Skipped", "Khong co OwnerEmail");
        skippedCount++;
        return;
      }

      try {
        sendTaskReminder_(task, type);
        logNotification_(task.taskId, type, task.ownerEmail, "Sent", buildNotificationMessage_(task, type));
        sentCount++;
      } catch (e) {
        logNotification_(task.taskId, type, task.ownerEmail, "Failed", String(e.message));
      }
    });
  });

  Logger.log("checkDeadlines() hoan tat. Da gui: " + sentCount + ", bo qua: " + skippedCount + ".");
}

// ==================================================
// XAC DINH LOAI NOTIFICATION AP DUNG CHO 1 TASK
// ==================================================

function getApplicableNotificationTypes_(task) {
  var types = [];

  var notifyOverdue = getSetting_("notificationEmailOnOverdue", DEFAULT_APP_SETTINGS.notificationEmailOnOverdue);
  var notifyDueToday = getSetting_("notificationEmailOnDueToday", DEFAULT_APP_SETTINGS.notificationEmailOnDueToday);
  var notifyDueTomorrow = getSetting_("notificationEmailOnDueTomorrow", DEFAULT_APP_SETTINGS.notificationEmailOnDueTomorrow);
  var notifyCritical = getSetting_("notificationEmailOnCritical", DEFAULT_APP_SETTINGS.notificationEmailOnCritical);

  if (notifyOverdue !== false && task.deadlineStatus === DEADLINE_STATUS.OVERDUE) {
    types.push(NOTIFICATION_TYPE.OVERDUE);
  }
  if (notifyDueToday !== false && task.deadlineStatus === DEADLINE_STATUS.DUE_TODAY) {
    types.push(NOTIFICATION_TYPE.DUE_TODAY);
  }
  if (notifyDueTomorrow === true && task.deadlineStatus === DEADLINE_STATUS.DUE_TOMORROW) {
    types.push(NOTIFICATION_TYPE.DUE_TOMORROW);
  }
  if (
    notifyCritical !== false &&
    task.priority === PRIORITY.CRITICAL &&
    Number(task.smartScore) >= 75 &&
    task.deadlineStatus !== DEADLINE_STATUS.OVERDUE &&
    task.deadlineStatus !== DEADLINE_STATUS.DUE_TODAY
  ) {
    types.push(NOTIFICATION_TYPE.CRITICAL);
  }

  return types;
}

// ==================================================
// GUI EMAIL
// ==================================================

function sendTaskReminder_(task, type) {
  var systemName = getSetting_("systemName", DEFAULT_APP_SETTINGS.systemName);
  var subject = "[" + systemName + "] " + notificationTypeLabel_(type) + ": " + task.taskName;
  var body = buildNotificationMessage_(task, type);

  MailApp.sendEmail({
    to: task.ownerEmail,
    subject: subject,
    body: body
  });
}

function notificationTypeLabel_(type) {
  switch (type) {
    case NOTIFICATION_TYPE.OVERDUE: return "Task qua han";
    case NOTIFICATION_TYPE.DUE_TODAY: return "Task den han hom nay";
    case NOTIFICATION_TYPE.DUE_TOMORROW: return "Task den han ngay mai";
    case NOTIFICATION_TYPE.CRITICAL: return "Task Critical can chu y";
    default: return "Nhac nho Task";
  }
}

function buildNotificationMessage_(task, type) {
  var lines = [
    notificationTypeLabel_(type) + ":",
    "",
    "Task: " + task.taskName + " (" + task.taskId + ")",
    "Priority: " + task.priority,
    "Status: " + task.status,
    "Due Date: " + (task.dueDate || "Khong co"),
    "Smart Score: " + task.smartScore + " (" + task.smartPriority + ")",
    "Recommended Action: " + task.recommendedAction
  ];
  return lines.join("\n");
}

// ==================================================
// CHONG SPAM - DOC LOG DA GUI HOM NAY
// ==================================================

function loadSentNotificationsToday_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  var result = {};
  if (!sheet) return result;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  var todayKey = formatDateOnly(getTodayDateOnly(new Date()));
  var data = sheet.getRange(2, 1, lastRow - 1, NOTIFICATIONS_HEADERS.length).getValues();

  data.forEach(function (row) {
    var timestamp = row[0];
    var taskId = row[1];
    var type = row[2];
    var status = row[4];
    if (!timestamp || status !== "Sent") return;

    var rowDateKey = normalizeTimestampToDateKey_(timestamp);
    if (rowDateKey === todayKey) {
      result[taskId + "|" + type] = true;
    }
  });

  return result;
}

function normalizeTimestampToDateKey_(value) {
  if (value instanceof Date) return formatDateOnly(value);
  var parsed = parseDateOnly(String(value));
  return parsed ? formatDateOnly(parsed) : "";
}

function logNotification_(taskId, type, recipient, status, message) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  if (!sheet) return;
  sheet.appendRow([formatDateTimeISO(new Date()), taskId, type, recipient, status, message]);
}

// ==================================================
// SETUP SHEET NOTIFICATIONS (idempotent, tu khoi tao khi can)
// ==================================================

function ensureNotificationsSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NOTIFICATIONS);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.NOTIFICATIONS + " chua duoc tao. Chay setupSmartTaskManager() truoc.");

  var headerRange = sheet.getRange(1, 1, 1, NOTIFICATIONS_HEADERS.length);
  var currentHeader = headerRange.getValues()[0];

  if (!arraysEqual_(currentHeader, NOTIFICATIONS_HEADERS)) {
    sheet.getRange(1, 1, 1, 6).clearContent();
    headerRange.setValues([NOTIFICATIONS_HEADERS]);
    headerRange
      .setFontWeight("bold")
      .setFontColor("#FFFFFF")
      .setBackground(DESIGN_COLORS.PRIMARY_DARK)
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
}

// ==================================================
// TRIGGER (chay checkDeadlines() moi ngay luc 07:00)
// ==================================================

function createTriggers() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  var alreadyExists = existingTriggers.some(function (t) {
    return t.getHandlerFunction() === "checkDeadlines";
  });

  if (alreadyExists) {
    Logger.log("Trigger checkDeadlines() da ton tai - khong tao trung.");
    return;
  }

  ScriptApp.newTrigger("checkDeadlines")
    .timeBased()
    .atHour(NOTIFICATION_TRIGGER_HOUR)
    .everyDays(1)
    .create();

  Logger.log("Da tao trigger checkDeadlines() chay hang ngay luc " + NOTIFICATION_TRIGGER_HOUR + ":00.");
}

function removeDeadlineTriggers_() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(function (t) {
    if (t.getHandlerFunction() === "checkDeadlines") {
      ScriptApp.deleteTrigger(t);
    }
  });
}