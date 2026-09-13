/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 5: Tasks (CRUD that tren sheet 02_Tasks)
 *
 * Day la layer duy nhat duoc phep doc/ghi truc tiep vao sheet Tasks.
 * Moi thao tac them/sua Task DEU phai di qua addTask()/updateTask() de
 * dam bao SmartScore/Risk/Deadline luon duoc tinh lai dung, khong bao gio
 * ghi thang gia tri "cu" vao cac cot computed.
 *
 * Phu thuoc global: types.gs, constants.gs, config.gs, utils.gs,
 * smart-engine.gs, setup.gs (dung TASK_COL, arraysEqual_ khong can thiet o day)
 */

// ==================================================
// CURRENT USER (safe - khong doi hoi scope userinfo.email)
// ==================================================

/**
 * Lay email nguoi dung hien tai neu duoc phep, fallback ve "system" neu
 * script chua duoc cap quyen "https://www.googleapis.com/auth/userinfo.email"
 * hoac dang chay boi trigger/service account khong co user context.
 * Khong dung Session.getActiveUser().getEmail() truc tiep o noi khac -
 * luon goi qua ham nay.
 * @returns {string}
 */
function getCurrentUserEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email || "system";
  } catch (e) {
    return "system";
  }
}

// ==================================================
// ID GENERATION
// ==================================================

/**
 * Sinh TaskId tiep theo dang TASK-000001, khong bao gio trung, khong phu
 * thuoc vi tri hang (an toan khi Sort/Filter) vi luon quet toan bo cot

 * TaskId hien co de lay so lon nhat + 1.
 * @returns {string}
 */
function generateTaskId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var lastRow = sheet.getLastRow();
  var maxNumber = 0;

  if (lastRow >= 2) {
    var ids = sheet.getRange(2, TASK_COL.taskId + 1, lastRow - 1, 1).getValues();
    var prefix = ID_CONFIG.TASK.prefix;
    for (var i = 0; i < ids.length; i++) {
      var idValue = String(ids[i][0] || "");
      if (idValue.indexOf(prefix) === 0) {
        var numberPart = parseInt(idValue.substring(prefix.length), 10);
        if (!isNaN(numberPart) && numberPart > maxNumber) {
          maxNumber = numberPart;
        }
      }
    }
  }

  var nextNumber = maxNumber + 1;
  var padded = String(nextNumber);
  while (padded.length < ID_CONFIG.TASK.padLength) {
    padded = "0" + padded;
  }
  return ID_CONFIG.TASK.prefix + padded;
}

// ==================================================
// READ HELPERS
// ==================================================

/**
 * Đọc toàn bộ Task hiện có thành mảng object (key = TASK_FIELDS).
 * Dùng batch getValues() 1 lần duy nhất - không loop từng cell.
 * @returns {Array<Object>}
 */
function getAllTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, TASK_FIELDS.length).getValues();
  var tasks = [];

  for (var r = 0; r < values.length; r++) {
    if (!values[r][TASK_COL.taskId]) continue; // bỏ qua hàng trống
    tasks.push(rowToTask_(values[r]));
  }
  return tasks;
}

/**
 * Tìm 1 Task theo taskId. Trả về null nếu không tồn tại.
 * @param {string} taskId
 * @returns {Object|null}
 */
function getTaskById(taskId) {
  var found = findTaskRow_(taskId);
  return found ? found.task : null;
}

/**
 * Tìm vị trí hàng (1-based, tính cả header) + dữ liệu Task theo taskId.
 * @param {string} taskId
 * @returns {{rowIndex:number, task:Object}|null}
 * @private
 */
function findTaskRow_(taskId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var values = sheet.getRange(2, 1, lastRow - 1, TASK_FIELDS.length).getValues();
  for (var r = 0; r < values.length; r++) {
    if (values[r][TASK_COL.taskId] === taskId) {
      return { rowIndex: r + 2, task: rowToTask_(values[r]) };
    }
  }
  return null;
}

/**
 * Chuyển 1 hàng values[] thành object Task theo đúng TASK_FIELDS.
 * @param {Array} row
 * @returns {Object}
 * @private
 */
function rowToTask_(row) {
  var task = {};
  for (var i = 0; i < TASK_FIELDS.length; i++) {
    task[TASK_FIELDS[i]] = row[i];
  }
  // Chuẩn hoá vài kiểu dữ liệu Google Sheets trả về không đúng ý muốn.
  task.progress = Number(task.progress) || 0;
  task.estimatedHours = Number(task.estimatedHours) || 0;
  task.actualHours = Number(task.actualHours) || 0;
  task.isOverdue = task.isOverdue === true || task.isOverdue === "TRUE";
  task.isBlocked = task.isBlocked === true || task.isBlocked === "TRUE";
  task.startDate = normalizeDateCell_(task.startDate);
  task.dueDate = normalizeDateCell_(task.dueDate);
  task.completedDate = normalizeDateCell_(task.completedDate);
  return task;
}

/**
 * Sheet có thể trả Date object (nếu ô được format Date) hoặc string.
 * Chuẩn hoá về "yyyy-MM-dd" hoặc null.
 * @param {Date|string|null} value
 * @returns {string|null}
 * @private
 */
function normalizeDateCell_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return formatDateOnly(value);
  }
  var parsed = parseDateOnly(String(value));
  return parsed ? formatDateOnly(parsed) : null;
}

// ==================================================
// DEPENDENCY RESOLUTION
// ==================================================

/**
 * true nếu Task không có dependency, hoặc dependency đã Completed.
 * false nếu dependency tồn tại và CHƯA Completed.
 * @param {string|null} dependencyTaskId
 * @returns {boolean}
 * @private
 */
function isDependencyResolved_(dependencyTaskId) {
  if (!dependencyTaskId) return true;
  var dependency = getTaskById(dependencyTaskId);
  if (!dependency) return true; // dependency không tồn tại (đã xoá) -> không chặn
  return dependency.status === STATUS.COMPLETED;
}

// ==================================================
// VALIDATION
// ==================================================

/**
 * @param {Object} input
 * @throws {Error} nếu input thiếu field bắt buộc hoặc sai giá trị enum
 * @private
 */
function validateTaskInput_(input) {
  if (!input || !input.taskName || String(input.taskName).trim() === "") {
    throw new Error("TaskName la bat buoc.");
  }
  if (!input.priority || PRIORITY_LIST.indexOf(input.priority) === -1) {
    throw new Error("Priority khong hop le: " + input.priority);
  }
  if (input.status && STATUS_LIST.indexOf(input.status) === -1) {
    throw new Error("Status khong hop le: " + input.status);
  }
  if (input.dependencyTaskId && !getTaskById(input.dependencyTaskId)) {
    throw new Error("DependencyTaskId khong ton tai: " + input.dependencyTaskId);
  }
  if (
    input.progress !== undefined &&
    input.progress !== null &&
    (isNaN(input.progress) || input.progress < 0 || input.progress > 100)
  ) {
    throw new Error("Progress phai trong khoang 0-100.");
  }
}

// ==================================================
// CREATE
// ==================================================

/**
 * Tạo Task mới. Tự động:
 * - sinh taskId
 * - set createdAt/updatedAt/lastStatusChangedAt
 * - tính deadlineStatus/smartScore/riskLevel/recommendedAction qua Smart Engine
 * - ghi Activity Log
 *
 * @param {Object} input - xem TaskInput trong types.gs
 * @returns {Object} Task vừa tạo (đầy đủ field, kể cả computed)
 */
function addTask(input) {
  validateTaskInput_(input);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var now = new Date();
  var nowISO = formatDateTimeISO(now);

  var status = input.status || getSetting_("defaultStatus", DEFAULT_APP_SETTINGS.defaultStatus);
  var dependencyTaskId = input.dependencyTaskId || null;
  var isBlocked = status === STATUS.BLOCKED;
  var isDependencyResolved = isDependencyResolved_(dependencyTaskId);

  var baseTask = {
    priority: input.priority,
    status: status,
    progress: Number(input.progress) || 0,
    isBlocked: isBlocked,
    dueDate: input.dueDate || null
  };

  var computed = computeTaskFields(baseTask, {
    referenceDate: now,
    isDependencyResolved: isDependencyResolved
  });

  var task = {
    taskId: generateTaskId(),
    taskName: String(input.taskName).trim(),
    description: input.description || "",

    projectId: input.projectId || "",
    projectName: input.projectName || "",
    category: input.category || "",
    tags: input.tags || "",

    ownerId: input.ownerId || "",
    ownerName: input.ownerName || "",
    ownerEmail: input.ownerEmail || "",
    createdBy: input.createdBy || getCurrentUserEmail_(),

    priority: input.priority,
    status: status,

    startDate: input.startDate || null,
    dueDate: input.dueDate || null,
    completedDate: status === STATUS.COMPLETED ? formatDateOnly(now) : null,

    progress: baseTask.progress,
    estimatedHours: Number(input.estimatedHours) || 0,
    actualHours: 0,

    daysRemaining: computed.daysRemaining,
    deadlineStatus: computed.deadlineStatus,

    smartScore: computed.smartScore,
    smartPriority: computed.smartPriority,

    riskLevel: computed.riskLevel,
    healthStatus: computed.healthStatus,

    isOverdue: computed.isOverdue,
    isBlocked: isBlocked,

    recommendedAction: computed.recommendedAction,

    dependencyTaskId: dependencyTaskId || "",
    recurringType: input.recurringType || RECURRING_TYPE.NONE,

    createdAt: nowISO,
    updatedAt: nowISO,
    lastStatusChangedAt: nowISO,

    notes: input.notes || ""
  };

  var row = TASK_FIELDS.map(function (field) {
    return task[field] === null || task[field] === undefined ? "" : task[field];
  });

  sheet.appendRow(row);

  writeActivityLog_(task.taskId, ACTIVITY_ACTION.CREATED, "", task.taskName);

  return task;
}

// ==================================================
// UPDATE
// ==================================================

/**
 * Cập nhật 1 Task theo taskId. Chỉ nhận các field cho phép sửa trực tiếp
 * (xem TaskUpdateInput trong types.gs) - các field computed luôn được
 * Smart Engine tính lại, không nhận giá trị patch cho chúng.
 *
 * @param {string} taskId
 * @param {Object} patch - xem TaskUpdateInput trong types.gs
 * @returns {Object} Task sau khi cập nhật
 */
function updateTask(taskId, patch) {
  var found = findTaskRow_(taskId);
  if (!found) throw new Error("Khong tim thay Task: " + taskId);

  var existing = found.task;
  var now = new Date();
  var nowISO = formatDateTimeISO(now);

  var statusChanged = patch.status && patch.status !== existing.status;
  var newStatus = patch.status || existing.status;

  if (patch.status && STATUS_LIST.indexOf(patch.status) === -1) {
    throw new Error("Status khong hop le: " + patch.status);
  }
  if (patch.priority && PRIORITY_LIST.indexOf(patch.priority) === -1) {
    throw new Error("Priority khong hop le: " + patch.priority);
  }
  if (
    patch.progress !== undefined &&
    patch.progress !== null &&
    (isNaN(patch.progress) || patch.progress < 0 || patch.progress > 100)
  ) {
    throw new Error("Progress phai trong khoang 0-100.");
  }
  if (patch.dependencyTaskId && !getTaskById(patch.dependencyTaskId)) {
    throw new Error("DependencyTaskId khong ton tai: " + patch.dependencyTaskId);
  }

  var dependencyTaskId =
    patch.dependencyTaskId !== undefined ? patch.dependencyTaskId : existing.dependencyTaskId;
  var isDependencyResolved = isDependencyResolved_(dependencyTaskId);
  var isBlocked = newStatus === STATUS.BLOCKED;

  var mergedForCompute = {
    priority: patch.priority || existing.priority,
    status: newStatus,
    progress: patch.progress !== undefined ? Number(patch.progress) : existing.progress,
    isBlocked: isBlocked,
    dueDate: patch.dueDate !== undefined ? patch.dueDate : existing.dueDate
  };

  var computed = computeTaskFields(mergedForCompute, {
    referenceDate: now,
    isDependencyResolved: isDependencyResolved
  });

  var updated = Object.assign({}, existing, patch, {
    priority: mergedForCompute.priority,
    status: newStatus,
    progress: mergedForCompute.progress,
    dueDate: mergedForCompute.dueDate,
    isBlocked: isBlocked,
    dependencyTaskId: dependencyTaskId || "",

    daysRemaining: computed.daysRemaining,
    deadlineStatus: computed.deadlineStatus,
    smartScore: computed.smartScore,
    smartPriority: computed.smartPriority,
    riskLevel: computed.riskLevel,
    healthStatus: computed.healthStatus,
    isOverdue: computed.isOverdue,
    recommendedAction: computed.recommendedAction,

    completedDate:
      newStatus === STATUS.COMPLETED
        ? existing.completedDate || formatDateOnly(now)
        : (statusChanged ? null : existing.completedDate),

    updatedAt: nowISO,
    lastStatusChangedAt: statusChanged ? nowISO : existing.lastStatusChangedAt
  });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var row = TASK_FIELDS.map(function (field) {
    return updated[field] === null || updated[field] === undefined ? "" : updated[field];
  });
  sheet.getRange(found.rowIndex, 1, 1, TASK_FIELDS.length).setValues([row]);

  if (statusChanged) {
    writeActivityLog_(taskId, ACTIVITY_ACTION.STATUS_CHANGED, existing.status, newStatus);
  }
  if (patch.progress !== undefined && Number(patch.progress) !== existing.progress) {
    writeActivityLog_(taskId, ACTIVITY_ACTION.PROGRESS_CHANGED, String(existing.progress), String(patch.progress));
  }
  if (patch.priority && patch.priority !== existing.priority) {
    writeActivityLog_(taskId, ACTIVITY_ACTION.PRIORITY_CHANGED, existing.priority, patch.priority);
  }
  if (patch.ownerId && patch.ownerId !== existing.ownerId) {
    writeActivityLog_(taskId, ACTIVITY_ACTION.OWNER_CHANGED, existing.ownerId, patch.ownerId);
  }
  if (patch.dueDate !== undefined && patch.dueDate !== existing.dueDate) {
    writeActivityLog_(taskId, ACTIVITY_ACTION.DEADLINE_CHANGED, String(existing.dueDate), String(patch.dueDate));
  }

  return updated;
}

// ==================================================
// ACTIVITY LOG WRITE
// ==================================================

/**
 * @param {string} taskId
 * @param {string} action - xem ACTIVITY_ACTION trong constants.gs
 * @param {string} oldValue
 * @param {string} newValue
 * @private
 */
function writeActivityLog_(taskId, action, oldValue, newValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ACTIVITY_LOG);
  if (!sheet) return; // chưa setup thì bỏ qua, không làm gãy addTask/updateTask

  var user = getCurrentUserEmail_();
  sheet.appendRow([formatDateTimeISO(new Date()), user, taskId, action, oldValue, newValue]);
}

// ==================================================
// RECOMPUTE ALL (dùng cho refresh thủ công / trigger hằng ngày ở Phase sau)
// ==================================================

/**
 * Tính lại field computed cho TOÀN BỘ Task hiện có, ghi lại 1 lần bằng
 * batch setValues() - không loop setValue() từng cell.
 * Dùng khi ngày đã sang hôm mới (Overdue/Due Today cần cập nhật) mà
 * người dùng không tự sửa gì trong Task.
 */
function recomputeAllTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TASKS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var now = new Date();
  var values = sheet.getRange(2, 1, lastRow - 1, TASK_FIELDS.length).getValues();
  var tasksById = {};
  var rows = values.map(rowToTask_);
  rows.forEach(function (t) {
    tasksById[t.taskId] = t;
  });

  var updatedRows = rows.map(function (task) {
    var isDependencyResolved = true;
    if (task.dependencyTaskId) {
      var dep = tasksById[task.dependencyTaskId];
      isDependencyResolved = !dep || dep.status === STATUS.COMPLETED;
    }

    var computed = computeTaskFields(
      {
        priority: task.priority,
        status: task.status,
        progress: task.progress,
        isBlocked: task.status === STATUS.BLOCKED,
        dueDate: task.dueDate
      },
      { referenceDate: now, isDependencyResolved: isDependencyResolved }
    );

    var merged = Object.assign({}, task, {
      daysRemaining: computed.daysRemaining,
      deadlineStatus: computed.deadlineStatus,
      smartScore: computed.smartScore,
      smartPriority: computed.smartPriority,
      riskLevel: computed.riskLevel,
      healthStatus: computed.healthStatus,
      isOverdue: computed.isOverdue,
      isBlocked: task.status === STATUS.BLOCKED,
      recommendedAction: computed.recommendedAction
    });

    return TASK_FIELDS.map(function (field) {
      return merged[field] === null || merged[field] === undefined ? "" : merged[field];
    });
  });

  sheet.getRange(2, 1, updatedRows.length, TASK_FIELDS.length).setValues(updatedRows);
  Logger.log("recomputeAllTasks() da cap nhat " + updatedRows.length + " task.");
}