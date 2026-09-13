/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 4: Smart Engine
 *
 * Chứa logic tính toán thuần, không đụng SpreadsheetApp. Layer setup/tasks.gs
 * (Phase 5) sẽ đọc dữ liệu sheet -> dựng object Task -> gọi computeTaskFields()
 * -> ghi kết quả trở lại sheet.
 *
 * ĐỒNG BỘ VỚI: excel/smart-engine.ts
 * Cùng công thức, cùng thứ tự rule, cùng ngưỡng số (lấy từ config.gs).
 *
 * Phụ thuộc global: constants.gs, config.gs, utils.gs (phải cùng nằm trong
 * project Apps Script - tất cả file .gs chia sẻ 1 global scope).
 */

// ==================================================
// DEADLINE ENGINE
// ==================================================

/**
 * @param {TaskStatus} status
 * @param {string|null} dueDate
 * @param {Date} [referenceDate]
 * @returns {DeadlineStatus}
 */
function calculateDeadlineStatus(status, dueDate, referenceDate) {
  referenceDate = referenceDate || new Date();

  if (CLOSED_STATUSES.indexOf(status) !== -1) {
    return status === STATUS.COMPLETED ? DEADLINE_STATUS.COMPLETED : DEADLINE_STATUS.NO_DEADLINE;
  }

  var due = parseDateOnly(dueDate);
  if (!due) return DEADLINE_STATUS.NO_DEADLINE;

  var daysRemaining = diffInCalendarDays(getTodayDateOnly(referenceDate), due);

  if (daysRemaining < 0) return DEADLINE_STATUS.OVERDUE;
  if (daysRemaining === 0) return DEADLINE_STATUS.DUE_TODAY;
  if (daysRemaining === 1) return DEADLINE_STATUS.DUE_TOMORROW;
  if (daysRemaining <= DEADLINE_CONFIG.dueInDays) return DEADLINE_STATUS.DUE_IN_3_DAYS;
  if (daysRemaining <= DEADLINE_CONFIG.dueThisWeekDays) return DEADLINE_STATUS.DUE_THIS_WEEK;
  return DEADLINE_STATUS.UPCOMING;
}

/**
 * @param {TaskStatus} status
 * @param {string|null} dueDate
 * @param {Date} [referenceDate]
 * @returns {number|null}
 */
function calculateDaysRemaining(status, dueDate, referenceDate) {
  referenceDate = referenceDate || new Date();
  if (CLOSED_STATUSES.indexOf(status) !== -1) return null;
  var due = parseDateOnly(dueDate);
  if (!due) return null;
  return diffInCalendarDays(getTodayDateOnly(referenceDate), due);
}

// ==================================================
// SMART SCORE ENGINE
// ==================================================

function priorityScoreOf_(priority) {
  return SMART_SCORE_WEIGHTS.PRIORITY_POINTS[priority] || 0;
}

function deadlineScoreOf_(deadlineStatus) {
  var points = SMART_SCORE_WEIGHTS.DEADLINE_POINTS;
  switch (deadlineStatus) {
    case DEADLINE_STATUS.DUE_TODAY:
      return points.DUE_TODAY;
    case DEADLINE_STATUS.DUE_TOMORROW:
      return points.DUE_TOMORROW;
    case DEADLINE_STATUS.DUE_IN_3_DAYS:
      return points.DUE_IN_3_DAYS;
    case DEADLINE_STATUS.DUE_THIS_WEEK:
      return points.DUE_THIS_WEEK;
    case DEADLINE_STATUS.UPCOMING:
      return points.UPCOMING;
    default:
      return points.NO_DEADLINE_OR_FAR;
  }
}

function overdueScoreOf_(daysRemaining) {
  if (daysRemaining === null || daysRemaining >= 0) return 0;
  var daysOverdue = Math.abs(daysRemaining);
  return clamp(daysOverdue * SMART_SCORE_WEIGHTS.OVERDUE_POINTS_PER_DAY, 0, SMART_SCORE_WEIGHTS.OVERDUE_MAX);
}

function progressRiskScoreOf_(progress, status) {
  if (CLOSED_STATUSES.indexOf(status) !== -1) return 0;
  var points = SMART_SCORE_WEIGHTS.PROGRESS_RISK_POINTS;
  if (progress < 25) return points.BELOW_25;
  if (progress < 50) return points.BELOW_50;
  if (progress < 75) return points.BELOW_75;
  return points.BELOW_100;
}

function projectWeightScoreOf_(projectPriority) {
  if (!projectPriority) return 0;
  return SMART_SCORE_WEIGHTS.PROJECT_WEIGHT_POINTS[projectPriority] || 0;
}

function bandOfSmartScore_(score) {
  for (var i = 0; i < SMART_SCORE_BANDS.length; i++) {
    if (score >= SMART_SCORE_BANDS[i].min) return SMART_SCORE_BANDS[i].band;
  }
  return SMART_SCORE_BANDS[SMART_SCORE_BANDS.length - 1].band;
}

/**
 * @param {{priority:TaskPriority, status:TaskStatus, progress:number, isBlocked:boolean}} task
 * @param {DeadlineStatus} deadlineStatus
 * @param {number|null} daysRemaining
 * @param {Object} [ctx] - { projectPriority, isDependencyResolved }
 * @returns {SmartScoreBreakdown}
 */
function calculateSmartScore(task, deadlineStatus, daysRemaining, ctx) {
  ctx = ctx || {};

  if (CLOSED_STATUSES.indexOf(task.status) !== -1) {
    return {
      priorityScore: 0,
      deadlineScore: 0,
      overdueScore: 0,
      progressRiskScore: 0,
      blockedScore: 0,
      dependencyScore: 0,
      projectWeightScore: 0,
      totalScore: 0,
      band: bandOfSmartScore_(0)
    };
  }

  var priorityScore = priorityScoreOf_(task.priority);
  var overdueScore = overdueScoreOf_(daysRemaining);
  var deadlineScore = overdueScore > 0 ? 0 : deadlineScoreOf_(deadlineStatus);
  var progressRiskScore = progressRiskScoreOf_(task.progress, task.status);
  var blockedScore = task.isBlocked ? SMART_SCORE_WEIGHTS.BLOCKED_POINTS : 0;
  var dependencyScore = ctx.isDependencyResolved === false ? SMART_SCORE_WEIGHTS.DEPENDENCY_POINTS : 0;
  var projectWeightScore = projectWeightScoreOf_(ctx.projectPriority);

  var rawTotal =
    priorityScore + deadlineScore + overdueScore + progressRiskScore + blockedScore + dependencyScore + projectWeightScore;

  var totalScore = roundSafe(clamp(rawTotal, SMART_SCORE_WEIGHTS.SCORE_MIN, SMART_SCORE_WEIGHTS.SCORE_MAX));

  return {
    priorityScore: priorityScore,
    deadlineScore: deadlineScore,
    overdueScore: overdueScore,
    progressRiskScore: progressRiskScore,
    blockedScore: blockedScore,
    dependencyScore: dependencyScore,
    projectWeightScore: projectWeightScore,
    totalScore: totalScore,
    band: bandOfSmartScore_(totalScore)
  };
}

// ==================================================
// RISK ENGINE
// ==================================================

function overdueRiskOf_(daysRemaining) {
  if (daysRemaining === null || daysRemaining >= 0) return 0;
  var daysOverdue = Math.abs(daysRemaining);
  return clamp(daysOverdue * RISK_CONFIG.WEIGHTS.OVERDUE_POINTS_PER_DAY, 0, RISK_CONFIG.WEIGHTS.OVERDUE_MAX);
}

function deadlineProximityRiskOf_(deadlineStatus, isOverdue) {
  if (isOverdue) return 0;
  var points = RISK_CONFIG.WEIGHTS.DEADLINE_PROXIMITY_POINTS;
  switch (deadlineStatus) {
    case DEADLINE_STATUS.DUE_TODAY:
      return points.DUE_TODAY;
    case DEADLINE_STATUS.DUE_TOMORROW:
      return points.DUE_TOMORROW;
    case DEADLINE_STATUS.DUE_IN_3_DAYS:
      return points.DUE_IN_3_DAYS;
    case DEADLINE_STATUS.DUE_THIS_WEEK:
      return points.DUE_THIS_WEEK;
    default:
      return points.OTHER;
  }
}

function progressRiskOf_(progress, status) {
  if (CLOSED_STATUSES.indexOf(status) !== -1) return 0;
  var points = RISK_CONFIG.WEIGHTS.PROGRESS_RISK_POINTS;
  if (progress < 25) return points.BELOW_25;
  if (progress < 50) return points.BELOW_50;
  if (progress < 75) return points.BELOW_75;
  return points.BELOW_100;
}

function priorityRiskOf_(priority) {
  return RISK_CONFIG.WEIGHTS.PRIORITY_POINTS[priority] || 0;
}

function agingRiskOf_(taskAgeDays) {
  if (!taskAgeDays) return 0;
  var thresholds = RISK_CONFIG.WEIGHTS.AGING_THRESHOLDS_DAYS;
  var points = RISK_CONFIG.WEIGHTS.AGING_POINTS;
  if (taskAgeDays >= thresholds.LEVEL_2) return points.LEVEL_2;
  if (taskAgeDays >= thresholds.LEVEL_1) return points.LEVEL_1;
  return 0;
}

function riskLevelOf_(score) {
  for (var i = 0; i < RISK_LEVEL_BANDS.length; i++) {
    if (score >= RISK_LEVEL_BANDS[i].min) return RISK_LEVEL_BANDS[i].level;
  }
  return RISK_LEVEL_BANDS[RISK_LEVEL_BANDS.length - 1].level;
}

/**
 * @param {{priority:TaskPriority, status:TaskStatus, progress:number, isBlocked:boolean}} task
 * @param {DeadlineStatus} deadlineStatus
 * @param {number|null} daysRemaining
 * @param {Object} [ctx] - { isDependencyResolved, taskAgeDays }
 * @returns {RiskBreakdown}
 */
function calculateRisk(task, deadlineStatus, daysRemaining, ctx) {
  ctx = ctx || {};

  if (CLOSED_STATUSES.indexOf(task.status) !== -1) {
    return {
      overdueRisk: 0,
      deadlineProximityRisk: 0,
      progressRisk: 0,
      priorityRisk: 0,
      blockedRisk: 0,
      dependencyRisk: 0,
      agingRisk: 0,
      totalRiskScore: 0,
      level: riskLevelOf_(0)
    };
  }

  var isOverdue = deadlineStatus === DEADLINE_STATUS.OVERDUE;
  var overdueRisk = overdueRiskOf_(daysRemaining);
  var deadlineProximityRisk = deadlineProximityRiskOf_(deadlineStatus, isOverdue);
  var progressRisk = progressRiskOf_(task.progress, task.status);
  var priorityRisk = priorityRiskOf_(task.priority);
  var blockedRisk = task.isBlocked ? RISK_CONFIG.WEIGHTS.BLOCKED_POINTS : 0;
  var dependencyRisk = ctx.isDependencyResolved === false ? RISK_CONFIG.WEIGHTS.DEPENDENCY_POINTS : 0;
  var agingRisk = agingRiskOf_(ctx.taskAgeDays);

  var rawTotal =
    overdueRisk + deadlineProximityRisk + progressRisk + priorityRisk + blockedRisk + dependencyRisk + agingRisk;

  var totalRiskScore = roundSafe(clamp(rawTotal, RISK_CONFIG.SCORE_MIN, RISK_CONFIG.SCORE_MAX));

  return {
    overdueRisk: overdueRisk,
    deadlineProximityRisk: deadlineProximityRisk,
    progressRisk: progressRisk,
    priorityRisk: priorityRisk,
    blockedRisk: blockedRisk,
    dependencyRisk: dependencyRisk,
    agingRisk: agingRisk,
    totalRiskScore: totalRiskScore,
    level: riskLevelOf_(totalRiskScore)
  };
}

// ==================================================
// RECOMMENDATION ENGINE
// ==================================================

/**
 * @param {{status:TaskStatus, isBlocked:boolean}} task
 * @param {DeadlineStatus} deadlineStatus
 * @param {number} smartScore
 * @param {boolean} isDependencyResolved
 * @returns {RecommendedAction}
 */
function calculateRecommendedAction(task, deadlineStatus, smartScore, isDependencyResolved) {
  if (task.status === STATUS.COMPLETED) {
    return RECOMMENDED_ACTION.COMPLETED;
  }
  if (task.isBlocked && task.status !== STATUS.CANCELLED) {
    return RECOMMENDED_ACTION.REVIEW_BLOCKED_TASK;
  }
  if (!isDependencyResolved) {
    return RECOMMENDED_ACTION.WAITING_FOR_DEPENDENCY;
  }
  if (deadlineStatus === DEADLINE_STATUS.OVERDUE) {
    return RECOMMENDED_ACTION.DO_NOW;
  }
  if (deadlineStatus === DEADLINE_STATUS.DUE_TODAY) {
    return RECOMMENDED_ACTION.COMPLETE_TODAY;
  }
  if (
    deadlineStatus === DEADLINE_STATUS.DUE_TOMORROW ||
    deadlineStatus === DEADLINE_STATUS.DUE_IN_3_DAYS ||
    deadlineStatus === DEADLINE_STATUS.DUE_THIS_WEEK
  ) {
    return RECOMMENDED_ACTION.DEADLINE_APPROACHING;
  }
  if (smartScore >= 75) {
    return RECOMMENDED_ACTION.FOLLOW_UP_WITH_OWNER;
  }
  if (smartScore < 40) {
    return RECOMMENDED_ACTION.LOW_PRIORITY;
  }
  return RECOMMENDED_ACTION.SCHEDULE_THIS_TASK;
}

// ==================================================
// HEALTH STATUS
// ==================================================

/**
 * @param {{status:TaskStatus, isBlocked:boolean}} task
 * @param {DeadlineStatus} deadlineStatus
 * @param {boolean} isDependencyResolved
 * @param {RiskLevel} riskLevel
 * @returns {string}
 */
function calculateHealthStatus(task, deadlineStatus, isDependencyResolved, riskLevel) {
  if (task.status === STATUS.COMPLETED) return "Completed";
  if (task.status === STATUS.CANCELLED) return "Cancelled";
  if (task.isBlocked) return "Blocked";
  if (!isDependencyResolved) return "Waiting Dependency";
  if (deadlineStatus === DEADLINE_STATUS.OVERDUE) return "Overdue";
  if (riskLevel === RISK_LEVEL.CRITICAL || riskLevel === RISK_LEVEL.HIGH) return "At Risk";
  return "On Track";
}

// ==================================================
// ORCHESTRATION
// ==================================================

/**
 * Hàm tổng hợp - dùng bởi tasks.gs (Phase 5) khi addTask()/updateTask()/
 * refreshDashboard() cần tính lại toàn bộ field computed cho 1 Task.
 *
 * @param {{priority:TaskPriority, status:TaskStatus, progress:number, isBlocked:boolean, dueDate:string|null}} task
 * @param {Object} [ctx] - { referenceDate, projectPriority, isDependencyResolved, taskAgeDays }
 * @returns {{
 *   daysRemaining:number|null,
 *   deadlineStatus:DeadlineStatus,
 *   smartScore:number,
 *   smartScoreBreakdown:SmartScoreBreakdown,
 *   smartPriority:SmartPriorityBand,
 *   riskLevel:RiskLevel,
 *   riskBreakdown:RiskBreakdown,
 *   healthStatus:string,
 *   isOverdue:boolean,
 *   recommendedAction:RecommendedAction
 * }}
 */
function computeTaskFields(task, ctx) {
  ctx = ctx || {};
  var referenceDate = ctx.referenceDate || new Date();
  var isDependencyResolved = ctx.isDependencyResolved !== false;

  var deadlineStatus = calculateDeadlineStatus(task.status, task.dueDate, referenceDate);
  var daysRemaining = calculateDaysRemaining(task.status, task.dueDate, referenceDate);
  var isOverdue = deadlineStatus === DEADLINE_STATUS.OVERDUE;

  var smartScoreBreakdown = calculateSmartScore(task, deadlineStatus, daysRemaining, ctx);
  var riskBreakdown = calculateRisk(task, deadlineStatus, daysRemaining, ctx);
  var recommendedAction = calculateRecommendedAction(
    task,
    deadlineStatus,
    smartScoreBreakdown.totalScore,
    isDependencyResolved
  );
  var healthStatus = calculateHealthStatus(task, deadlineStatus, isDependencyResolved, riskBreakdown.level);

  return {
    daysRemaining: daysRemaining,
    deadlineStatus: deadlineStatus,
    smartScore: smartScoreBreakdown.totalScore,
    smartScoreBreakdown: smartScoreBreakdown,
    smartPriority: smartScoreBreakdown.band,
    riskLevel: riskBreakdown.level,
    riskBreakdown: riskBreakdown,
    healthStatus: healthStatus,
    isOverdue: isOverdue,
    recommendedAction: recommendedAction
  };
}