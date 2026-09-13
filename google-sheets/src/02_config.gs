/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 1-3: Configuration (Smart Engine tuning parameters)
 *
 * Chỉ chứa CẤU HÌNH (số liệu, ngưỡng, trọng số). Logic tính toán thực sự
 * thuộc Phase 4 (smart-engine.gs).
 *
 * ĐỒNG BỘ VỚI: excel/config.ts
 * Mọi con số ở đây PHẢI khớp với bản TypeScript để SmartScore/Risk tính
 * ra kết quả giống nhau trên cả 2 platform.
 *
 * Phụ thuộc: constants.gs (phải load trước - Apps Script tự gộp file theo
 * thứ tự alphabet trong Editor, nhưng vì tất cả đều ở global scope nên thứ
 * tự chỉ quan trọng tại thời điểm gọi hàm, không phải thời điểm khai báo
 * "var" đơn thuần).
 */

// ==================================================
// SMART SCORE WEIGHTS
// ==================================================
// Tổng điểm tối đa = 100.
// SmartScore = PriorityScore + DeadlineScore + OverdueScore
//              + ProgressRiskScore + BlockedScore + DependencyScore
//              + ProjectWeightScore (đã clamp 0-100)

var SMART_SCORE_WEIGHTS = Object.freeze({
  PRIORITY_POINTS: Object.freeze((function () {
    var map = {};
    map[PRIORITY.CRITICAL] = 25;
    map[PRIORITY.URGENT] = 20;
    map[PRIORITY.HIGH] = 15;
    map[PRIORITY.MEDIUM] = 8;
    map[PRIORITY.LOW] = 3;
    return map;
  })()),
  PRIORITY_MAX: 25,

  DEADLINE_POINTS: Object.freeze({
    DUE_TODAY: 25,
    DUE_TOMORROW: 20,
    DUE_IN_3_DAYS: 15,
    DUE_THIS_WEEK: 10,
    UPCOMING: 5,
    NO_DEADLINE_OR_FAR: 0
  }),
  DEADLINE_MAX: 25,

  OVERDUE_POINTS_PER_DAY: 4,
  OVERDUE_MAX: 20,

  PROGRESS_RISK_POINTS: Object.freeze({
    BELOW_25: 15,
    BELOW_50: 10,
    BELOW_75: 5,
    BELOW_100: 0
  }),
  PROGRESS_RISK_MAX: 15,

  BLOCKED_POINTS: 10,
  BLOCKED_MAX: 10,

  DEPENDENCY_POINTS: 5,
  DEPENDENCY_MAX: 5,

  PROJECT_WEIGHT_POINTS: Object.freeze((function () {
    var map = {};
    map[PRIORITY.CRITICAL] = 10;
    map[PRIORITY.URGENT] = 7;
    map[PRIORITY.HIGH] = 5;
    map[PRIORITY.MEDIUM] = 2;
    map[PRIORITY.LOW] = 0;
    return map;
  })()),
  PROJECT_WEIGHT_MAX: 10,

  SCORE_MIN: 0,
  SCORE_MAX: 100
});

/** Ngưỡng phân loại SmartScore -> SmartPriorityBand. Duyệt từ trên xuống, dùng ">=". */
var SMART_SCORE_BANDS = Object.freeze([
  { min: 90, band: SMART_PRIORITY_BAND.CRITICAL_ATTENTION },
  { min: 75, band: SMART_PRIORITY_BAND.VERY_HIGH_PRIORITY },
  { min: 60, band: SMART_PRIORITY_BAND.HIGH_PRIORITY },
  { min: 40, band: SMART_PRIORITY_BAND.NORMAL },
  { min: 0, band: SMART_PRIORITY_BAND.LOW_ATTENTION }
]);

// ==================================================
// DEADLINE ENGINE CONFIG
// ==================================================

var DEADLINE_CONFIG = Object.freeze({
  dueSoonDays: 3,
  dueInDays: 3,
  dueThisWeekDays: 7,
  upcomingHorizonDays: 30
});

// ==================================================
// RISK ENGINE CONFIG
// ==================================================

var RISK_CONFIG = Object.freeze({
  WEIGHTS: Object.freeze({
    OVERDUE_POINTS_PER_DAY: 5,
    OVERDUE_MAX: 30,

    DEADLINE_PROXIMITY_POINTS: Object.freeze({
      DUE_TODAY: 20,
      DUE_TOMORROW: 15,
      DUE_IN_3_DAYS: 10,
      DUE_THIS_WEEK: 5,
      OTHER: 0
    }),
    DEADLINE_PROXIMITY_MAX: 20,

    PROGRESS_RISK_POINTS: Object.freeze({
      BELOW_25: 20,
      BELOW_50: 12,
      BELOW_75: 5,
      BELOW_100: 0
    }),
    PROGRESS_RISK_MAX: 20,

    PRIORITY_POINTS: Object.freeze((function () {
      var map = {};
      map[PRIORITY.CRITICAL] = 10;
      map[PRIORITY.URGENT] = 7;
      map[PRIORITY.HIGH] = 4;
      map[PRIORITY.MEDIUM] = 1;
      map[PRIORITY.LOW] = 0;
      return map;
    })()),
    PRIORITY_MAX: 10,

    BLOCKED_POINTS: 10,
    BLOCKED_MAX: 10,

    DEPENDENCY_POINTS: 5,
    DEPENDENCY_MAX: 5,

    AGING_MAX: 5,
    AGING_THRESHOLDS_DAYS: Object.freeze({
      LEVEL_1: 15,
      LEVEL_2: 30
    }),
    AGING_POINTS: Object.freeze({
      LEVEL_1: 2,
      LEVEL_2: 5
    })
  }),

  SCORE_MIN: 0,
  SCORE_MAX: 100
});

/** Ngưỡng phân loại Risk Score -> RiskLevel. Duyệt từ trên xuống, dùng ">=". */
var RISK_LEVEL_BANDS = Object.freeze([
  { min: 80, level: RISK_LEVEL.CRITICAL },
  { min: 55, level: RISK_LEVEL.HIGH },
  { min: 30, level: RISK_LEVEL.MEDIUM },
  { min: 0, level: RISK_LEVEL.LOW }
]);

// ==================================================
// TASK AGING BUCKETS
// ==================================================

var TASK_AGING_BUCKETS = Object.freeze([
  { label: "0-3 Days", minDays: 0, maxDays: 3 },
  { label: "4-7 Days", minDays: 4, maxDays: 7 },
  { label: "8-14 Days", minDays: 8, maxDays: 14 },
  { label: "15-30 Days", minDays: 15, maxDays: 30 },
  { label: "30+ Days", minDays: 31, maxDays: null }
]);

// ==================================================
// RECOMMENDATION ENGINE RULES
// ==================================================
// Danh sách rule được đánh giá THEO THỨ TỰ ưu tiên từ trên xuống. Rule đầu
// tiên khớp điều kiện sẽ quyết định RecommendedAction. Điều kiện thực tế
// được implement ở Phase 4 (smart-engine.gs) - ở đây chỉ khai báo thứ tự
// và action tương ứng để logic Phase 4 tra cứu.

var RECOMMENDATION_RULES = Object.freeze([
  {
    ruleId: "COMPLETED",
    action: RECOMMENDED_ACTION.COMPLETED,
    conditionDescription: "status === " + STATUS.COMPLETED
  },
  {
    ruleId: "BLOCKED_REVIEW",
    action: RECOMMENDED_ACTION.REVIEW_BLOCKED_TASK,
    conditionDescription: "isBlocked === true && status !== " + STATUS.CANCELLED
  },
  {
    ruleId: "WAITING_DEPENDENCY",
    action: RECOMMENDED_ACTION.WAITING_FOR_DEPENDENCY,
    conditionDescription:
      "dependencyTaskId tồn tại && dependency chưa Completed && chưa bị Blocked"
  },
  {
    ruleId: "OVERDUE_DO_NOW",
    action: RECOMMENDED_ACTION.DO_NOW,
    conditionDescription: "deadlineStatus === Overdue"
  },
  {
    ruleId: "DUE_TODAY",
    action: RECOMMENDED_ACTION.COMPLETE_TODAY,
    conditionDescription: "deadlineStatus === Due Today"
  },
  {
    ruleId: "DEADLINE_APPROACHING",
    action: RECOMMENDED_ACTION.DEADLINE_APPROACHING,
    conditionDescription:
      "deadlineStatus in [Due Tomorrow, Due in 3 Days, Due This Week]"
  },
  {
    ruleId: "HIGH_SMARTSCORE_FOLLOW_UP",
    action: RECOMMENDED_ACTION.FOLLOW_UP_WITH_OWNER,
    conditionDescription:
      "smartScore >= 75 && progress thấp bất thường so với thời gian đã trôi qua"
  },
  {
    ruleId: "LOW_PRIORITY",
    action: RECOMMENDED_ACTION.LOW_PRIORITY,
    conditionDescription: "smartScore < 40"
  },
  {
    ruleId: "DEFAULT_SCHEDULE",
    action: RECOMMENDED_ACTION.SCHEDULE_THIS_TASK,
    conditionDescription: "Không khớp rule nào ở trên (fallback mặc định)"
  }
]);

// ==================================================
// PRODUCTIVITY SCORE CONFIG
// ==================================================

var PRODUCTIVITY_CONFIG = Object.freeze({
  WEIGHTS: Object.freeze({
    COMPLETION_RATE_MAX: 35,
    ON_TIME_RATE_MAX: 30,
    OVERDUE_PENALTY_PER_TASK: 5,
    OVERDUE_PENALTY_MAX: 25,
    BLOCKED_PENALTY_PER_TASK: 3,
    BLOCKED_PENALTY_MAX: 15,
    BASE_SCORE: 50
  }),
  SCORE_MIN: 0,
  SCORE_MAX: 100
});

var PRODUCTIVITY_BANDS = Object.freeze([
  { min: 85, band: "Excellent" },
  { min: 70, band: "Good" },
  { min: 50, band: "Normal" },
  { min: 30, band: "Needs Attention" },
  { min: 0, band: "Critical" }
]);

// ==================================================
// WORKLOAD CONFIG
// ==================================================

var WORKLOAD_CONFIG = Object.freeze({
  DEFAULT_WEEKLY_CAPACITY_HOURS: 40,
  THRESHOLDS: Object.freeze({
    AVAILABLE_MAX_PERCENT: 50,
    NORMAL_MAX_PERCENT: 80,
    BUSY_MAX_PERCENT: 100
  })
});

// ==================================================
// DEFAULT APP SETTINGS
// ==================================================
// Giá trị mặc định khi setupSettingsSheet() chạy lần đầu (Phase 5).
// Sau khi user thay đổi trong sheet 12_Settings, đọc giá trị thật từ sheet,
// KHÔNG đọc lại từ đây.

/** @type {AppSettings} */
var DEFAULT_APP_SETTINGS = Object.freeze({
  systemName: "Smart Task Manager",
  defaultStatus: STATUS.TO_DO,
  defaultPriority: PRIORITY.MEDIUM,
  dueSoonDays: DEADLINE_CONFIG.dueSoonDays,
  workingDays: [1, 2, 3, 4, 5],
  workingHoursPerDay: 8,
  weekendDays: [6, 7],
  notificationsEnabled: true,
  notificationEmailOnOverdue: true,
  notificationEmailOnDueToday: true,
  notificationEmailOnDueTomorrow: false,
  notificationEmailOnCritical: true,
  theme: "Light",
  dateFormat: "dd/MM/yyyy",
  currency: "VND",
  language: "vi"
});
