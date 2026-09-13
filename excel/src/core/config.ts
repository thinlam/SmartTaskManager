/**
 * SMART TASK MANAGER - EXCEL / OFFICE SCRIPTS
 * Phase 1-3: Configuration (Smart Engine tuning parameters)
 *
 * File này chỉ chứa CẤU HÌNH (số liệu, ngưỡng, trọng số).
 * KHÔNG chứa logic tính toán - logic thực sự thuộc Phase 4 (smart-engine.ts).
 *
 * Mục đích tách riêng: cho phép chỉnh trọng số / ngưỡng qua Settings sheet
 * trong tương lai mà không phải sửa code logic.
 *
 * ĐỒNG BỘ VỚI: google-sheets/config.gs
 * Mọi con số ở đây PHẢI khớp với bản Google Apps Script để SmartScore/Risk
 * tính ra kết quả giống nhau trên cả 2 platform.
 */

import type { AppSettings, RecommendedAction, SmartPriorityBand } from "./types";
import {
    PRIORITY,
    RECOMMENDED_ACTION,
    RISK_LEVEL,
    SMART_PRIORITY_BAND,
    STATUS,
} from "./constants";

// ==================================================
// SMART SCORE WEIGHTS
// ==================================================
// Tổng điểm tối đa (maxPoints cộng lại) = 100.
// SmartScore = PriorityScore + DeadlineScore + OverdueScore
//              + ProgressRiskScore + BlockedScore + DependencyScore
//              + ProjectWeightScore  (đã clamp về 0-100)

export const SMART_SCORE_WEIGHTS = Object.freeze({
    // Điểm theo Priority - tối đa 25
    PRIORITY_POINTS: Object.freeze({
        [PRIORITY.CRITICAL]: 25,
        [PRIORITY.URGENT]: 20,
        [PRIORITY.HIGH]: 15,
        [PRIORITY.MEDIUM]: 8,
        [PRIORITY.LOW]: 3,
    }),
    PRIORITY_MAX: 25,

    // Điểm theo khoảng cách deadline (chỉ áp dụng khi CHƯA overdue) - tối đa 25
    DEADLINE_POINTS: Object.freeze({
        DUE_TODAY: 25,
        DUE_TOMORROW: 20,
        DUE_IN_3_DAYS: 15,
        DUE_THIS_WEEK: 10,
        UPCOMING: 5,
        NO_DEADLINE_OR_FAR: 0,
    }),
    DEADLINE_MAX: 25,

    // Điểm theo số ngày trễ hạn (overdue) - tối đa 20
    // Công thức: min(daysOverdue * OVERDUE_POINTS_PER_DAY, OVERDUE_MAX)
    OVERDUE_POINTS_PER_DAY: 4,
    OVERDUE_MAX: 20,

    // Điểm rủi ro do Progress thấp so với kỳ vọng - tối đa 15
    PROGRESS_RISK_POINTS: Object.freeze({
        BELOW_25: 15,
        BELOW_50: 10,
        BELOW_75: 5,
        BELOW_100: 0,
    }),
    PROGRESS_RISK_MAX: 15,

    // Điểm nếu Task đang Blocked - tối đa 10
    BLOCKED_POINTS: 10,
    BLOCKED_MAX: 10,

    // Điểm nếu có Dependency chưa hoàn thành - tối đa 5
    DEPENDENCY_POINTS: 5,
    DEPENDENCY_MAX: 5,

    // Trọng số ưu tiên dự án (Project Importance) - cộng thêm tối đa theo Priority của Project
    // Giữ ở mức nhỏ vì đây là yếu tố phụ, không lấn át yếu tố của chính Task.
    PROJECT_WEIGHT_POINTS: Object.freeze({
        [PRIORITY.CRITICAL]: 10,
        [PRIORITY.URGENT]: 7,
        [PRIORITY.HIGH]: 5,
        [PRIORITY.MEDIUM]: 2,
        [PRIORITY.LOW]: 0,
    }),
    PROJECT_WEIGHT_MAX: 10,

    // Giới hạn tổng điểm cuối cùng
    SCORE_MIN: 0,
    SCORE_MAX: 100,
} as const);

/**
 * Ngưỡng phân loại SmartScore -> SmartPriorityBand.
 * Duyệt từ trên xuống, dùng ngưỡng "min" (>=).
 */
export const SMART_SCORE_BANDS: ReadonlyArray<{
    min: number;
    band: SmartPriorityBand;
}> = Object.freeze([
    { min: 90, band: SMART_PRIORITY_BAND.CRITICAL_ATTENTION },
    { min: 75, band: SMART_PRIORITY_BAND.VERY_HIGH_PRIORITY },
    { min: 60, band: SMART_PRIORITY_BAND.HIGH_PRIORITY },
    { min: 40, band: SMART_PRIORITY_BAND.NORMAL },
    { min: 0, band: SMART_PRIORITY_BAND.LOW_ATTENTION },
]);

// ==================================================
// DEADLINE ENGINE CONFIG
// ==================================================

export const DEADLINE_CONFIG = Object.freeze({
    /** Số ngày được coi là "Due Soon" trên KPI Card DUE SOON (không tính Due Today/Tomorrow). */
    dueSoonDays: 3,
    /** Ngưỡng "Due in 3 Days" (số ngày remaining, không tính hôm nay/ngày mai). */
    dueInDays: 3,
    /** Số ngày tối đa để coi là "Due This Week". */
    dueThisWeekDays: 7,
    /** Số ngày tối đa để coi là "Upcoming" trước khi rơi vào nhóm xa (vẫn hiển thị Upcoming). */
    upcomingHorizonDays: 30,
} as const);

// ==================================================
// RISK ENGINE CONFIG
// ==================================================
// Risk được tính độc lập với SmartScore (mục đích khác nhau), nhưng dùng
// chung một số yếu tố đầu vào. Thang điểm riêng 0-100 rồi map ra RiskLevel.

export const RISK_CONFIG = Object.freeze({
    WEIGHTS: Object.freeze({
        // Overdue càng lâu, risk càng cao - tối đa 30
        OVERDUE_POINTS_PER_DAY: 5,
        OVERDUE_MAX: 30,

        // Deadline càng gần, risk càng cao (khi chưa overdue) - tối đa 20
        DEADLINE_PROXIMITY_POINTS: Object.freeze({
            DUE_TODAY: 20,
            DUE_TOMORROW: 15,
            DUE_IN_3_DAYS: 10,
            DUE_THIS_WEEK: 5,
            OTHER: 0,
        }),
        DEADLINE_PROXIMITY_MAX: 20,

        // Progress thấp trong khi thời gian còn ít - tối đa 20
        PROGRESS_RISK_POINTS: Object.freeze({
            BELOW_25: 20,
            BELOW_50: 12,
            BELOW_75: 5,
            BELOW_100: 0,
        }),
        PROGRESS_RISK_MAX: 20,

        // Priority cao góp phần tăng risk - tối đa 10
        PRIORITY_POINTS: Object.freeze({
            [PRIORITY.CRITICAL]: 10,
            [PRIORITY.URGENT]: 7,
            [PRIORITY.HIGH]: 4,
            [PRIORITY.MEDIUM]: 1,
            [PRIORITY.LOW]: 0,
        }),
        PRIORITY_MAX: 10,

        // Task đang Blocked - tối đa 10
        BLOCKED_POINTS: 10,
        BLOCKED_MAX: 10,

        // Dependency chưa hoàn thành khi deadline gần - tối đa 5
        DEPENDENCY_POINTS: 5,
        DEPENDENCY_MAX: 5,

        // Task tồn tại quá lâu mà chưa xong (Task Aging) - tối đa 5
        AGING_MAX: 5,
        AGING_THRESHOLDS_DAYS: Object.freeze({
            LEVEL_1: 15, // >=15 ngày tuổi -> +2
            LEVEL_2: 30, // >=30 ngày tuổi -> +5
        }),
        AGING_POINTS: Object.freeze({
            LEVEL_1: 2,
            LEVEL_2: 5,
        }),
    }),

    SCORE_MIN: 0,
    SCORE_MAX: 100,
} as const);

/** Ngưỡng phân loại Risk Score -> RiskLevel. Duyệt từ trên xuống, dùng ">=". */
export const RISK_LEVEL_BANDS = Object.freeze([
    { min: 80, level: RISK_LEVEL.CRITICAL },
    { min: 55, level: RISK_LEVEL.HIGH },
    { min: 30, level: RISK_LEVEL.MEDIUM },
    { min: 0, level: RISK_LEVEL.LOW },
] as const);

// ==================================================
// TASK AGING BUCKETS (dùng chung cho Dashboard/Reports/Risk)
// ==================================================

export const TASK_AGING_BUCKETS = Object.freeze([
    { label: "0-3 Days", minDays: 0, maxDays: 3 },
    { label: "4-7 Days", minDays: 4, maxDays: 7 },
    { label: "8-14 Days", minDays: 8, maxDays: 14 },
    { label: "15-30 Days", minDays: 15, maxDays: 30 },
    { label: "30+ Days", minDays: 31, maxDays: null },
] as const);

// ==================================================
// RECOMMENDATION ENGINE RULES
// ==================================================
// Danh sách rule được đánh giá THEO THỨ TỰ ưu tiên từ trên xuống.
// Rule đầu tiên khớp điều kiện sẽ quyết định RecommendedAction.
// Điều kiện thực tế (so sánh status/deadline/risk...) được implement ở
// Phase 4 (smart-engine.ts) - ở đây chỉ khai báo THỨ TỰ và ACTION tương ứng
// để logic Phase 4 tra cứu, tránh hard-code chuỗi rải rác.

export interface RecommendationRuleDef {
    /** Tên rule để log/debug, không phải giá trị hiển thị. */
    ruleId: string;
    action: RecommendedAction;
    /** Mô tả điều kiện áp dụng (tài liệu hoá, Phase 4 sẽ implement đúng theo mô tả này). */
    conditionDescription: string;
}

export const RECOMMENDATION_RULES: ReadonlyArray<RecommendationRuleDef> = Object.freeze([
    {
        ruleId: "COMPLETED",
        action: RECOMMENDED_ACTION.COMPLETED,
        conditionDescription: `status === ${STATUS.COMPLETED}`,
    },
    {
        ruleId: "BLOCKED_REVIEW",
        action: RECOMMENDED_ACTION.REVIEW_BLOCKED_TASK,
        conditionDescription: `isBlocked === true && status !== ${STATUS.CANCELLED}`,
    },
    {
        ruleId: "WAITING_DEPENDENCY",
        action: RECOMMENDED_ACTION.WAITING_FOR_DEPENDENCY,
        conditionDescription:
            "dependencyTaskId tồn tại && dependency chưa Completed && chưa bị Blocked",
    },
    {
        ruleId: "OVERDUE_DO_NOW",
        action: RECOMMENDED_ACTION.DO_NOW,
        conditionDescription: "deadlineStatus === Overdue",
    },
    {
        ruleId: "DUE_TODAY",
        action: RECOMMENDED_ACTION.COMPLETE_TODAY,
        conditionDescription: "deadlineStatus === Due Today",
    },
    {
        ruleId: "DEADLINE_APPROACHING",
        action: RECOMMENDED_ACTION.DEADLINE_APPROACHING,
        conditionDescription:
            "deadlineStatus in [Due Tomorrow, Due in 3 Days, Due This Week]",
    },
    {
        ruleId: "HIGH_SMARTSCORE_FOLLOW_UP",
        action: RECOMMENDED_ACTION.FOLLOW_UP_WITH_OWNER,
        conditionDescription:
            "smartScore >= 75 && progress thấp bất thường so với thời gian đã trôi qua",
    },
    {
        ruleId: "LOW_PRIORITY",
        action: RECOMMENDED_ACTION.LOW_PRIORITY,
        conditionDescription: "smartScore < 40",
    },
    {
        ruleId: "DEFAULT_SCHEDULE",
        action: RECOMMENDED_ACTION.SCHEDULE_THIS_TASK,
        conditionDescription: "Không khớp rule nào ở trên (fallback mặc định)",
    },
]);

// ==================================================
// PRODUCTIVITY SCORE CONFIG
// ==================================================

export const PRODUCTIVITY_CONFIG = Object.freeze({
    WEIGHTS: Object.freeze({
        COMPLETION_RATE_MAX: 35, // % completed / total, quy đổi tối đa 35 điểm
        ON_TIME_RATE_MAX: 30, // % completed đúng hạn / completed, tối đa 30 điểm
        OVERDUE_PENALTY_PER_TASK: 5, // trừ 5 điểm / task overdue, không cho âm quá mức
        OVERDUE_PENALTY_MAX: 25,
        BLOCKED_PENALTY_PER_TASK: 3,
        BLOCKED_PENALTY_MAX: 15,
        BASE_SCORE: 50, // điểm nền trước khi cộng/trừ theo completion & on-time
    }),
    SCORE_MIN: 0,
    SCORE_MAX: 100,
} as const);

export const PRODUCTIVITY_BANDS = Object.freeze([
    { min: 85, band: "Excellent" },
    { min: 70, band: "Good" },
    { min: 50, band: "Normal" },
    { min: 30, band: "Needs Attention" },
    { min: 0, band: "Critical" },
] as const);

// ==================================================
// WORKLOAD CONFIG
// ==================================================

export const WORKLOAD_CONFIG = Object.freeze({
    /** Số giờ ước tính mặc định/tuần dùng làm mẫu số nếu Member.workloadLimit = 0. */
    DEFAULT_WEEKLY_CAPACITY_HOURS: 40,
    THRESHOLDS: Object.freeze({
        AVAILABLE_MAX_PERCENT: 50, // <=50% capacity đang dùng
        NORMAL_MAX_PERCENT: 80, // 50-80%
        BUSY_MAX_PERCENT: 100, // 80-100%
        // >100% => Overloaded
    }),
} as const);

// ==================================================
// DEFAULT APP SETTINGS
// ==================================================
// Giá trị mặc định khi khởi tạo sheet 12_Settings lần đầu (Phase 5/6).
// Sau khi user thay đổi trong Settings, giá trị thực tế đọc từ sheet,
// KHÔNG đọc lại từ đây (đây chỉ là default cho lần setup đầu tiên).

export const DEFAULT_APP_SETTINGS: AppSettings = Object.freeze({
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
    language: "vi",
});