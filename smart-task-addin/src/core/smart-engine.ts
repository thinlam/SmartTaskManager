/**
 * SMART TASK MANAGER - EXCEL / OFFICE SCRIPTS
 * Phase 4: Smart Engine
 *
 * Chứa logic tính toán thuần (pure function), không đụng tới Office.js/
 * Excel Range. Nhận Task (+ context phụ) làm input, trả về field computed.
 * Layer gọi Excel API (Phase 5/6) sẽ đọc Range -> map thành Task -> gọi các
 * hàm ở đây -> ghi kết quả trở lại Range.
 *
 * ĐỒNG BỘ VỚI: google-sheets/smart-engine.gs
 * Cùng công thức, cùng thứ tự rule, cùng ngưỡng số (lấy từ config.ts).
 */

import type {
    DeadlineStatus,
    RecommendedAction,
    RiskBreakdown,
    RiskLevel,
    SmartPriorityBand,
    SmartScoreBreakdown,
    Task,
    TaskPriority,
    TaskStatus,
} from "./types";
import {
    CLOSED_STATUSES,
    DEADLINE_STATUS,
    RECOMMENDED_ACTION,
    RISK_LEVEL,
    STATUS,
} from "./constants";
import {
    DEADLINE_CONFIG,
    RISK_CONFIG,
    RISK_LEVEL_BANDS,
    SMART_SCORE_BANDS,
    SMART_SCORE_WEIGHTS,
} from "./config";
import { clamp, diffInCalendarDays, getTodayDateOnly, parseDateOnly, roundSafe } from "./utils";

/** Context phụ cần thiết để tính SmartScore/Risk mà bản thân Task không tự chứa đủ. */
export interface SmartEngineContext {
    /** Mốc thời gian tham chiếu, mặc định là "bây giờ". Cho phép truyền vào để unit test/deterministic. */
    referenceDate?: Date;
    /** Priority của Project chứa Task (dùng cho Project Weight Score). Nếu không có Project, bỏ qua field này. */
    projectPriority?: TaskPriority;
    /** true nếu Task không có dependency, hoặc dependency đã Completed. false nếu dependency tồn tại và chưa xong. */
    isDependencyResolved?: boolean;
    /** Số ngày kể từ khi Task được tạo (dùng cho Task Aging trong Risk Engine). */
    taskAgeDays?: number;
}

// ==================================================
// DEADLINE ENGINE
// ==================================================

/**
 * Tính DeadlineStatus dựa trên status hiện tại và dueDate.
 * Không yêu cầu user nhập tay - luôn suy ra từ dữ liệu thật.
 */
export function calculateDeadlineStatus(
    status: TaskStatus,
    dueDate: string | null,
    referenceDate: Date = new Date()
): DeadlineStatus {
    if (CLOSED_STATUSES.includes(status)) {
        return status === STATUS.COMPLETED ? DEADLINE_STATUS.COMPLETED : DEADLINE_STATUS.NO_DEADLINE;
    }

    const due = parseDateOnly(dueDate);
    if (!due) return DEADLINE_STATUS.NO_DEADLINE;

    const daysRemaining = diffInCalendarDays(getTodayDateOnly(referenceDate), due);

    if (daysRemaining < 0) return DEADLINE_STATUS.OVERDUE;
    if (daysRemaining === 0) return DEADLINE_STATUS.DUE_TODAY;
    if (daysRemaining === 1) return DEADLINE_STATUS.DUE_TOMORROW;
    if (daysRemaining <= DEADLINE_CONFIG.dueInDays) return DEADLINE_STATUS.DUE_IN_3_DAYS;
    if (daysRemaining <= DEADLINE_CONFIG.dueThisWeekDays) return DEADLINE_STATUS.DUE_THIS_WEEK;
    return DEADLINE_STATUS.UPCOMING;
}

/**
 * Số ngày còn lại tới deadline. Âm nghĩa là đã trễ (số ngày trễ).
 * Trả về null nếu không có dueDate hoặc Task đã đóng (Completed/Cancelled).
 */
export function calculateDaysRemaining(
    status: TaskStatus,
    dueDate: string | null,
    referenceDate: Date = new Date()
): number | null {
    if (CLOSED_STATUSES.includes(status)) return null;
    const due = parseDateOnly(dueDate);
    if (!due) return null;
    return diffInCalendarDays(getTodayDateOnly(referenceDate), due);
}

// ==================================================
// SMART SCORE ENGINE
// ==================================================

function priorityScoreOf(priority: TaskPriority): number {
    return SMART_SCORE_WEIGHTS.PRIORITY_POINTS[priority] ?? 0;
}

function deadlineScoreOf(deadlineStatus: DeadlineStatus): number {
    const points = SMART_SCORE_WEIGHTS.DEADLINE_POINTS;
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

function overdueScoreOf(daysRemaining: number | null): number {
    if (daysRemaining === null || daysRemaining >= 0) return 0;
    const daysOverdue = Math.abs(daysRemaining);
    return clamp(
        daysOverdue * SMART_SCORE_WEIGHTS.OVERDUE_POINTS_PER_DAY,
        0,
        SMART_SCORE_WEIGHTS.OVERDUE_MAX
    );
}

function progressRiskScoreOf(progress: number, status: TaskStatus): number {
    if (CLOSED_STATUSES.includes(status)) return 0;
    const points = SMART_SCORE_WEIGHTS.PROGRESS_RISK_POINTS;
    if (progress < 25) return points.BELOW_25;
    if (progress < 50) return points.BELOW_50;
    if (progress < 75) return points.BELOW_75;
    return points.BELOW_100;
}

function projectWeightScoreOf(projectPriority: TaskPriority | undefined): number {
    if (!projectPriority) return 0;
    return SMART_SCORE_WEIGHTS.PROJECT_WEIGHT_POINTS[projectPriority] ?? 0;
}

function bandOfSmartScore(score: number): SmartPriorityBand {
    for (const entry of SMART_SCORE_BANDS) {
        if (score >= entry.min) return entry.band;
    }
    return SMART_SCORE_BANDS[SMART_SCORE_BANDS.length - 1]!.band;
}

/**
 * Tính SmartScore (0-100) + breakdown giải thích được cho 1 Task.
 * Không random - hoàn toàn suy ra từ dữ liệu Task + context.
 */
export function calculateSmartScore(
    task: Pick<Task, "priority" | "status" | "progress" | "isBlocked" | "dueDate">,
    deadlineStatus: DeadlineStatus,
    daysRemaining: number | null,
    ctx: SmartEngineContext = {}
): SmartScoreBreakdown {
    // Task đã Completed/Cancelled thì không cần "chú ý" nữa -> điểm thấp có chủ đích.
    if (CLOSED_STATUSES.includes(task.status)) {
        return {
            priorityScore: 0,
            deadlineScore: 0,
            overdueScore: 0,
            progressRiskScore: 0,
            blockedScore: 0,
            dependencyScore: 0,
            projectWeightScore: 0,
            totalScore: 0,
            band: bandOfSmartScore(0),
        };
    }

    const priorityScore = priorityScoreOf(task.priority);
    const overdueScore = overdueScoreOf(daysRemaining);
    // Deadline score chỉ tính khi CHƯA overdue (overdue đã có điểm riêng, tránh cộng dồn 2 lần).
    const deadlineScore = overdueScore > 0 ? 0 : deadlineScoreOf(deadlineStatus);
    const progressRiskScore = progressRiskScoreOf(task.progress, task.status);
    const blockedScore = task.isBlocked ? SMART_SCORE_WEIGHTS.BLOCKED_POINTS : 0;
    const dependencyScore =
        ctx.isDependencyResolved === false ? SMART_SCORE_WEIGHTS.DEPENDENCY_POINTS : 0;
    const projectWeightScore = projectWeightScoreOf(ctx.projectPriority);

    const rawTotal =
        priorityScore +
        deadlineScore +
        overdueScore +
        progressRiskScore +
        blockedScore +
        dependencyScore +
        projectWeightScore;

    const totalScore = roundSafe(
        clamp(rawTotal, SMART_SCORE_WEIGHTS.SCORE_MIN, SMART_SCORE_WEIGHTS.SCORE_MAX)
    );

    return {
        priorityScore,
        deadlineScore,
        overdueScore,
        progressRiskScore,
        blockedScore,
        dependencyScore,
        projectWeightScore,
        totalScore,
        band: bandOfSmartScore(totalScore),
    };
}

// ==================================================
// RISK ENGINE
// ==================================================

function overdueRiskOf(daysRemaining: number | null): number {
    if (daysRemaining === null || daysRemaining >= 0) return 0;
    const daysOverdue = Math.abs(daysRemaining);
    return clamp(
        daysOverdue * RISK_CONFIG.WEIGHTS.OVERDUE_POINTS_PER_DAY,
        0,
        RISK_CONFIG.WEIGHTS.OVERDUE_MAX
    );
}

function deadlineProximityRiskOf(deadlineStatus: DeadlineStatus, isOverdue: boolean): number {
    if (isOverdue) return 0; // đã tính ở overdueRisk, tránh cộng dồn
    const points = RISK_CONFIG.WEIGHTS.DEADLINE_PROXIMITY_POINTS;
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

function progressRiskOf(progress: number, status: TaskStatus): number {
    if (CLOSED_STATUSES.includes(status)) return 0;
    const points = RISK_CONFIG.WEIGHTS.PROGRESS_RISK_POINTS;
    if (progress < 25) return points.BELOW_25;
    if (progress < 50) return points.BELOW_50;
    if (progress < 75) return points.BELOW_75;
    return points.BELOW_100;
}

function priorityRiskOf(priority: TaskPriority): number {
    return RISK_CONFIG.WEIGHTS.PRIORITY_POINTS[priority] ?? 0;
}

function agingRiskOf(taskAgeDays: number | undefined): number {
    if (!taskAgeDays) return 0;
    const thresholds = RISK_CONFIG.WEIGHTS.AGING_THRESHOLDS_DAYS;
    const points = RISK_CONFIG.WEIGHTS.AGING_POINTS;
    if (taskAgeDays >= thresholds.LEVEL_2) return points.LEVEL_2;
    if (taskAgeDays >= thresholds.LEVEL_1) return points.LEVEL_1;
    return 0;
}

function riskLevelOf(score: number): RiskLevel {
    for (const entry of RISK_LEVEL_BANDS) {
        if (score >= entry.min) return entry.level;
    }
    return RISK_LEVEL_BANDS[RISK_LEVEL_BANDS.length - 1]!.level;
}

/**
 * Tính Risk (0-100) + breakdown cho 1 Task. Tách biệt với SmartScore vì
 * mục đích khác nhau: SmartScore trả lời "nên làm gì trước", Risk trả lời
 * "task này có nguy cơ trễ/hỏng tiến độ không".
 */
export function calculateRisk(
    task: Pick<Task, "priority" | "status" | "progress" | "isBlocked">,
    deadlineStatus: DeadlineStatus,
    daysRemaining: number | null,
    ctx: SmartEngineContext = {}
): RiskBreakdown {
    if (CLOSED_STATUSES.includes(task.status)) {
        return {
            overdueRisk: 0,
            deadlineProximityRisk: 0,
            progressRisk: 0,
            priorityRisk: 0,
            blockedRisk: 0,
            dependencyRisk: 0,
            agingRisk: 0,
            totalRiskScore: 0,
            level: riskLevelOf(0),
        };
    }

    const isOverdue = deadlineStatus === DEADLINE_STATUS.OVERDUE;
    const overdueRisk = overdueRiskOf(daysRemaining);
    const deadlineProximityRisk = deadlineProximityRiskOf(deadlineStatus, isOverdue);
    const progressRisk = progressRiskOf(task.progress, task.status);
    const priorityRisk = priorityRiskOf(task.priority);
    const blockedRisk = task.isBlocked ? RISK_CONFIG.WEIGHTS.BLOCKED_POINTS : 0;
    const dependencyRisk =
        ctx.isDependencyResolved === false ? RISK_CONFIG.WEIGHTS.DEPENDENCY_POINTS : 0;
    const agingRisk = agingRiskOf(ctx.taskAgeDays);

    const rawTotal =
        overdueRisk +
        deadlineProximityRisk +
        progressRisk +
        priorityRisk +
        blockedRisk +
        dependencyRisk +
        agingRisk;

    const totalRiskScore = roundSafe(clamp(rawTotal, RISK_CONFIG.SCORE_MIN, RISK_CONFIG.SCORE_MAX));

    return {
        overdueRisk,
        deadlineProximityRisk,
        progressRisk,
        priorityRisk,
        blockedRisk,
        dependencyRisk,
        agingRisk,
        totalRiskScore,
        level: riskLevelOf(totalRiskScore),
    };
}

// ==================================================
// RECOMMENDATION ENGINE
// ==================================================
// Áp dụng đúng thứ tự rule đã khai báo ở RECOMMENDATION_RULES (config.ts).
// Rule đầu tiên khớp điều kiện sẽ quyết định kết quả.

export function calculateRecommendedAction(
    task: Pick<Task, "status" | "isBlocked">,
    deadlineStatus: DeadlineStatus,
    smartScore: number,
    isDependencyResolved: boolean
): RecommendedAction {
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
// HEALTH STATUS (chuỗi mô tả ngắn hiển thị trên Dashboard/Task detail)
// ==================================================

export function calculateHealthStatus(
    task: Pick<Task, "status" | "isBlocked">,
    deadlineStatus: DeadlineStatus,
    isDependencyResolved: boolean,
    riskLevel: RiskLevel
): string {
    if (task.status === STATUS.COMPLETED) return "Completed";
    if (task.status === STATUS.CANCELLED) return "Cancelled";
    if (task.isBlocked) return "Blocked";
    if (!isDependencyResolved) return "Waiting Dependency";
    if (deadlineStatus === DEADLINE_STATUS.OVERDUE) return "Overdue";
    if (riskLevel === RISK_LEVEL.CRITICAL || riskLevel === RISK_LEVEL.HIGH) return "At Risk";
    return "On Track";
}

// ==================================================
// ORCHESTRATION - tính toàn bộ field computed cho 1 Task trong 1 lần gọi
// ==================================================

export interface ComputedTaskFields {
    daysRemaining: number | null;
    deadlineStatus: DeadlineStatus;
    smartScore: number;
    smartScoreBreakdown: SmartScoreBreakdown;
    smartPriority: SmartPriorityBand;
    riskLevel: RiskLevel;
    riskBreakdown: RiskBreakdown;
    healthStatus: string;
    isOverdue: boolean;
    recommendedAction: RecommendedAction;
}

/**
 * Hàm tổng hợp - input là 1 Task (đã có status/priority/progress/isBlocked/
 * dueDate) + context, output là toàn bộ field computed cần ghi lại vào
 * sheet Tasks. Đây là hàm mà layer Excel API (Phase 5/6) sẽ gọi cho mỗi
 * dòng Task khi refresh.
 */
export function computeTaskFields(
    task: Pick<Task, "priority" | "status" | "progress" | "isBlocked" | "dueDate">,
    ctx: SmartEngineContext = {}
): ComputedTaskFields {
    const referenceDate = ctx.referenceDate ?? new Date();
    const isDependencyResolved = ctx.isDependencyResolved ?? true;

    const deadlineStatus = calculateDeadlineStatus(task.status, task.dueDate, referenceDate);
    const daysRemaining = calculateDaysRemaining(task.status, task.dueDate, referenceDate);
    const isOverdue = deadlineStatus === DEADLINE_STATUS.OVERDUE;

    const smartScoreBreakdown = calculateSmartScore(task, deadlineStatus, daysRemaining, ctx);
    const riskBreakdown = calculateRisk(task, deadlineStatus, daysRemaining, ctx);
    const recommendedAction = calculateRecommendedAction(
        task,
        deadlineStatus,
        smartScoreBreakdown.totalScore,
        isDependencyResolved
    );
    const healthStatus = calculateHealthStatus(
        task,
        deadlineStatus,
        isDependencyResolved,
        riskBreakdown.level
    );

    return {
        daysRemaining,
        deadlineStatus,
        smartScore: smartScoreBreakdown.totalScore,
        smartScoreBreakdown,
        smartPriority: smartScoreBreakdown.band,
        riskLevel: riskBreakdown.level,
        riskBreakdown,
        healthStatus,
        isOverdue,
        recommendedAction,
    };
}