/**
 * SMART TASK MANAGER - EXCEL / OFFICE SCRIPTS
 * Phase 1-3: Constants
 *
 * Toàn bộ magic string trong hệ thống PHẢI đi qua file này.
 * Không được viết lại chuỗi "In Progress", "Critical", v.v. trực tiếp
 * trong logic - luôn tham chiếu qua STATUS.* / PRIORITY.*
 *
 * ĐỒNG BỘ VỚI: google-sheets/constants.gs
 * Mọi giá trị string ở đây phải khớp 100% với bản Google Apps Script.
 */

import type { TaskPriority, TaskStatus } from "./types";

// ==================================================
// SHEET NAMES (WORKBOOK STRUCTURE)
// ==================================================

export const SHEET_NAMES = Object.freeze({
    DASHBOARD: "01_Dashboard",
    TASKS: "02_Tasks",
    CALENDAR: "03_Calendar",
    KANBAN: "04_Kanban",
    GANTT: "05_Gantt",
    PROJECTS: "06_Projects",
    MEMBERS: "07_Members",
    CATEGORIES: "08_Categories",
    REPORTS: "09_Reports",
    ACTIVITY_LOG: "10_Activity_Log",
    NOTIFICATIONS: "11_Notifications",
    SETTINGS: "12_Settings",
    LISTS: "13_Lists",
    INSTRUCTIONS: "14_Instructions",
} as const);

export type SheetNameKey = keyof typeof SHEET_NAMES;

// ==================================================
// EXCEL TABLE NAMES
// ==================================================

export const TABLE_NAMES = Object.freeze({
    TASKS: "tblTasks",
    PROJECTS: "tblProjects",
    MEMBERS: "tblMembers",
    CATEGORIES: "tblCategories",
    ACTIVITY_LOG: "tblActivityLog",
} as const);

// ==================================================
// ID PREFIX / PADDING CONFIG
// ==================================================

export const ID_CONFIG = Object.freeze({
    TASK: { prefix: "TASK-", padLength: 6 }, // TASK-000001
    PROJECT: { prefix: "PRJ-", padLength: 4 }, // PRJ-0001
    MEMBER: { prefix: "USR-", padLength: 4 }, // USR-0001
} as const);

// ==================================================
// STATUS (giá trị hiển thị = giá trị lưu trong cell)
// ==================================================

export const STATUS = Object.freeze({
    NOT_STARTED: "Not Started",
    TO_DO: "To Do",
    IN_PROGRESS: "In Progress",
    REVIEW: "Review",
    BLOCKED: "Blocked",
    ON_HOLD: "On Hold",
    COMPLETED: "Completed",
    CANCELLED: "Cancelled",
} as const);

export const STATUS_LIST: TaskStatus[] = [
    STATUS.NOT_STARTED,
    STATUS.TO_DO,
    STATUS.IN_PROGRESS,
    STATUS.REVIEW,
    STATUS.BLOCKED,
    STATUS.ON_HOLD,
    STATUS.COMPLETED,
    STATUS.CANCELLED,
];

/** Các status được coi là "đã kết thúc vòng đời" của Task. */
export const CLOSED_STATUSES: TaskStatus[] = [STATUS.COMPLETED, STATUS.CANCELLED];

/** Các status được coi là "đang hoạt động" (dùng cho workload, KPI...). */
export const ACTIVE_STATUSES: TaskStatus[] = [
    STATUS.NOT_STARTED,
    STATUS.TO_DO,
    STATUS.IN_PROGRESS,
    STATUS.REVIEW,
    STATUS.BLOCKED,
    STATUS.ON_HOLD,
];

// ==================================================
// PRIORITY
// ==================================================

export const PRIORITY = Object.freeze({
    CRITICAL: "Critical",
    URGENT: "Urgent",
    HIGH: "High",
    MEDIUM: "Medium",
    LOW: "Low",
} as const);

export const PRIORITY_LIST: TaskPriority[] = [
    PRIORITY.CRITICAL,
    PRIORITY.URGENT,
    PRIORITY.HIGH,
    PRIORITY.MEDIUM,
    PRIORITY.LOW,
];

// ==================================================
// RECURRING TYPE
// ==================================================

export const RECURRING_TYPE = Object.freeze({
    NONE: "None",
    DAILY: "Daily",
    WEEKLY: "Weekly",
    MONTHLY: "Monthly",
    QUARTERLY: "Quarterly",
} as const);

// ==================================================
// DEADLINE STATUS
// ==================================================

export const DEADLINE_STATUS = Object.freeze({
    COMPLETED: "Completed",
    OVERDUE: "Overdue",
    DUE_TODAY: "Due Today",
    DUE_TOMORROW: "Due Tomorrow",
    DUE_IN_3_DAYS: "Due in 3 Days",
    DUE_THIS_WEEK: "Due This Week",
    UPCOMING: "Upcoming",
    NO_DEADLINE: "No Deadline",
} as const);

// ==================================================
// RISK LEVEL
// ==================================================

export const RISK_LEVEL = Object.freeze({
    LOW: "Low",
    MEDIUM: "Medium",
    HIGH: "High",
    CRITICAL: "Critical",
} as const);

// ==================================================
// SMART PRIORITY BAND (phân loại SmartScore)
// ==================================================

export const SMART_PRIORITY_BAND = Object.freeze({
    CRITICAL_ATTENTION: "Critical Attention",
    VERY_HIGH_PRIORITY: "Very High Priority",
    HIGH_PRIORITY: "High Priority",
    NORMAL: "Normal",
    LOW_ATTENTION: "Low Attention",
} as const);

// ==================================================
// RECOMMENDED ACTION
// ==================================================

export const RECOMMENDED_ACTION = Object.freeze({
    DO_NOW: "Do now",
    COMPLETE_TODAY: "Complete today",
    DEADLINE_APPROACHING: "Deadline approaching",
    REVIEW_BLOCKED_TASK: "Review blocked task",
    FOLLOW_UP_WITH_OWNER: "Follow up with owner",
    WAITING_FOR_DEPENDENCY: "Waiting for dependency",
    SCHEDULE_THIS_TASK: "Schedule this task",
    LOW_PRIORITY: "Low priority",
    COMPLETED: "Completed",
} as const);

// ==================================================
// PROJECT HEALTH / STATUS
// ==================================================

export const PROJECT_HEALTH = Object.freeze({
    HEALTHY: "Healthy",
    ATTENTION: "Attention",
    AT_RISK: "At Risk",
    CRITICAL: "Critical",
} as const);

export const PROJECT_STATUS = Object.freeze({
    NOT_STARTED: "Not Started",
    IN_PROGRESS: "In Progress",
    ON_HOLD: "On Hold",
    COMPLETED: "Completed",
    CANCELLED: "Cancelled",
} as const);

// ==================================================
// WORKLOAD LEVEL
// ==================================================

export const WORKLOAD_LEVEL = Object.freeze({
    AVAILABLE: "Available",
    NORMAL: "Normal",
    BUSY: "Busy",
    OVERLOADED: "Overloaded",
} as const);

// ==================================================
// PRODUCTIVITY BAND
// ==================================================

export const PRODUCTIVITY_BAND = Object.freeze({
    EXCELLENT: "Excellent",
    GOOD: "Good",
    NORMAL: "Normal",
    NEEDS_ATTENTION: "Needs Attention",
    CRITICAL: "Critical",
} as const);

// ==================================================
// ACTIVITY LOG ACTIONS
// ==================================================

export const ACTIVITY_ACTION = Object.freeze({
    CREATED: "Created",
    UPDATED: "Updated",
    STATUS_CHANGED: "StatusChanged",
    PROGRESS_CHANGED: "ProgressChanged",
    OWNER_CHANGED: "OwnerChanged",
    PRIORITY_CHANGED: "PriorityChanged",
    DEADLINE_CHANGED: "DeadlineChanged",
    DELETED: "Deleted",
    COMPLETED: "Completed",
} as const);

// ==================================================
// COLOR SYSTEM (DESIGN TOKENS)
// ==================================================
// Dùng hex string vì Office Scripts / Office.js Range.format.fill.color
// nhận trực tiếp hex string. Đây là single source of truth cho màu sắc,
// không hard-code hex ở nơi khác.

export const DESIGN_COLORS = Object.freeze({
    PRIMARY: "#2563EB",
    PRIMARY_DARK: "#1E40AF",
    SUCCESS: "#16A34A",
    WARNING: "#F59E0B",
    DANGER: "#DC2626",
    INFO: "#0EA5E9",
    BACKGROUND: "#F8FAFC",
    SURFACE: "#FFFFFF",
    TEXT_PRIMARY: "#0F172A",
    TEXT_SECONDARY: "#64748B",
    BORDER: "#E2E8F0",
} as const);

/** Màu nền theo Status - dùng cho Conditional Formatting & badge trong Dashboard. */
export const STATUS_COLORS: Record<TaskStatus, string> = Object.freeze({
    [STATUS.NOT_STARTED]: "#94A3B8", // Gray
    [STATUS.TO_DO]: "#64748B", // Blue Gray
    [STATUS.IN_PROGRESS]: "#2563EB", // Blue
    [STATUS.REVIEW]: "#7C3AED", // Purple
    [STATUS.BLOCKED]: "#DC2626", // Red
    [STATUS.ON_HOLD]: "#F59E0B", // Orange
    [STATUS.COMPLETED]: "#16A34A", // Green
    [STATUS.CANCELLED]: "#334155", // Dark Gray
});

/** Màu theo Priority - dùng cho Conditional Formatting & badge. */
export const PRIORITY_COLORS: Record<TaskPriority, string> = Object.freeze({
    [PRIORITY.CRITICAL]: "#DC2626",
    [PRIORITY.URGENT]: "#EA580C",
    [PRIORITY.HIGH]: "#F59E0B",
    [PRIORITY.MEDIUM]: "#0EA5E9",
    [PRIORITY.LOW]: "#94A3B8",
});

/** Màu theo Risk Level. */
export const RISK_COLORS = Object.freeze({
    [RISK_LEVEL.LOW]: "#16A34A",
    [RISK_LEVEL.MEDIUM]: "#F59E0B",
    [RISK_LEVEL.HIGH]: "#EA580C",
    [RISK_LEVEL.CRITICAL]: "#DC2626",
} as const);

/** Màu theo Deadline Status - dùng cho Calendar highlight & Dashboard alert. */
export const DEADLINE_COLORS = Object.freeze({
    [DEADLINE_STATUS.COMPLETED]: "#16A34A",
    [DEADLINE_STATUS.OVERDUE]: "#DC2626",
    [DEADLINE_STATUS.DUE_TODAY]: "#EA580C",
    [DEADLINE_STATUS.DUE_TOMORROW]: "#F59E0B",
    [DEADLINE_STATUS.DUE_IN_3_DAYS]: "#FACC15",
    [DEADLINE_STATUS.DUE_THIS_WEEK]: "#0EA5E9",
    [DEADLINE_STATUS.UPCOMING]: "#94A3B8",
    [DEADLINE_STATUS.NO_DEADLINE]: "#CBD5E1",
} as const);

// ==================================================
// SHEET COLUMN SCHEMA (chuẩn bị cho Phase 6 - Office Scripts)
// ==================================================
// ĐỒNG BỘ VỚI: google-sheets/constants.gs (TASK_FIELDS / TASK_HEADERS)
// TASK_FIELDS quyết định thứ tự cột thật trong bảng tblTasks.

export const TASK_FIELDS = [
    "taskId", "taskName", "description",
    "projectId", "projectName", "category", "tags",
    "ownerId", "ownerName", "ownerEmail", "createdBy",
    "priority", "status",
    "startDate", "dueDate", "completedDate",
    "progress", "estimatedHours", "actualHours",
    "daysRemaining", "deadlineStatus",
    "smartScore", "smartPriority",
    "riskLevel", "healthStatus",
    "isOverdue", "isBlocked",
    "recommendedAction",
    "dependencyTaskId", "recurringType",
    "createdAt", "updatedAt", "lastStatusChangedAt",
    "notes",
] as const;

export const TASK_HEADERS = [
    "Task ID", "Task Name", "Description",
    "Project ID", "Project Name", "Category", "Tags",
    "Owner ID", "Owner Name", "Owner Email", "Created By",
    "Priority", "Status",
    "Start Date", "Due Date", "Completed Date",
    "Progress (%)", "Estimated Hours", "Actual Hours",
    "Days Remaining", "Deadline Status",
    "Smart Score", "Smart Priority",
    "Risk Level", "Health Status",
    "Is Overdue", "Is Blocked",
    "Recommended Action",
    "Dependency Task ID", "Recurring Type",
    "Created At", "Updated At", "Last Status Changed At",
    "Notes",
] as const;

export const ACTIVITY_LOG_FIELDS = [
    "timestamp", "user", "taskId", "action", "oldValue", "newValue",
] as const;

export const ACTIVITY_LOG_HEADERS = [
    "Timestamp", "User", "Task ID", "Action", "Old Value", "New Value",
] as const;

/** Index (0-based) của các cột quan trọng, dùng khi đọc/ghi Range/Table theo vị trí. */
export const TASK_COL: Record<(typeof TASK_FIELDS)[number], number> = TASK_FIELDS.reduce(
    (map, field, index) => {
        map[field] = index;
        return map;
    },
    {} as Record<(typeof TASK_FIELDS)[number], number>
);