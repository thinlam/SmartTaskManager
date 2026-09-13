/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 1-3: Constants
 *
 * Toàn bộ magic string PHẢI đi qua file này. Không hard-code
 * "In Progress", "Critical", v.v. trực tiếp trong logic.
 *
 * ĐỒNG BỘ VỚI: excel/constants.ts
 * Mọi giá trị string ở đây phải khớp 100% với bản TypeScript.
 *
 * Apps Script chạy toàn bộ file .gs trong cùng một global scope, nên các
 * hằng số dưới đây khả dụng trực tiếp ở mọi file .gs khác trong project
 * (không cần require/import).
 */

// ==================================================
// SHEET NAMES (WORKBOOK STRUCTURE)
// ==================================================

var SHEET_NAMES = Object.freeze({
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
  INSTRUCTIONS: "14_Instructions"
});

// ==================================================
// TABLE / NAMED RANGE NAMES
// ==================================================
// Google Sheets không có "Table" như Excel Table, nhưng ta vẫn đặt tên
// logic thống nhất để dùng làm Named Range bao phủ vùng dữ liệu tương ứng,
// giữ cùng khái niệm "table" giữa 2 platform.

var TABLE_NAMES = Object.freeze({
  TASKS: "tblTasks",
  PROJECTS: "tblProjects",
  MEMBERS: "tblMembers",
  CATEGORIES: "tblCategories",
  ACTIVITY_LOG: "tblActivityLog"
});

// ==================================================
// ID PREFIX / PADDING CONFIG
// ==================================================

var ID_CONFIG = Object.freeze({
  TASK: { prefix: "TASK-", padLength: 6 }, // TASK-000001
  PROJECT: { prefix: "PRJ-", padLength: 4 }, // PRJ-0001
  MEMBER: { prefix: "USR-", padLength: 4 } // USR-0001
});

// ==================================================
// STATUS
// ==================================================

var STATUS = Object.freeze({
  NOT_STARTED: "Not Started",
  TO_DO: "To Do",
  IN_PROGRESS: "In Progress",
  REVIEW: "Review",
  BLOCKED: "Blocked",
  ON_HOLD: "On Hold",
  COMPLETED: "Completed",
  CANCELLED: "Cancelled"
});

var STATUS_LIST = Object.freeze([
  STATUS.NOT_STARTED,
  STATUS.TO_DO,
  STATUS.IN_PROGRESS,
  STATUS.REVIEW,
  STATUS.BLOCKED,
  STATUS.ON_HOLD,
  STATUS.COMPLETED,
  STATUS.CANCELLED
]);

var CLOSED_STATUSES = Object.freeze([STATUS.COMPLETED, STATUS.CANCELLED]);

var ACTIVE_STATUSES = Object.freeze([
  STATUS.NOT_STARTED,
  STATUS.TO_DO,
  STATUS.IN_PROGRESS,
  STATUS.REVIEW,
  STATUS.BLOCKED,
  STATUS.ON_HOLD
]);

// ==================================================
// PRIORITY
// ==================================================

var PRIORITY = Object.freeze({
  CRITICAL: "Critical",
  URGENT: "Urgent",
  HIGH: "High",
  MEDIUM: "Medium",
  LOW: "Low"
});

var PRIORITY_LIST = Object.freeze([
  PRIORITY.CRITICAL,
  PRIORITY.URGENT,
  PRIORITY.HIGH,
  PRIORITY.MEDIUM,
  PRIORITY.LOW
]);

// ==================================================
// RECURRING TYPE
// ==================================================

var RECURRING_TYPE = Object.freeze({
  NONE: "None",
  DAILY: "Daily",
  WEEKLY: "Weekly",
  MONTHLY: "Monthly",
  QUARTERLY: "Quarterly"
});

// ==================================================
// DEADLINE STATUS
// ==================================================

var DEADLINE_STATUS = Object.freeze({
  COMPLETED: "Completed",
  OVERDUE: "Overdue",
  DUE_TODAY: "Due Today",
  DUE_TOMORROW: "Due Tomorrow",
  DUE_IN_3_DAYS: "Due in 3 Days",
  DUE_THIS_WEEK: "Due This Week",
  UPCOMING: "Upcoming",
  NO_DEADLINE: "No Deadline"
});

// ==================================================
// RISK LEVEL
// ==================================================

var RISK_LEVEL = Object.freeze({
  LOW: "Low",
  MEDIUM: "Medium",
  HIGH: "High",
  CRITICAL: "Critical"
});

// ==================================================
// SMART PRIORITY BAND
// ==================================================

var SMART_PRIORITY_BAND = Object.freeze({
  CRITICAL_ATTENTION: "Critical Attention",
  VERY_HIGH_PRIORITY: "Very High Priority",
  HIGH_PRIORITY: "High Priority",
  NORMAL: "Normal",
  LOW_ATTENTION: "Low Attention"
});

// ==================================================
// RECOMMENDED ACTION
// ==================================================

var RECOMMENDED_ACTION = Object.freeze({
  DO_NOW: "Do now",
  COMPLETE_TODAY: "Complete today",
  DEADLINE_APPROACHING: "Deadline approaching",
  REVIEW_BLOCKED_TASK: "Review blocked task",
  FOLLOW_UP_WITH_OWNER: "Follow up with owner",
  WAITING_FOR_DEPENDENCY: "Waiting for dependency",
  SCHEDULE_THIS_TASK: "Schedule this task",
  LOW_PRIORITY: "Low priority",
  COMPLETED: "Completed"
});

// ==================================================
// PROJECT HEALTH / STATUS
// ==================================================

var PROJECT_HEALTH = Object.freeze({
  HEALTHY: "Healthy",
  ATTENTION: "Attention",
  AT_RISK: "At Risk",
  CRITICAL: "Critical"
});

var PROJECT_STATUS = Object.freeze({
  NOT_STARTED: "Not Started",
  IN_PROGRESS: "In Progress",
  ON_HOLD: "On Hold",
  COMPLETED: "Completed",
  CANCELLED: "Cancelled"
});

// ==================================================
// WORKLOAD LEVEL
// ==================================================

var WORKLOAD_LEVEL = Object.freeze({
  AVAILABLE: "Available",
  NORMAL: "Normal",
  BUSY: "Busy",
  OVERLOADED: "Overloaded"
});

// ==================================================
// PRODUCTIVITY BAND
// ==================================================

var PRODUCTIVITY_BAND = Object.freeze({
  EXCELLENT: "Excellent",
  GOOD: "Good",
  NORMAL: "Normal",
  NEEDS_ATTENTION: "Needs Attention",
  CRITICAL: "Critical"
});

// ==================================================
// ACTIVITY LOG ACTIONS
// ==================================================

var ACTIVITY_ACTION = Object.freeze({
  CREATED: "Created",
  UPDATED: "Updated",
  STATUS_CHANGED: "StatusChanged",
  PROGRESS_CHANGED: "ProgressChanged",
  OWNER_CHANGED: "OwnerChanged",
  PRIORITY_CHANGED: "PriorityChanged",
  DEADLINE_CHANGED: "DeadlineChanged",
  DELETED: "Deleted",
  COMPLETED: "Completed"
});

// ==================================================
// COLOR SYSTEM (DESIGN TOKENS)
// ==================================================
// Dùng hex string vì Range.setBackground()/setFontColor() của Apps Script
// nhận trực tiếp hex string giống Office.js.

var DESIGN_COLORS = Object.freeze({
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
  BORDER: "#E2E8F0"
});

/** Màu nền theo Status - dùng cho Conditional Formatting & badge. */
var STATUS_COLORS = Object.freeze((function () {
  var map = {};
  map[STATUS.NOT_STARTED] = "#94A3B8"; // Gray
  map[STATUS.TO_DO] = "#64748B"; // Blue Gray
  map[STATUS.IN_PROGRESS] = "#2563EB"; // Blue
  map[STATUS.REVIEW] = "#7C3AED"; // Purple
  map[STATUS.BLOCKED] = "#DC2626"; // Red
  map[STATUS.ON_HOLD] = "#F59E0B"; // Orange
  map[STATUS.COMPLETED] = "#16A34A"; // Green
  map[STATUS.CANCELLED] = "#334155"; // Dark Gray
  return map;
})());

/** Màu theo Priority. */
var PRIORITY_COLORS = Object.freeze((function () {
  var map = {};
  map[PRIORITY.CRITICAL] = "#DC2626";
  map[PRIORITY.URGENT] = "#EA580C";
  map[PRIORITY.HIGH] = "#F59E0B";
  map[PRIORITY.MEDIUM] = "#0EA5E9";
  map[PRIORITY.LOW] = "#94A3B8";
  return map;
})());

/** Màu theo Risk Level. */
var RISK_COLORS = Object.freeze((function () {
  var map = {};
  map[RISK_LEVEL.LOW] = "#16A34A";
  map[RISK_LEVEL.MEDIUM] = "#F59E0B";
  map[RISK_LEVEL.HIGH] = "#EA580C";
  map[RISK_LEVEL.CRITICAL] = "#DC2626";
  return map;
})());

/** Màu theo Deadline Status. */
var DEADLINE_COLORS = Object.freeze((function () {
  var map = {};
  map[DEADLINE_STATUS.COMPLETED] = "#16A34A";
  map[DEADLINE_STATUS.OVERDUE] = "#DC2626";
  map[DEADLINE_STATUS.DUE_TODAY] = "#EA580C";
  map[DEADLINE_STATUS.DUE_TOMORROW] = "#F59E0B";
  map[DEADLINE_STATUS.DUE_IN_3_DAYS] = "#FACC15";
  map[DEADLINE_STATUS.DUE_THIS_WEEK] = "#0EA5E9";
  map[DEADLINE_STATUS.UPCOMING] = "#94A3B8";
  map[DEADLINE_STATUS.NO_DEADLINE] = "#CBD5E1";
  return map;
})());
