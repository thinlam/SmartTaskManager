/**
 * SMART TASK MANAGER - EXCEL / OFFICE SCRIPTS
 * Phase 1-3: Data Model (Types & Interfaces)
 *
 * File này định nghĩa toàn bộ type/interface dùng chung cho:
 * - Tasks
 * - Projects
 * - Members
 * - Categories
 * - Activity Log
 * - Settings
 *
 * QUAN TRỌNG:
 * Các union type string literal bên dưới (TaskStatus, TaskPriority, ...)
 * PHẢI có giá trị chuỗi giống hệt file constants.gs / types.gs của bản
 * Google Sheets. Đây là "business rule contract" giữa 2 platform.
 * Không được đổi giá trị chuỗi ở một bên mà không đổi bên còn lại.
 */

// ==================================================
// ENUM-LIKE STRING LITERAL TYPES (SOURCE OF TRUTH)
// ==================================================

export type TaskStatus =
    | "Not Started"
    | "To Do"
    | "In Progress"
    | "Review"
    | "Blocked"
    | "On Hold"
    | "Completed"
    | "Cancelled";

export type TaskPriority =
    | "Critical"
    | "Urgent"
    | "High"
    | "Medium"
    | "Low";

export type DeadlineStatus =
    | "Completed"
    | "Overdue"
    | "Due Today"
    | "Due Tomorrow"
    | "Due in 3 Days"
    | "Due This Week"
    | "Upcoming"
    | "No Deadline";

export type RiskLevel = "Low" | "Medium" | "High" | "Critical";

export type SmartPriorityBand =
    | "Critical Attention"
    | "Very High Priority"
    | "High Priority"
    | "Normal"
    | "Low Attention";

export type RecommendedAction =
    | "Do now"
    | "Complete today"
    | "Deadline approaching"
    | "Review blocked task"
    | "Follow up with owner"
    | "Waiting for dependency"
    | "Schedule this task"
    | "Low priority"
    | "Completed";

export type RecurringType =
    | "None"
    | "Daily"
    | "Weekly"
    | "Monthly"
    | "Quarterly";

export type ProjectHealth = "Healthy" | "Attention" | "At Risk" | "Critical";

export type ProjectStatus =
    | "Not Started"
    | "In Progress"
    | "On Hold"
    | "Completed"
    | "Cancelled";

export type WorkloadLevel = "Available" | "Normal" | "Busy" | "Overloaded";

export type ProductivityBand =
    | "Excellent"
    | "Good"
    | "Normal"
    | "Needs Attention"
    | "Critical";

export type ActivityAction =
    | "Created"
    | "Updated"
    | "StatusChanged"
    | "ProgressChanged"
    | "OwnerChanged"
    | "PriorityChanged"
    | "DeadlineChanged"
    | "Deleted"
    | "Completed";

// ==================================================
// CORE DATA MODEL
// ==================================================

/**
 * Task - Single Source of Truth của toàn hệ thống.
 * Mọi view (Dashboard, Kanban, Calendar, Gantt, Reports) đều suy ra từ Task.
 */
export interface Task {
    // Identity
    taskId: string; // TASK-000001
    taskName: string;
    description: string;

    // Relations
    projectId: string; // PRJ-0001
    projectName: string;
    category: string;
    tags: string; // comma-separated, ví dụ: "backend,urgent"

    // Ownership
    ownerId: string; // USR-0001
    ownerName: string;
    ownerEmail: string;
    createdBy: string;

    // Classification
    priority: TaskPriority;
    status: TaskStatus;

    // Timeline
    startDate: string | null; // ISO date string "yyyy-MM-dd"
    dueDate: string | null;
    completedDate: string | null;

    // Progress
    progress: number; // 0 - 100
    estimatedHours: number;
    actualHours: number;

    // Derived / Computed (do Smart Engine tính, không cho user nhập tay)
    daysRemaining: number | null;
    deadlineStatus: DeadlineStatus;
    smartScore: number; // 0 - 100
    smartPriority: SmartPriorityBand;
    riskLevel: RiskLevel;
    healthStatus: string; // ví dụ: "On Track" | "Waiting Dependency" | "At Risk"
    isOverdue: boolean;
    isBlocked: boolean;
    recommendedAction: RecommendedAction;

    // Dependency & Recurrence
    dependencyTaskId: string | null;
    recurringType: RecurringType;

    // Audit
    createdAt: string; // ISO datetime
    updatedAt: string;
    lastStatusChangedAt: string;

    // Free text
    notes: string;
}

/**
 * Input tối thiểu cần thiết khi tạo Task mới (Quick Add Task).
 * Các field còn lại sẽ được Smart Engine / setup logic tự tính ở Phase 4+.
 */
export interface TaskInput {
    taskName: string;
    description?: string;
    projectId?: string;
    category?: string;
    tags?: string;
    ownerId?: string;
    priority: TaskPriority;
    status?: TaskStatus;
    startDate?: string | null;
    dueDate?: string | null;
    progress?: number;
    estimatedHours?: number;
    dependencyTaskId?: string | null;
    recurringType?: RecurringType;
    notes?: string;
    createdBy?: string;
}

/**
 * Patch dùng cho Quick Update Task - chỉ chứa field được phép chỉnh sửa
 * trực tiếp bởi người dùng (không cho sửa field computed).
 */
export interface TaskUpdateInput {
    taskName?: string;
    description?: string;
    projectId?: string;
    category?: string;
    tags?: string;
    ownerId?: string;
    priority?: TaskPriority;
    status?: TaskStatus;
    startDate?: string | null;
    dueDate?: string | null;
    progress?: number;
    estimatedHours?: number;
    actualHours?: number;
    dependencyTaskId?: string | null;
    recurringType?: RecurringType;
    notes?: string;
}

export interface Project {
    projectId: string; // PRJ-0001
    projectName: string;
    owner: string;
    startDate: string | null;
    endDate: string | null;
    priority: TaskPriority;
    progress: number; // 0-100, tính từ Tasks thuộc project
    health: ProjectHealth;
    status: ProjectStatus;
    description: string;

    // Computed (Phase 4+)
    totalTasks: number;
    completedTasks: number;
    inProgressTasks: number;
    overdueTasks: number;
    blockedTasks: number;
    healthScore: number; // 0-100

    createdAt: string;
    updatedAt: string;
}

export interface Member {
    memberId: string; // USR-0001
    fullName: string;
    email: string;
    role: string;
    department: string;
    active: boolean;
    workloadLimit: number; // số giờ hoặc số task tối đa/tuần, xem SETTINGS

    // Computed (Phase 4+)
    assignedTasks: number;
    completedTasks: number;
    activeTasks: number;
    overdueTasks: number;
    completionRate: number; // %
    onTimeRate: number; // %
    averageCompletionTimeDays: number;
    workloadLevel: WorkloadLevel;
    productivityScore: number; // 0-100
    productivityBand: ProductivityBand;

    createdAt: string;
    updatedAt: string;
}

export interface Category {
    categoryName: string;
    description: string;
    colorHex: string;
    active: boolean;
}

export interface ActivityLogEntry {
    timestamp: string; // ISO datetime
    user: string;
    taskId: string;
    action: ActivityAction;
    oldValue: string;
    newValue: string;
}

/**
 * Cấu hình có thể thay đổi bởi người dùng qua sheet 12_Settings.
 * Không hard-code các giá trị này rải rác trong business logic.
 */
export interface AppSettings {
    systemName: string;
    defaultStatus: TaskStatus;
    defaultPriority: TaskPriority;
    dueSoonDays: number;
    workingDays: number[]; // 1=Monday ... 7=Sunday
    workingHoursPerDay: number;
    weekendDays: number[];
    notificationsEnabled: boolean;
    notificationEmailOnOverdue: boolean;
    notificationEmailOnDueToday: boolean;
    notificationEmailOnDueTomorrow: boolean;
    notificationEmailOnCritical: boolean;
    theme: "Light" | "Dark";
    dateFormat: string; // ví dụ "dd/MM/yyyy"
    currency: string;
    language: "vi" | "en";
}

/**
 * Kết quả breakdown khi tính SmartScore - dùng để giải thích điểm số
 * (yêu cầu XXXIV: SmartScore phải giải thích được).
 * Được implement thực tế ở Phase 4 (smart-engine).
 */
export interface SmartScoreBreakdown {
    priorityScore: number;
    deadlineScore: number;
    overdueScore: number;
    progressRiskScore: number;
    blockedScore: number;
    dependencyScore: number;
    projectWeightScore: number;
    totalScore: number; // tổng, đã clamp 0-100
    band: SmartPriorityBand;
}

/**
 * Kết quả breakdown khi tính Risk - tách biệt với SmartScore vì
 * mục đích sử dụng khác nhau (Risk Engine vs Priority Engine).
 */
export interface RiskBreakdown {
    overdueRisk: number;
    deadlineProximityRisk: number;
    progressRisk: number;
    priorityRisk: number;
    blockedRisk: number;
    dependencyRisk: number;
    agingRisk: number;
    totalRiskScore: number; // 0-100
    level: RiskLevel;
}