/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 1-3: Data Model (JSDoc Typedefs)
 *
 * Google Apps Script không có type system thật, nên file này dùng JSDoc
 * @typedef để mô tả "hợp đồng dữ liệu" (data contract), giúp autocomplete
 * trong Apps Script Editor và làm tài liệu tham chiếu.
 *
 * QUAN TRỌNG:
 * Field name và giá trị enum-string PHẢI khớp với excel/types.ts.
 * Đây là bản GAS tương đương của Task/Project/Member/... bên Excel.
 *
 * File này KHÔNG chứa logic, chỉ chứa định nghĩa cấu trúc dữ liệu.
 */

/**
 * @typedef {"Not Started"|"To Do"|"In Progress"|"Review"|"Blocked"|"On Hold"|"Completed"|"Cancelled"} TaskStatus
 */

/**
 * @typedef {"Critical"|"Urgent"|"High"|"Medium"|"Low"} TaskPriority
 */

/**
 * @typedef {"Completed"|"Overdue"|"Due Today"|"Due Tomorrow"|"Due in 3 Days"|"Due This Week"|"Upcoming"|"No Deadline"} DeadlineStatus
 */

/**
 * @typedef {"Low"|"Medium"|"High"|"Critical"} RiskLevel
 */

/**
 * @typedef {"Critical Attention"|"Very High Priority"|"High Priority"|"Normal"|"Low Attention"} SmartPriorityBand
 */

/**
 * @typedef {"Do now"|"Complete today"|"Deadline approaching"|"Review blocked task"|"Follow up with owner"|"Waiting for dependency"|"Schedule this task"|"Low priority"|"Completed"} RecommendedAction
 */

/**
 * @typedef {"None"|"Daily"|"Weekly"|"Monthly"|"Quarterly"} RecurringType
 */

/**
 * @typedef {"Healthy"|"Attention"|"At Risk"|"Critical"} ProjectHealth
 */

/**
 * @typedef {"Not Started"|"In Progress"|"On Hold"|"Completed"|"Cancelled"} ProjectStatus
 */

/**
 * @typedef {"Available"|"Normal"|"Busy"|"Overloaded"} WorkloadLevel
 */

/**
 * @typedef {"Excellent"|"Good"|"Normal"|"Needs Attention"|"Critical"} ProductivityBand
 */

/**
 * @typedef {"Created"|"Updated"|"StatusChanged"|"ProgressChanged"|"OwnerChanged"|"PriorityChanged"|"DeadlineChanged"|"Deleted"|"Completed"} ActivityAction
 */

/**
 * Task - Single Source of Truth. Thứ tự field bên dưới cũng chính là
 * THỨ TỰ CỘT trong sheet 02_Tasks (dùng bởi setup.gs khi tạo header).
 *
 * @typedef {Object} Task
 * @property {string} taskId - "TASK-000001"
 * @property {string} taskName
 * @property {string} description
 * @property {string} projectId - "PRJ-0001"
 * @property {string} projectName
 * @property {string} category
 * @property {string} tags - comma-separated
 * @property {string} ownerId - "USR-0001"
 * @property {string} ownerName
 * @property {string} ownerEmail
 * @property {string} createdBy
 * @property {TaskPriority} priority
 * @property {TaskStatus} status
 * @property {string|null} startDate - ISO date "yyyy-MM-dd"
 * @property {string|null} dueDate - ISO date "yyyy-MM-dd"
 * @property {string|null} completedDate - ISO date "yyyy-MM-dd"
 * @property {number} progress - 0-100
 * @property {number} estimatedHours
 * @property {number} actualHours
 * @property {number|null} daysRemaining
 * @property {DeadlineStatus} deadlineStatus
 * @property {number} smartScore - 0-100
 * @property {SmartPriorityBand} smartPriority
 * @property {RiskLevel} riskLevel
 * @property {string} healthStatus
 * @property {boolean} isOverdue
 * @property {boolean} isBlocked
 * @property {RecommendedAction} recommendedAction
 * @property {string|null} dependencyTaskId
 * @property {RecurringType} recurringType
 * @property {string} createdAt - ISO datetime
 * @property {string} updatedAt - ISO datetime
 * @property {string} lastStatusChangedAt - ISO datetime
 * @property {string} notes
 */

/**
 * @typedef {Object} TaskInput
 * @property {string} taskName
 * @property {string} [description]
 * @property {string} [projectId]
 * @property {string} [category]
 * @property {string} [tags]
 * @property {string} [ownerId]
 * @property {TaskPriority} priority
 * @property {TaskStatus} [status]
 * @property {string|null} [startDate]
 * @property {string|null} [dueDate]
 * @property {number} [estimatedHours]
 * @property {string|null} [dependencyTaskId]
 * @property {RecurringType} [recurringType]
 * @property {string} [notes]
 * @property {string} [createdBy]
 */

/**
 * @typedef {Object} TaskUpdateInput
 * @property {string} [taskName]
 * @property {string} [description]
 * @property {string} [projectId]
 * @property {string} [category]
 * @property {string} [tags]
 * @property {string} [ownerId]
 * @property {TaskPriority} [priority]
 * @property {TaskStatus} [status]
 * @property {string|null} [startDate]
 * @property {string|null} [dueDate]
 * @property {number} [progress]
 * @property {number} [estimatedHours]
 * @property {number} [actualHours]
 * @property {string|null} [dependencyTaskId]
 * @property {RecurringType} [recurringType]
 * @property {string} [notes]
 */

/**
 * @typedef {Object} Project
 * @property {string} projectId - "PRJ-0001"
 * @property {string} projectName
 * @property {string} owner
 * @property {string|null} startDate
 * @property {string|null} endDate
 * @property {TaskPriority} priority
 * @property {number} progress
 * @property {ProjectHealth} health
 * @property {ProjectStatus} status
 * @property {string} description
 * @property {number} totalTasks
 * @property {number} completedTasks
 * @property {number} inProgressTasks
 * @property {number} overdueTasks
 * @property {number} blockedTasks
 * @property {number} healthScore
 * @property {string} createdAt
 * @property {string} updatedAt
 */

/**
 * @typedef {Object} Member
 * @property {string} memberId - "USR-0001"
 * @property {string} fullName
 * @property {string} email
 * @property {string} role
 * @property {string} department
 * @property {boolean} active
 * @property {number} workloadLimit
 * @property {number} assignedTasks
 * @property {number} completedTasks
 * @property {number} activeTasks
 * @property {number} overdueTasks
 * @property {number} completionRate
 * @property {number} onTimeRate
 * @property {number} averageCompletionTimeDays
 * @property {WorkloadLevel} workloadLevel
 * @property {number} productivityScore
 * @property {ProductivityBand} productivityBand
 * @property {string} createdAt
 * @property {string} updatedAt
 */

/**
 * @typedef {Object} Category
 * @property {string} categoryName
 * @property {string} description
 * @property {string} colorHex
 * @property {boolean} active
 */

/**
 * @typedef {Object} ActivityLogEntry
 * @property {string} timestamp
 * @property {string} user
 * @property {string} taskId
 * @property {ActivityAction} action
 * @property {string} oldValue
 * @property {string} newValue
 */

/**
 * @typedef {Object} AppSettings
 * @property {string} systemName
 * @property {TaskStatus} defaultStatus
 * @property {TaskPriority} defaultPriority
 * @property {number} dueSoonDays
 * @property {number[]} workingDays - 1=Monday ... 7=Sunday
 * @property {number} workingHoursPerDay
 * @property {number[]} weekendDays
 * @property {boolean} notificationsEnabled
 * @property {boolean} notificationEmailOnOverdue
 * @property {boolean} notificationEmailOnDueToday
 * @property {boolean} notificationEmailOnDueTomorrow
 * @property {boolean} notificationEmailOnCritical
 * @property {"Light"|"Dark"} theme
 * @property {string} dateFormat
 * @property {string} currency
 * @property {"vi"|"en"} language
 */

/**
 * @typedef {Object} SmartScoreBreakdown
 * @property {number} priorityScore
 * @property {number} deadlineScore
 * @property {number} overdueScore
 * @property {number} progressRiskScore
 * @property {number} blockedScore
 * @property {number} dependencyScore
 * @property {number} projectWeightScore
 * @property {number} totalScore
 * @property {SmartPriorityBand} band
 */

/**
 * @typedef {Object} RiskBreakdown
 * @property {number} overdueRisk
 * @property {number} deadlineProximityRisk
 * @property {number} progressRisk
 * @property {number} priorityRisk
 * @property {number} blockedRisk
 * @property {number} dependencyRisk
 * @property {number} agingRisk
 * @property {number} totalRiskScore
 * @property {RiskLevel} level
 */

// Không có export trong Apps Script (global scope theo file .gs trong cùng
// project) - các file khác (constants.gs, config.gs, tasks.gs, ...) dùng
// trực tiếp các typedef này qua JSDoc @param / @returns để được gợi ý type.
