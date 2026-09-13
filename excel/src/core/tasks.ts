/**
 * SMART TASK MANAGER - EXCEL (Office.js Task Pane Add-in)
 * Phase 6: Tasks (CRUD thật trên bảng tblTasks)
 *
 * Tương đương google-sheets/tasks.gs. Dùng Excel Table API
 * (table.getDataBodyRange()) để đọc/ghi theo BATCH - đúng nguyên tắc
 * "không context.sync() trong loop từng cell" (mục XI/XII).
 *
 * Mọi thao tác thêm/sửa Task đều đi qua addTask()/updateTask() để đảm bảo
 * SmartScore/Risk/Deadline luôn được Smart Engine tính lại đúng.
 *
 * ĐỒNG BỘ VỚI: google-sheets/tasks.gs (cùng field, cùng validation, cùng
 * hành vi dependency/activity log).
 */

import type { RecommendedAction, RiskLevel, SmartPriorityBand, Task, TaskInput, TaskUpdateInput } from "./types";
import {
    ACTIVITY_ACTION,
    ID_CONFIG,
    PRIORITY_LIST,
    RECURRING_TYPE,
    SHEET_NAMES,
    STATUS,
    STATUS_LIST,
    TABLE_NAMES,
    TASK_COL,
    TASK_FIELDS,
} from "./constants";
import { DEFAULT_APP_SETTINGS } from "./config";
import { computeTaskFields } from "./smart-engine";
import { formatDateOnly, formatDateTimeISO } from "./utils";

type TaskRow = Array<string | number | boolean>;

// ==================================================
// ROW <-> TASK MAPPING
// ==================================================

function rowToTask_(row: ReadonlyArray<unknown>): Task {
    const task = {} as Record<string, unknown>;
    TASK_FIELDS.forEach((field, i) => {
        task[field] = row[i];
    });
    task.progress = Number(task.progress) || 0;
    task.estimatedHours = Number(task.estimatedHours) || 0;
    task.actualHours = Number(task.actualHours) || 0;
    task.isOverdue = task.isOverdue === true || task.isOverdue === "TRUE";
    task.isBlocked = task.isBlocked === true || task.isBlocked === "TRUE";
    task.startDate = normalizeDateCell_(task.startDate);
    task.dueDate = normalizeDateCell_(task.dueDate);
    task.completedDate = normalizeDateCell_(task.completedDate);
    return task as unknown as Task;
}

function normalizeDateCell_(value: unknown): string | null {
    if (!value) return null;
    if (value instanceof Date) return formatDateOnly(value);
    const str = String(value).trim();
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(str);
    return match ? str.substring(0, 10) : null;
}

function taskToRow_(task: Record<string, unknown>): TaskRow {
    return TASK_FIELDS.map((field) => {
        const value = task[field];
        if (value === null || value === undefined) return "";
        if (typeof value === "boolean") return value;
        if (typeof value === "number") return value;
        return String(value);
    });
}

// ==================================================
// ID GENERATION
// ==================================================

/**
 * Sinh TaskId tiếp theo dạng TASK-000001 dựa trên số lớn nhất hiện có
 * trong cột TaskId - không phụ thuộc vị trí hàng (an toàn khi Sort/Filter).
 */
function generateTaskIdFromExisting_(existingTaskIds: ReadonlyArray<string>): string {
    let maxNumber = 0;
    const prefix = ID_CONFIG.TASK.prefix;

    for (const id of existingTaskIds) {
        if (id && id.indexOf(prefix) === 0) {
            const numberPart = parseInt(id.substring(prefix.length), 10);
            if (!Number.isNaN(numberPart) && numberPart > maxNumber) maxNumber = numberPart;
        }
    }

    const nextNumber = maxNumber + 1;
    const padded = String(nextNumber).padStart(ID_CONFIG.TASK.padLength, "0");
    return prefix + padded;
}

/** Sinh TaskId tiếp theo bằng cách đọc thật cột TaskId trên bảng tblTasks. */
export async function generateTaskId(): Promise<string> {
    return Excel.run(async (context) => {
        const table = context.workbook.worksheets
            .getItem(SHEET_NAMES.TASKS)
            .tables.getItem(TABLE_NAMES.TASKS);
        const idColumnRange = table.columns.getItem("Task ID").getDataBodyRange();
        idColumnRange.load("values");
        await context.sync();

        const ids = idColumnRange.values.map((row) => String(row[0] ?? ""));
        return generateTaskIdFromExisting_(ids);
    });
}

// ==================================================
// READ
// ==================================================

/** Đọc toàn bộ Task hiện có - 1 lần getValues() duy nhất, không loop cell. */
export async function getAllTasks(): Promise<Task[]> {
    return Excel.run(async (context) => {
        const table = context.workbook.worksheets
            .getItem(SHEET_NAMES.TASKS)
            .tables.getItem(TABLE_NAMES.TASKS);
        const bodyRange = table.getDataBodyRange();
        bodyRange.load("values, rowCount");
        await context.sync();

        const tasks: Task[] = [];
        for (let r = 0; r < bodyRange.rowCount; r++) {
            const row = bodyRange.values[r];
            if (row && row[TASK_COL.taskId]) tasks.push(rowToTask_(row));
        }
        return tasks;
    });
}

/** Tìm 1 Task theo taskId. Trả về null nếu không tồn tại. */
export async function getTaskById(taskId: string): Promise<Task | null> {
    const tasks = await getAllTasks();
    return tasks.find((t) => t.taskId === taskId) ?? null;
}

function isDependencyResolvedFrom_(tasks: ReadonlyArray<Task>, dependencyTaskId: string | null): boolean {
    if (!dependencyTaskId) return true;
    const dependency = tasks.find((t) => t.taskId === dependencyTaskId);
    if (!dependency) return true;
    return dependency.status === STATUS.COMPLETED;
}

// ==================================================
// VALIDATION
// ==================================================

function validateTaskInput_(input: TaskInput, existingTasks: ReadonlyArray<Task>): void {
    if (!input.taskName || input.taskName.trim() === "") {
        throw new Error("TaskName la bat buoc.");
    }
    if (!input.priority || !PRIORITY_LIST.includes(input.priority)) {
        throw new Error("Priority khong hop le: " + input.priority);
    }
    if (input.status && !STATUS_LIST.includes(input.status)) {
        throw new Error("Status khong hop le: " + input.status);
    }
    if (input.dependencyTaskId && !existingTasks.some((t) => t.taskId === input.dependencyTaskId)) {
        throw new Error("DependencyTaskId khong ton tai: " + input.dependencyTaskId);
    }
    if (
        input.estimatedHours !== undefined &&
        input.estimatedHours !== null &&
        Number.isNaN(Number(input.estimatedHours))
    ) {
        throw new Error("EstimatedHours phai la so.");
    }
}

// ==================================================
// CREATE
// ==================================================

/**
 * Tạo Task mới. Tự động sinh taskId, set createdAt/updatedAt, tính
 * deadlineStatus/smartScore/riskLevel/recommendedAction qua Smart Engine,
 * ghi Activity Log. Toàn bộ trong 1 Excel.run để hạn chế context.sync().
 */
export async function addTask(input: TaskInput): Promise<Task> {
    return Excel.run(async (context) => {
        const tasksTable = context.workbook.worksheets
            .getItem(SHEET_NAMES.TASKS)
            .tables.getItem(TABLE_NAMES.TASKS);
        const bodyRange = tasksTable.getDataBodyRange();
        bodyRange.load("values, rowCount");
        await context.sync();

        const existingTasks: Task[] = [];
        for (let r = 0; r < bodyRange.rowCount; r++) {
            const row = bodyRange.values[r];
            if (row && row[TASK_COL.taskId]) existingTasks.push(rowToTask_(row));
        }

        validateTaskInput_(input, existingTasks);

        const now = new Date();
        const nowISO = formatDateTimeISO(now);
        const status = input.status ?? DEFAULT_APP_SETTINGS.defaultStatus;
        const dependencyTaskId = input.dependencyTaskId ?? null;
        const isBlocked = status === STATUS.BLOCKED;
        const isDependencyResolved = isDependencyResolvedFrom_(existingTasks, dependencyTaskId);

        const computed = computeTaskFields(
            {
                priority: input.priority,
                status,
                progress: Number(input.progress) || 0,
                isBlocked,
                dueDate: input.dueDate ?? null,
            },
            { referenceDate: now, isDependencyResolved }
        );

        const taskId = generateTaskIdFromExisting_(existingTasks.map((t) => t.taskId));

        const task: Record<string, unknown> = {
            taskId,
            taskName: input.taskName.trim(),
            description: input.description ?? "",

            projectId: input.projectId ?? "",
            projectName: "",
            category: input.category ?? "",
            tags: input.tags ?? "",

            ownerId: input.ownerId ?? "",
            ownerName: "",
            ownerEmail: "",
            createdBy: input.createdBy ?? "system",

            priority: input.priority,
            status,

            startDate: input.startDate ?? null,
            dueDate: input.dueDate ?? null,
            completedDate: status === STATUS.COMPLETED ? formatDateOnly(now) : null,

            progress: Number(input.progress) || 0,
            estimatedHours: Number(input.estimatedHours) || 0,
            actualHours: 0,

            daysRemaining: computed.daysRemaining,
            deadlineStatus: computed.deadlineStatus,

            smartScore: computed.smartScore,
            smartPriority: computed.smartPriority,

            riskLevel: computed.riskLevel,
            healthStatus: computed.healthStatus,

            isOverdue: computed.isOverdue,
            isBlocked,

            recommendedAction: computed.recommendedAction,

            dependencyTaskId: dependencyTaskId ?? "",
            recurringType: input.recurringType ?? RECURRING_TYPE.NONE,

            createdAt: nowISO,
            updatedAt: nowISO,
            lastStatusChangedAt: nowISO,

            notes: input.notes ?? "",
        };

        tasksTable.rows.add(undefined, [taskToRow_(task)]);
        await context.sync();

        await appendActivityLog_(context, taskId, ACTIVITY_ACTION.CREATED, "", task.taskName as string);
        await context.sync();

        return task as unknown as Task;
    });
}

// ==================================================
// UPDATE
// ==================================================

/**
 * Cập nhật 1 Task theo taskId. Chỉ nhận field cho phép sửa trực tiếp
 * (TaskUpdateInput) - các field computed luôn được Smart Engine tính lại.
 */
export async function updateTask(taskId: string, patch: TaskUpdateInput): Promise<Task> {
    return Excel.run(async (context) => {
        const tasksTable = context.workbook.worksheets
            .getItem(SHEET_NAMES.TASKS)
            .tables.getItem(TABLE_NAMES.TASKS);
        const bodyRange = tasksTable.getDataBodyRange();
        bodyRange.load("values, rowCount");
        await context.sync();

        let rowIndex = -1;
        const allTasks: Task[] = [];
        for (let r = 0; r < bodyRange.rowCount; r++) {
            const row = bodyRange.values[r];
            if (!row || !row[TASK_COL.taskId]) continue;
            const t = rowToTask_(row);
            allTasks.push(t);
            if (t.taskId === taskId) rowIndex = r;
        }

        if (rowIndex === -1) throw new Error("Khong tim thay Task: " + taskId);
        const existing = allTasks.find((t) => t.taskId === taskId)!;

        if (patch.status && !STATUS_LIST.includes(patch.status)) {
            throw new Error("Status khong hop le: " + patch.status);
        }
        if (patch.priority && !PRIORITY_LIST.includes(patch.priority)) {
            throw new Error("Priority khong hop le: " + patch.priority);
        }
        if (
            patch.progress !== undefined &&
            patch.progress !== null &&
            (Number.isNaN(patch.progress) || patch.progress < 0 || patch.progress > 100)
        ) {
            throw new Error("Progress phai trong khoang 0-100.");
        }
        if (patch.dependencyTaskId && !allTasks.some((t) => t.taskId === patch.dependencyTaskId)) {
            throw new Error("DependencyTaskId khong ton tai: " + patch.dependencyTaskId);
        }

        const now = new Date();
        const nowISO = formatDateTimeISO(now);
        const statusChanged = Boolean(patch.status && patch.status !== existing.status);
        const newStatus = patch.status ?? existing.status;
        const dependencyTaskId =
            patch.dependencyTaskId !== undefined ? patch.dependencyTaskId : existing.dependencyTaskId;
        const isDependencyResolved = isDependencyResolvedFrom_(allTasks, dependencyTaskId);
        const isBlocked = newStatus === STATUS.BLOCKED;

        const mergedForCompute = {
            priority: patch.priority ?? existing.priority,
            status: newStatus,
            progress: patch.progress !== undefined ? Number(patch.progress) : existing.progress,
            isBlocked,
            dueDate: patch.dueDate !== undefined ? patch.dueDate : existing.dueDate,
        };

        const computed = computeTaskFields(mergedForCompute, {
            referenceDate: now,
            isDependencyResolved,
        });

        const updated: Record<string, unknown> = {
            ...existing,
            ...patch,
            priority: mergedForCompute.priority,
            status: newStatus,
            progress: mergedForCompute.progress,
            dueDate: mergedForCompute.dueDate,
            isBlocked,
            dependencyTaskId: dependencyTaskId ?? "",

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
                    ? existing.completedDate ?? formatDateOnly(now)
                    : statusChanged
                    ? null
                    : existing.completedDate,

            updatedAt: nowISO,
            lastStatusChangedAt: statusChanged ? nowISO : existing.lastStatusChangedAt,
        };

        bodyRange.getRow(rowIndex).values = [taskToRow_(updated)];
        await context.sync();

        if (statusChanged) {
            await appendActivityLog_(context, taskId, ACTIVITY_ACTION.STATUS_CHANGED, existing.status, newStatus);
        }
        if (patch.progress !== undefined && Number(patch.progress) !== existing.progress) {
            await appendActivityLog_(
                context,
                taskId,
                ACTIVITY_ACTION.PROGRESS_CHANGED,
                String(existing.progress),
                String(patch.progress)
            );
        }
        if (patch.priority && patch.priority !== existing.priority) {
            await appendActivityLog_(context, taskId, ACTIVITY_ACTION.PRIORITY_CHANGED, existing.priority, patch.priority);
        }
        if (patch.dueDate !== undefined && patch.dueDate !== existing.dueDate) {
            await appendActivityLog_(
                context,
                taskId,
                ACTIVITY_ACTION.DEADLINE_CHANGED,
                String(existing.dueDate),
                String(patch.dueDate)
            );
        }
        await context.sync();

        return updated as unknown as Task;
    });
}

// ==================================================
// ACTIVITY LOG
// ==================================================

async function appendActivityLog_(
    context: Excel.RequestContext,
    taskId: string,
    action: string,
    oldValue: string,
    newValue: string
): Promise<void> {
    const logTable = context.workbook.worksheets
        .getItem(SHEET_NAMES.ACTIVITY_LOG)
        .tables.getItemOrNullObject(TABLE_NAMES.ACTIVITY_LOG);
    logTable.load("name");
    await context.sync();
    if (logTable.isNullObject) return; // chưa setup thì bỏ qua, không làm gãy addTask/updateTask

    logTable.rows.add(undefined, [[formatDateTimeISO(new Date()), "system", taskId, action, oldValue, newValue]]);
}

// ==================================================
// RECOMPUTE ALL
// ==================================================

/**
 * Tính lại field computed cho TOÀN BỘ Task - đọc 1 lần, ghi 1 lần bằng
 * batch values (không loop setValue từng cell). Dùng khi cần refresh
 * Overdue/Due Today sau khi ngày đã sang hôm mới.
 */
export async function recomputeAllTasks(): Promise<number> {
    return Excel.run(async (context) => {
        const table = context.workbook.worksheets
            .getItem(SHEET_NAMES.TASKS)
            .tables.getItem(TABLE_NAMES.TASKS);
        const bodyRange = table.getDataBodyRange();
        bodyRange.load("values, rowCount");
        await context.sync();

        if (bodyRange.rowCount === 0) return 0;

        const now = new Date();
        const rows = bodyRange.values.map(rowToTask_);
        const tasksById = new Map(rows.map((t) => [t.taskId, t] as const));

        const updatedRows = rows.map((task) => {
            let isDependencyResolved = true;
            if (task.dependencyTaskId) {
                const dep = tasksById.get(task.dependencyTaskId);
                isDependencyResolved = !dep || dep.status === STATUS.COMPLETED;
            }

            const computed = computeTaskFields(
                {
                    priority: task.priority,
                    status: task.status,
                    progress: task.progress,
                    isBlocked: task.status === STATUS.BLOCKED,
                    dueDate: task.dueDate,
                },
                { referenceDate: now, isDependencyResolved }
            );

            const merged: Record<string, unknown> = {
                ...task,
                daysRemaining: computed.daysRemaining,
                deadlineStatus: computed.deadlineStatus,
                smartScore: computed.smartScore,
                smartPriority: computed.smartPriority as SmartPriorityBand,
                riskLevel: computed.riskLevel as RiskLevel,
                healthStatus: computed.healthStatus,
                isOverdue: computed.isOverdue,
                isBlocked: task.status === STATUS.BLOCKED,
                recommendedAction: computed.recommendedAction as RecommendedAction,
            };

            return taskToRow_(merged);
        });

        bodyRange.values = updatedRows;
        await context.sync();

        return updatedRows.length;
    });
}