/**
 * SMART TASK MANAGER - EXCEL (Office.js Task Pane Add-in)
 * Phase 6: Setup (tạo workbook structure)
 *
 * Tương đương google-sheets/setup.gs, dùng Excel.run + Excel Table
 * (tblTasks) thay vì Range thô - đúng nguyên tắc "ưu tiên Excel Table cho
 * structured data" đã thống nhất.
 *
 * setupSmartTaskManager() PHẢI idempotent: chạy lại nhiều lần không được
 * tạo trùng worksheet/table, không xóa dữ liệu Task đã có.
 *
 * ĐỒNG BỘ VỚI: google-sheets/setup.gs (cùng SHEET_NAMES, TASK_HEADERS,
 * cùng danh sách key Settings mặc định).
 */

import {
    ACTIVITY_LOG_HEADERS,
    DESIGN_COLORS,
    PRIORITY_COLORS,
    PRIORITY_LIST,
    RECURRING_TYPE,
    RISK_COLORS,
    RISK_LEVEL,
    SHEET_NAMES,
    STATUS_COLORS,
    STATUS_LIST,
    TABLE_NAMES,
    TASK_HEADERS,
} from "./constants";
import { DEFAULT_APP_SETTINGS } from "./config";

const RESERVED_FOR_LATER_PHASE: ReadonlyArray<string> = [
    SHEET_NAMES.DASHBOARD,
    SHEET_NAMES.CALENDAR,
    SHEET_NAMES.KANBAN,
    SHEET_NAMES.GANTT,
    SHEET_NAMES.PROJECTS,
    SHEET_NAMES.MEMBERS,
    SHEET_NAMES.CATEGORIES,
    SHEET_NAMES.REPORTS,
    SHEET_NAMES.NOTIFICATIONS,
    SHEET_NAMES.LISTS,
    SHEET_NAMES.INSTRUCTIONS,
];

const ORDERED_SHEET_NAMES: ReadonlyArray<string> = [
    SHEET_NAMES.DASHBOARD,
    SHEET_NAMES.TASKS,
    SHEET_NAMES.CALENDAR,
    SHEET_NAMES.KANBAN,
    SHEET_NAMES.GANTT,
    SHEET_NAMES.PROJECTS,
    SHEET_NAMES.MEMBERS,
    SHEET_NAMES.CATEGORIES,
    SHEET_NAMES.REPORTS,
    SHEET_NAMES.ACTIVITY_LOG,
    SHEET_NAMES.NOTIFICATIONS,
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.LISTS,
    SHEET_NAMES.INSTRUCTIONS,
];

/**
 * Entry point chính - gọi từ nút "Khởi tạo" trên Task Pane. An toàn khi
 * chạy lại nhiều lần (idempotent).
 */
export async function setupSmartTaskManager(): Promise<void> {
    await Excel.run(async (context) => {
        await createAllWorksheets_(context);
        await setupTasksTable_(context);
        await setupActivityLogTable_(context);
        await setupSettingsSheet_(context);
        await context.sync();
    });
}

// ==================================================
// WORKSHEET CREATION (idempotent)
// ==================================================

async function createAllWorksheets_(context: Excel.RequestContext): Promise<void> {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const existingNames = new Set(worksheets.items.map((ws) => ws.name));

    for (const name of ORDERED_SHEET_NAMES) {
        if (!existingNames.has(name)) {
            const sheet = worksheets.add(name);
            if (RESERVED_FOR_LATER_PHASE.includes(name)) {
                sheet.getRange("A1").values = [
                    [`[${name}] Se duoc xay dung o phase sau (Dashboard/Kanban/Calendar/Gantt/Reports...).`],
                ];
                sheet.getRange("A1").format.font.italic = true;
                sheet.getRange("A1").format.font.color = DESIGN_COLORS.TEXT_SECONDARY;
            }
        }
    }
    await context.sync();

    // Sắp xếp lại đúng thứ tự tab theo ORDERED_SHEET_NAMES.
    for (let i = 0; i < ORDERED_SHEET_NAMES.length; i++) {
        const sheet = context.workbook.worksheets.getItemOrNullObject(ORDERED_SHEET_NAMES[i]!);
        sheet.position = i;
    }
    await context.sync();
}

// ==================================================
// TASKS TABLE SETUP
// ==================================================

/**
 * Tạo bảng tblTasks trên sheet 02_Tasks nếu chưa tồn tại. Idempotent:
 * không tạo lại table nếu đã có, không xóa dữ liệu Task hiện có.
 */
async function setupTasksTable_(context: Excel.RequestContext): Promise<void> {
    const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.TASKS);
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    const existing = tables.items.find((t) => t.name === TABLE_NAMES.TASKS);
    if (existing) {
        // Table đã tồn tại - không đụng vào dữ liệu, chỉ đảm bảo style nhất quán.
        existing.style = "TableStyleMedium2";
        await context.sync();
        return;
    }

    const headerRange = sheet.getRange("A1").getResizedRange(0, TASK_HEADERS.length - 1);
    headerRange.values = [TASK_HEADERS as unknown as string[]];

    const table = sheet.tables.add(headerRange, true);
    table.name = TABLE_NAMES.TASKS;
    table.style = "TableStyleMedium2";

    sheet.freezePanes.freezeRows(1);

    await context.sync();

    await setupTasksDataValidation_(context, sheet, table);
    await setupTasksConditionalFormatting_(context, sheet);
}

async function setupTasksDataValidation_(
    context: Excel.RequestContext,
    sheet: Excel.Worksheet,
    table: Excel.Table
): Promise<void> {
    const statusColIndex = TASK_HEADERS.indexOf("Status");
    const priorityColIndex = TASK_HEADERS.indexOf("Priority");
    const recurringColIndex = TASK_HEADERS.indexOf("Recurring Type");
    const progressColIndex = TASK_HEADERS.indexOf("Progress (%)");

    const maxDataRows = 2000; // giới hạn hợp lý, tránh áp dụng validation lên toàn triệu dòng

    const statusRange = sheet
        .getRange("A2")
        .getOffsetRange(0, statusColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    statusRange.dataValidation.rule = {
        list: { inCellDropDown: true, source: STATUS_LIST.join(",") },
    };

    const priorityRange = sheet
        .getRange("A2")
        .getOffsetRange(0, priorityColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    priorityRange.dataValidation.rule = {
        list: { inCellDropDown: true, source: PRIORITY_LIST.join(",") },
    };

    const recurringValues = [
        RECURRING_TYPE.NONE,
        RECURRING_TYPE.DAILY,
        RECURRING_TYPE.WEEKLY,
        RECURRING_TYPE.MONTHLY,
        RECURRING_TYPE.QUARTERLY,
    ];
    const recurringRange = sheet
        .getRange("A2")
        .getOffsetRange(0, recurringColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    recurringRange.dataValidation.rule = {
        list: { inCellDropDown: true, source: recurringValues.join(",") },
    };

    const progressRange = sheet
        .getRange("A2")
        .getOffsetRange(0, progressColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    progressRange.dataValidation.rule = {
        wholeNumber: { formula1: 0, formula2: 100, operator: Excel.DataValidationOperator.between },
    };

    void table;
    await context.sync();
}

async function setupTasksConditionalFormatting_(
    context: Excel.RequestContext,
    sheet: Excel.Worksheet
): Promise<void> {
    const statusColIndex = TASK_HEADERS.indexOf("Status");
    const priorityColIndex = TASK_HEADERS.indexOf("Priority");
    const riskColIndex = TASK_HEADERS.indexOf("Risk Level");
    const maxDataRows = 2000;

    const statusRange = sheet
        .getRange("A2")
        .getOffsetRange(0, statusColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    const priorityRange = sheet
        .getRange("A2")
        .getOffsetRange(0, priorityColIndex)
        .getResizedRange(maxDataRows - 1, 0);
    const riskRange = sheet
        .getRange("A2")
        .getOffsetRange(0, riskColIndex)
        .getResizedRange(maxDataRows - 1, 0);

    // Xóa rule cũ của đúng 3 range này trước khi thêm lại - tránh cộng dồn
    // qua nhiều lần chạy setup (idempotent).
    statusRange.conditionalFormats.clearAll();
    priorityRange.conditionalFormats.clearAll();
    riskRange.conditionalFormats.clearAll();
    await context.sync();

    STATUS_LIST.forEach((statusValue) => {
        const cf = statusRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.rule = {
            formula1: `="${statusValue}"`,
            operator: Excel.ConditionalCellValueOperator.equalTo,
        };
        cf.cellValue.format.fill.color = STATUS_COLORS[statusValue];
        cf.cellValue.format.font.color = "#FFFFFF";
    });

    PRIORITY_LIST.forEach((priorityValue) => {
        const cf = priorityRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.rule = {
            formula1: `="${priorityValue}"`,
            operator: Excel.ConditionalCellValueOperator.equalTo,
        };
        cf.cellValue.format.fill.color = PRIORITY_COLORS[priorityValue];
        cf.cellValue.format.font.color = "#FFFFFF";
    });

    ([RISK_LEVEL.LOW, RISK_LEVEL.MEDIUM, RISK_LEVEL.HIGH, RISK_LEVEL.CRITICAL] as const).forEach(
        (riskValue) => {
            const cf = riskRange.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cf.cellValue.rule = {
                formula1: `="${riskValue}"`,
                operator: Excel.ConditionalCellValueOperator.equalTo,
            };
            cf.cellValue.format.fill.color = RISK_COLORS[riskValue];
            cf.cellValue.format.font.color = "#FFFFFF";
        }
    );

    await context.sync();
}

// ==================================================
// ACTIVITY LOG TABLE SETUP
// ==================================================

async function setupActivityLogTable_(context: Excel.RequestContext): Promise<void> {
    const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.ACTIVITY_LOG);
    const tables = sheet.tables;
    tables.load("items/name");
    await context.sync();

    const existing = tables.items.find((t) => t.name === TABLE_NAMES.ACTIVITY_LOG);
    if (existing) return;

    const headerRange = sheet.getRange("A1").getResizedRange(0, ACTIVITY_LOG_HEADERS.length - 1);
    headerRange.values = [ACTIVITY_LOG_HEADERS as unknown as string[]];

    const table = sheet.tables.add(headerRange, true);
    table.name = TABLE_NAMES.ACTIVITY_LOG;
    table.style = "TableStyleMedium3";

    sheet.freezePanes.freezeRows(1);
    await context.sync();
}

// ==================================================
// SETTINGS SHEET SETUP
// ==================================================

/**
 * Settings lưu dạng 3 cột Key/Value/Description (không dùng Table, vì
 * đây là danh sách cấu hình cố định, không cần filter/sort như Tasks).
 * Chỉ ghi key CHƯA tồn tại - không ghi đè key user đã tự chỉnh.
 */
async function setupSettingsSheet_(context: Excel.RequestContext): Promise<void> {
    const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.SETTINGS);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("values, rowCount");
    await context.sync();

    const existingKeys = new Set<string>();
    let nextRow = 1;

    if (!usedRange.isNullObject && usedRange.rowCount > 1) {
        const values = usedRange.values as unknown[][];
        for (let r = 1; r < values.length; r++) {
            const key = values[r]?.[0];
            if (key) existingKeys.add(String(key));
        }
        nextRow = values.length;
    } else {
        const headerRange = sheet.getRange("A1:C1");
        headerRange.values = [["Key", "Value", "Description"]];
        headerRange.format.font.bold = true;
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.fill.color = DESIGN_COLORS.PRIMARY;
        sheet.freezePanes.freezeRows(1);
        nextRow = 1;
    }

    const defaultRows: [string, string | number | boolean, string][] = [
        ["systemName", DEFAULT_APP_SETTINGS.systemName, "Ten he thong hien thi tren Dashboard"],
        ["defaultStatus", DEFAULT_APP_SETTINGS.defaultStatus, "Status mac dinh khi tao Task moi"],
        ["defaultPriority", DEFAULT_APP_SETTINGS.defaultPriority, "Priority mac dinh khi tao Task moi"],
        ["dueSoonDays", DEFAULT_APP_SETTINGS.dueSoonDays, "So ngay tinh la 'Due Soon' tren KPI"],
        ["workingDays", DEFAULT_APP_SETTINGS.workingDays.join(","), "Cac ngay lam viec trong tuan (1=Mon..7=Sun)"],
        ["workingHoursPerDay", DEFAULT_APP_SETTINGS.workingHoursPerDay, "So gio lam viec/ngay"],
        ["weekendDays", DEFAULT_APP_SETTINGS.weekendDays.join(","), "Cac ngay cuoi tuan"],
        ["notificationsEnabled", DEFAULT_APP_SETTINGS.notificationsEnabled, "Bat/tat toan bo notification"],
        ["notificationEmailOnOverdue", DEFAULT_APP_SETTINGS.notificationEmailOnOverdue, "Gui email khi Task Overdue"],
        ["notificationEmailOnDueToday", DEFAULT_APP_SETTINGS.notificationEmailOnDueToday, "Gui email khi Task Due Today"],
        ["notificationEmailOnDueTomorrow", DEFAULT_APP_SETTINGS.notificationEmailOnDueTomorrow, "Gui email khi Task Due Tomorrow"],
        ["notificationEmailOnCritical", DEFAULT_APP_SETTINGS.notificationEmailOnCritical, "Gui email khi Task Critical"],
        ["theme", DEFAULT_APP_SETTINGS.theme, "Light hoac Dark"],
        ["dateFormat", DEFAULT_APP_SETTINGS.dateFormat, "Dinh dang ngay hien thi"],
        ["currency", DEFAULT_APP_SETTINGS.currency, "Don vi tien te"],
        ["language", DEFAULT_APP_SETTINGS.language, "Ngon ngu he thong"],
    ];

    const rowsToAppend = defaultRows.filter((row) => !existingKeys.has(row[0]));

    if (rowsToAppend.length > 0) {
        const appendRange = sheet
            .getRange("A1")
            .getOffsetRange(nextRow, 0)
            .getResizedRange(rowsToAppend.length - 1, 2);
        appendRange.values = rowsToAppend;
    }

    sheet.getRange("A:A").format.columnWidth = 220;
    sheet.getRange("B:B").format.columnWidth = 200;
    sheet.getRange("C:C").format.columnWidth = 320;

    await context.sync();
}

/**
 * Đọc 1 giá trị Setting theo key. Trả về defaultValue nếu key không tồn tại.
 */
export async function getSetting(key: string, defaultValue: unknown): Promise<unknown> {
    return Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.SETTINGS);
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount");
        await context.sync();

        if (sheet.isNullObject || usedRange.isNullObject || usedRange.rowCount < 2) {
            return defaultValue;
        }

        const values = usedRange.values as unknown[][];
        for (let r = 1; r < values.length; r++) {
            if (values[r]?.[0] === key) return values[r]?.[1] ?? defaultValue;
        }
        return defaultValue;
    });
}