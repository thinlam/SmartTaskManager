/**
 * SMART TASK MANAGER - EXCEL / OFFICE SCRIPTS
 * Phase 4: Utils (date & number helpers dùng cho Smart Engine)
 *
 * Chỉ chứa hàm thuần (pure function), không phụ thuộc Office.js/Excel API,
 * để có thể unit test độc lập bằng Node/ts-node.
 *
 * ĐỒNG BỘ VỚI: google-sheets/utils.gs
 * Cùng công thức tính ngày, cùng cách làm tròn/clamp số.
 */

/**
 * Parse chuỗi ISO date "yyyy-MM-dd" thành Date ở mốc 00:00:00 (local time).
 * Trả về null nếu input rỗng/không hợp lệ.
 */
export function parseDateOnly(dateStr: string | null | undefined): Date | null {
    if (!dateStr) return null;
    const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(dateStr.trim());
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const date = new Date(year, month - 1, day, 0, 0, 0, 0);
    if (Number.isNaN(date.getTime())) return null;
    return date;
}

/** Format Date thành chuỗi ISO date "yyyy-MM-dd" (bỏ qua giờ). */
export function formatDateOnly(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

/** Format Date thành chuỗi ISO datetime đầy đủ (dùng cho createdAt/updatedAt). */
export function formatDateTimeISO(date: Date): string {
    return date.toISOString();
}

/** Trả về Date ở mốc 00:00:00 hôm nay (local time). */
export function getTodayDateOnly(referenceDate?: Date): Date {
    const base = referenceDate ?? new Date();
    return new Date(base.getFullYear(), base.getMonth(), base.getDate(), 0, 0, 0, 0);
}

/**
 * Số ngày chênh lệch giữa 2 mốc ngày (đã bỏ giờ), targetDate - fromDate.
 * Dương: targetDate ở tương lai. Âm: targetDate ở quá khứ.
 */
export function diffInCalendarDays(fromDate: Date, targetDate: Date): number {
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    const from = getTodayDateOnly(fromDate);
    const target = getTodayDateOnly(targetDate);
    return Math.round((target.getTime() - from.getTime()) / MS_PER_DAY);
}

/** Giới hạn number trong khoảng [min, max]. */
export function clamp(value: number, min: number, max: number): number {
    if (Number.isNaN(value)) return min;
    return Math.max(min, Math.min(max, value));
}

/** Làm tròn về số nguyên, an toàn với NaN. */
export function roundSafe(value: number): number {
    return Number.isFinite(value) ? Math.round(value) : 0;
}