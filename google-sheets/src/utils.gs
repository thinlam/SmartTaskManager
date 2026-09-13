/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 4: Utils (date & number helpers dùng cho Smart Engine)
 *
 * ĐỒNG BỘ VỚI: excel/utils.ts
 * Cùng công thức tính ngày, cùng cách clamp/round số.
 */

/**
 * Parse chuỗi ISO date "yyyy-MM-dd" thành Date ở mốc 00:00:00 (local time
 * theo timeZone khai báo trong appsscript.json).
 * @param {string|null|undefined} dateStr
 * @returns {Date|null}
 */
function parseDateOnly(dateStr) {
  if (!dateStr) return null;
  var match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(dateStr).trim());
  if (!match) return null;
  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);
  var date = new Date(year, month - 1, day, 0, 0, 0, 0);
  if (isNaN(date.getTime())) return null;
  return date;
}

/**
 * @param {Date} date
 * @returns {string} "yyyy-MM-dd"
 */
function formatDateOnly(date) {
  var year = date.getFullYear();
  var month = String(date.getMonth() + 1).padStart(2, "0");
  var day = String(date.getDate()).padStart(2, "0");
  return year + "-" + month + "-" + day;
}

/**
 * @param {Date} date
 * @returns {string} ISO datetime
 */
function formatDateTimeISO(date) {
  return date.toISOString();
}

/**
 * @param {Date} [referenceDate]
 * @returns {Date} mốc 00:00:00 hôm nay
 */
function getTodayDateOnly(referenceDate) {
  var base = referenceDate || new Date();
  return new Date(base.getFullYear(), base.getMonth(), base.getDate(), 0, 0, 0, 0);
}

/**
 * @param {Date} fromDate
 * @param {Date} targetDate
 * @returns {number} targetDate - fromDate, tính theo ngày (âm nếu targetDate ở quá khứ)
 */
function diffInCalendarDays(fromDate, targetDate) {
  var MS_PER_DAY = 24 * 60 * 60 * 1000;
  var from = getTodayDateOnly(fromDate);
  var target = getTodayDateOnly(targetDate);
  return Math.round((target.getTime() - from.getTime()) / MS_PER_DAY);
}

/**
 * @param {number} value
 * @param {number} min
 * @param {number} max
 * @returns {number}
 */
function clamp(value, min, max) {
  if (isNaN(value)) return min;
  return Math.max(min, Math.min(max, value));
}

/**
 * @param {number} value
 * @returns {number}
 */
function roundSafe(value) {
  return isFinite(value) ? Math.round(value) : 0;
}