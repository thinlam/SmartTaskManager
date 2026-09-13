/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 8a: Kanban
 *
 * refreshKanban() la view phai sinh tu 02_Tasks (dung nguyen tac Single
 * Source of Truth, muc XLVI: "Kanban phai lay du lieu tu Tasks. Khong tao
 * mot bo Task rieng trong Kanban"). Moi lan goi se doc lai toan bo Task,
 * nhom theo Status, ghi de TOAN BO vung Kanban - khong con Task cu sot lai
 * khi Status da doi cot.
 *
 * Mapping Status -> cot Kanban (muc XLVI):
 *   Not Started, To Do  -> TO DO
 *   In Progress          -> IN PROGRESS
 *   Review                -> REVIEW
 *   Blocked                -> BLOCKED
 *   Completed              -> DONE
 *   Cancelled               -> khong hien thi tren board (task da huy,
 *                              van con nguyen trong 02_Tasks va Reports)
 *
 * Phu thuoc global: types.gs, constants.gs, tasks.gs
 */

var KANBAN_COLUMNS = Object.freeze([
  { title: "TO DO", statuses: [STATUS.NOT_STARTED, STATUS.TO_DO], color: STATUS_COLORS[STATUS.TO_DO] },
  { title: "IN PROGRESS", statuses: [STATUS.IN_PROGRESS], color: STATUS_COLORS[STATUS.IN_PROGRESS] },
  { title: "REVIEW", statuses: [STATUS.REVIEW], color: STATUS_COLORS[STATUS.REVIEW] },
  { title: "BLOCKED", statuses: [STATUS.BLOCKED], color: STATUS_COLORS[STATUS.BLOCKED] },
  { title: "DONE", statuses: [STATUS.COMPLETED], color: STATUS_COLORS[STATUS.COMPLETED] }
]);

var KANBAN_LAYOUT = Object.freeze({
  HEADER_ROW: 1,
  CARD_START_ROW: 2,
  MAX_CARDS_PER_COLUMN: 200,
  COL_WIDTH: 2
});

/**
 * Tinh lai toan bo Kanban tu du lieu Tasks hien tai. An toan khi goi
 * nhieu lan (idempotent) - luon ghi de dung vung quan ly.
 */
function refreshKanban() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.KANBAN);
  if (!sheet) throw new Error("Sheet " + SHEET_NAMES.KANBAN + " chua duoc tao. Chay setupSmartTaskManager() truoc.");

  var tasks = getAllTasks();

  clearKanbanManagedRange_(sheet);
  writeKanbanColumns_(sheet, tasks);

  SpreadsheetApp.flush();
  var visibleCount = tasks.filter(function (t) { return t.status !== STATUS.CANCELLED; }).length;
  Logger.log("refreshKanban() hoan tat. Tong so Task hien thi: " + visibleCount);
}

function clearKanbanManagedRange_(sheet) {
  var maxCol = KANBAN_COLUMNS.length * KANBAN_LAYOUT.COL_WIDTH;
  var maxRow = KANBAN_LAYOUT.CARD_START_ROW + KANBAN_LAYOUT.MAX_CARDS_PER_COLUMN;
  sheet.getRange(1, 1, maxRow, maxCol).clearContent();
  sheet.getRange(1, 1, maxRow, maxCol).clearFormat();
}

function writeKanbanColumns_(sheet, tasks) {
  for (var c = 0; c < KANBAN_COLUMNS.length; c++) {
    var column = KANBAN_COLUMNS[c];
    var startCol = c * KANBAN_LAYOUT.COL_WIDTH + 1;

    var columnTasks = tasks.filter(function (t) {
      return column.statuses.indexOf(t.status) !== -1;
    });

    columnTasks.sort(function (a, b) {
      return Number(b.smartScore) - Number(a.smartScore);
    });

    var headerRange = sheet.getRange(KANBAN_LAYOUT.HEADER_ROW, startCol, 1, KANBAN_LAYOUT.COL_WIDTH);
    headerRange.merge();
    headerRange.setValue(column.title + " (" + columnTasks.length + ")");
    headerRange
      .setFontWeight("bold")
      .setFontColor("#FFFFFF")
      .setBackground(column.color)
      .setHorizontalAlignment("center");

    sheet.setColumnWidth(startCol, 200);
    sheet.setColumnWidth(startCol + 1, 20);

    if (columnTasks.length === 0) {
      var emptyCell = sheet.getRange(KANBAN_LAYOUT.CARD_START_ROW, startCol);
      emptyCell.setValue("(khong co task)");
      emptyCell.setFontColor(DESIGN_COLORS.TEXT_SECONDARY).setFontStyle("italic");
      continue;
    }

    var maxCards = Math.min(columnTasks.length, KANBAN_LAYOUT.MAX_CARDS_PER_COLUMN);
    for (var r = 0; r < maxCards; r++) {
      var t = columnTasks[r];
      var cardRow = KANBAN_LAYOUT.CARD_START_ROW + r;
      var cardText = t.taskId + "\n" + t.taskName + "\n[" + t.priority + "]";

      var cardCell = sheet.getRange(cardRow, startCol);
      cardCell.setValue(cardText);
      cardCell.setWrap(true);
      cardCell.setVerticalAlignment("top");
      cardCell.setBackground(DESIGN_COLORS.SURFACE);
      cardCell.setBorder(true, true, true, true, false, false, DESIGN_COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

      var priorityColor = PRIORITY_COLORS[t.priority];
      if (priorityColor) cardCell.setFontColor(priorityColor);

      sheet.setRowHeight(cardRow, 60);
    }
  }
}