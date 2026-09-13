/**
 * SMART TASK MANAGER - GOOGLE SHEETS (Apps Script)
 * Phase 9 (rut gon): Custom Menu
 *
 * onOpen() la trigger dac biet cua Google Sheets - tu dong chay moi khi
 * mo file, khong can goi thu cong. Tao menu "SMART TASK MANAGER" tren
 * thanh menu chinh de nguoi dung khong can vao Apps Script Editor.
 *
 * Phu thuoc global: setup.gs, tasks.gs, 07_dashboard.gs
 */

/**
 * Trigger dac biet - Google tu goi ham nay moi khi Spreadsheet duoc mo.
 * Khong duoc doi ten ham nay.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ Smart Task Manager")
    .addItem("📊 Refresh Dashboard", "menuRefreshDashboard_")
    .addItem("🔄 Recompute All Tasks", "menuRecomputeAllTasks_")
    .addItem("🗂️ Refresh Kanban", "menuRefreshKanban_")
    .addItem("📅 Refresh Calendar", "menuRefreshCalendar_")
    .addSeparator()
    .addItem("🔔 Check Deadlines Now (gui email neu can)", "menuCheckDeadlines_")
    .addItem("⏰ Bat nhac nho hang ngay (07:00)", "menuEnableDailyReminders_")
    .addItem("🔕 Tat nhac nho hang ngay", "menuDisableDailyReminders_")
    .addSeparator()
    .addItem("🛠️ Run Setup (Khoi tao / Sua workbook)", "menuRunSetup_")
    .addToUi();
}

// ==================================================
// MENU HANDLERS (wrap ham logic + hien thong bao cho nguoi dung)
// ==================================================

function menuRefreshDashboard_() {
  var ui = SpreadsheetApp.getUi();
  try {
    refreshDashboard();
    ui.alert("Dashboard da duoc cap nhat.");
  } catch (e) {
    ui.alert("Loi khi refresh Dashboard: " + e.message);
  }
}

function menuRecomputeAllTasks_() {
  var ui = SpreadsheetApp.getUi();
  try {
    recomputeAllTasks();
    ui.alert("Da tinh lai SmartScore/Risk/Deadline cho toan bo Task.");
  } catch (e) {
    ui.alert("Loi khi recompute: " + e.message);
  }
}

function menuRefreshKanban_() {
  var ui = SpreadsheetApp.getUi();
  try {
    refreshKanban();
    ui.alert("Kanban da duoc cap nhat.");
  } catch (e) {
    ui.alert("Loi khi refresh Kanban: " + e.message);
  }
}

function menuRefreshCalendar_() {
  var ui = SpreadsheetApp.getUi();
  try {
    refreshCalendar();
    ui.alert("Calendar da duoc cap nhat cho thang hien tai.");
  } catch (e) {
    ui.alert("Loi khi refresh Calendar: " + e.message);
  }
}

function menuCheckDeadlines_() {
  var ui = SpreadsheetApp.getUi();
  try {
    checkDeadlines();
    ui.alert("Da kiem tra deadline. Xem chi tiet o sheet 11_Notifications.");
  } catch (e) {
    ui.alert("Loi khi check deadlines: " + e.message);
  }
}

function menuEnableDailyReminders_() {
  var ui = SpreadsheetApp.getUi();
  try {
    createTriggers();
    ui.alert("Da bat nhac nho hang ngay luc 07:00.");
  } catch (e) {
    ui.alert("Loi khi tao trigger: " + e.message);
  }
}

function menuDisableDailyReminders_() {
  var ui = SpreadsheetApp.getUi();
  try {
    removeDeadlineTriggers_();
    ui.alert("Da tat nhac nho hang ngay.");
  } catch (e) {
    ui.alert("Loi khi xoa trigger: " + e.message);
  }
}

function menuRunSetup_() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "Xac nhan",
    "Chay lai Setup se dam bao du 14 sheet, header, data validation dung chuan. Du lieu Task hien co se KHONG bi mat. Tiep tuc?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    setupSmartTaskManager();
    ui.alert("Setup hoan tat.");
  } catch (e) {
    ui.alert("Loi khi chay Setup: " + e.message);
  }
}