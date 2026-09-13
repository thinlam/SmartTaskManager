import { setupSmartTaskManager } from "../core/setup";
import { addTask } from "../core/tasks";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("setup-button")!.onclick = runSetup;
        document.getElementById("add-task-button")!.onclick = runAddDemoTask;
    }
});

async function runSetup(): Promise<void> {
    const status = document.getElementById("status")!;
    status.textContent = "Đang khởi tạo...";
    try {
        await setupSmartTaskManager();
        status.textContent = "Khởi tạo thành công! Kiểm tra các sheet mới.";
    } catch (error) {
        status.textContent = "Lỗi: " + (error as Error).message;
        console.error(error);
    }
}

async function runAddDemoTask(): Promise<void> {
    const status = document.getElementById("status")!;
    status.textContent = "Đang thêm Task...";
    try {
        const task = await addTask({
            taskName: "Fix authentication bug",
            priority: "Critical",
            status: "In Progress",
            progress: 10,
            dueDate: "2026-09-10",
        });
        status.textContent = `Đã thêm ${task.taskId} - SmartScore: ${task.smartScore}`;
    } catch (error) {
        status.textContent = "Lỗi: " + (error as Error).message;
        console.error(error);
    }
}