from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from update_st_close import process_excel_file


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ST 收盘价更新")
        self.root.geometry("560x180")
        self.root.resizable(False, False)

        self.file_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择 Excel 文件")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = tk.Frame(self.root, padx=16, pady=16)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Excel 文件:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.file_var, width=55).grid(row=1, column=0, padx=(0, 8), pady=(6, 10))
        tk.Button(frame, text="选择文件", width=10, command=self.choose_file).grid(row=1, column=1, pady=(6, 10))

        self.confirm_btn = tk.Button(frame, text="确定", width=10, command=self.run_process)
        self.confirm_btn.grid(row=2, column=0, sticky="w")

        tk.Button(frame, text="退出", width=10, command=self.root.destroy).grid(row=2, column=1, sticky="e")

        tk.Label(frame, textvariable=self.status_var, fg="#444").grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls")],
        )
        if path:
            self.file_var.set(path)
            self.status_var.set("已选择文件，点击确定开始处理")

    def run_process(self) -> None:
        path_str = self.file_var.get().strip()
        if not path_str:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return

        excel_path = Path(path_str)
        if not excel_path.exists():
            messagebox.showerror("错误", f"文件不存在:\n{excel_path}")
            return

        self.confirm_btn.config(state=tk.DISABLED)
        self.status_var.set("处理中，请稍候...")
        worker = threading.Thread(target=self._do_process, args=(excel_path,), daemon=True)
        worker.start()

    def _do_process(self, excel_path: Path) -> None:
        try:
            out = process_excel_file(excel_path)
            self.root.after(0, self._on_success, out)
        except Exception as e:  # noqa: BLE001
            self.root.after(0, self._on_error, e)

    def _on_success(self, output_path: Path) -> None:
        self.confirm_btn.config(state=tk.NORMAL)
        self.status_var.set(f"完成: {output_path}")
        messagebox.showinfo("完成", f"处理完成，输出文件:\n{output_path}")

    def _on_error(self, err: Exception) -> None:
        self.confirm_btn.config(state=tk.NORMAL)
        self.status_var.set("处理失败")
        messagebox.showerror("处理失败", str(err))


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
