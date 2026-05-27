# -*- coding: utf-8 -*-
"""Desktop window UI for the monthly report tools."""

from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import app as backend


DEFAULT_START = "2026-04-20"
DEFAULT_END = "2026-05-19"


class FileField(ttk.Frame):
    def __init__(self, master, label: str, filetypes: list[tuple[str, str]], *, save_dir: bool = False):
        super().__init__(master)
        self.filetypes = filetypes
        self.save_dir = save_dir
        self.value = tk.StringVar()

        ttk.Label(self, text=label).grid(row=0, column=0, sticky="w", pady=(0, 6))
        row = ttk.Frame(self)
        row.grid(row=1, column=0, sticky="ew")
        row.columnconfigure(0, weight=1)
        self.entry = ttk.Entry(row, textvariable=self.value)
        self.entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(row, text="选择", command=self.browse).grid(row=0, column=1)
        self.columnconfigure(0, weight=1)

    def browse(self) -> None:
        if self.save_dir:
            selected = filedialog.askdirectory(
                title="选择文件夹",
                initialdir=self.value.get() or str(backend.DEFAULT_OUTPUT_DIR),
                mustexist=False,
            )
        else:
            selected = filedialog.askopenfilename(title="选择文件", filetypes=self.filetypes)
        if selected:
            self.value.set(selected)

    def path(self) -> Path:
        text = self.value.get().strip().strip('"')
        if not text:
            raise ValueError("请先选择文件或文件夹")
        return Path(text)


class DateRange(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.start = tk.StringVar(value=DEFAULT_START)
        self.end = tk.StringVar(value=DEFAULT_END)
        ttk.Label(self, text="开始日期").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(self, text="结束日期").grid(row=0, column=1, sticky="w", padx=(14, 0), pady=(0, 6))
        ttk.Entry(self, textvariable=self.start, width=18).grid(row=1, column=0, sticky="ew")
        ttk.Entry(self, textvariable=self.end, width=18).grid(row=1, column=1, sticky="ew", padx=(14, 0))
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

    def values(self):
        start = backend.parse_ui_date(self.start.get())
        end = backend.parse_ui_date(self.end.get(), end_of_day=True)
        if start > end:
            raise ValueError("开始日期不能晚于结束日期")
        return start, end


class StatusBox(ttk.Frame):
    def __init__(self, master, initial: str):
        super().__init__(master)
        self.text = tk.Text(self, height=5, wrap="word", relief="solid", borderwidth=1)
        self.text.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.set(initial)

    def set(self, message: str) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", message)
        self.text.configure(state="disabled")


def add_field(parent, row: int, widget: ttk.Frame) -> int:
    widget.grid(row=row, column=0, sticky="ew", pady=(0, 14))
    return row + 1


def run_async(root: tk.Misc, button: ttk.Button, status: StatusBox, action, working_text: str) -> None:
    button.configure(state="disabled")
    status.set(working_text)

    def worker() -> None:
        try:
            message = action()
        except Exception as exc:
            error_text = str(exc)
            root.after(0, lambda: (status.set(f"处理失败：{error_text}"), messagebox.showerror("处理失败", error_text)))
        else:
            root.after(0, lambda: status.set(message))
        finally:
            root.after(0, lambda: button.configure(state="normal"))

    threading.Thread(target=worker, daemon=True).start()


def open_output(path: Path) -> None:
    target = path if path.is_dir() else path.parent
    try:
        if target.exists():
            os.startfile(target)
    except OSError:
        pass


def make_scrolled_frame(parent):
    canvas = tk.Canvas(parent, highlightthickness=0)
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    frame = ttk.Frame(canvas, padding=18)
    frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
    window_id = canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(0, weight=1)
    return frame


def build_summary_tab(parent, root: tk.Misc):
    parent.columnconfigure(0, weight=1)
    template = FileField(parent, "汇总 xlsx 模板", [("Excel 文件", "*.xlsx")])
    source = FileField(parent, "数据源台账", [("Excel 文件", "*.xlsx")])
    dates = DateRange(parent)
    output_dir = FileField(parent, "文件保存地址", [("文件夹", "*")], save_dir=True)
    output_dir.value.set(str(backend.DEFAULT_OUTPUT_DIR))
    status = StatusBox(parent, "等待选择文件。")
    button = ttk.Button(parent, text="生成汇总表")

    row = 0
    for widget in (template, source, dates, output_dir, status):
        row = add_field(parent, row, widget)
    button.grid(row=row, column=0, sticky="w")

    def action() -> str:
        start, end = dates.values()
        out_dir = backend.resolve_output_dir(output_dir.value.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        output_file = out_dir / backend.output_name_for(end)
        result = backend.generate_summary(source.path(), template.path(), start, end, output_file)
        lines = [f"生成完成：{output_file.name}", f"保存位置：{output_file}"]
        if result.get("row_counts"):
            lines.append("")
            lines.append("生成结果：")
            lines.extend(f"{name}：{count}" for name, count in result["row_counts"].items())
        if result.get("missing"):
            lines.append("")
            lines.append("缺失项：")
            lines.extend(str(item) for item in result["missing"])
        open_output(output_file)
        return "\n".join(lines)

    button.configure(command=lambda: run_async(root, button, status, action, "正在生成汇总表，请稍候..."))


def build_purity_tab(parent, root: tk.Misc):
    parent.columnconfigure(0, weight=1)
    source = FileField(parent, "上传区：带容器总数的原始台账", [("Excel 文件", "*.xlsx *.xls")])
    summary = FileField(parent, "待修改文件区：已生成的月报汇总表", [("Excel 文件", "*.xlsx")])
    dates = DateRange(parent)
    output_dir = FileField(parent, "输出区：更新后汇总表保存地址", [("文件夹", "*")], save_dir=True)
    output_dir.value.set(str(backend.DEFAULT_OUTPUT_DIR))
    status = StatusBox(parent, "等待上传原始台账和待修改汇总表。")
    hint = ttk.Label(
        parent,
        text="支持 xlsx / xls。后台会识别“有1个垃圾桶、1组桶、1个桶”等具体问题列，并生成容器总数与纯净率对照表。",
        wraplength=720,
        foreground="#667085",
    )
    button = ttk.Button(parent, text="更新容器总数与纯净率")

    row = 0
    for widget in (source, summary, dates, output_dir):
        row = add_field(parent, row, widget)
    hint.grid(row=row, column=0, sticky="w", pady=(0, 14))
    row += 1
    status.grid(row=row, column=0, sticky="ew", pady=(0, 14))
    row += 1
    button.grid(row=row, column=0, sticky="w")

    def action() -> str:
        start, end = dates.values()
        out_dir = backend.resolve_output_dir(output_dir.value.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        backend.PROCESS_DIR.mkdir(parents=True, exist_ok=True)
        output_file = out_dir / summary.path().name
        process_file = backend.PROCESS_DIR / "容器总数与纯净率对照表.xlsx"
        bad_counts = backend.read_summary_bad_counts(summary.path())
        comparison = backend.build_container_comparison(source.path(), start, end, process_file, bad_counts)
        updated = backend.update_summary_container_and_purity(summary.path(), output_file, comparison)
        open_output(output_file)
        return (
            f"更新完成：{output_file.name}\n"
            f"汇总表位置：{output_file}\n"
            f"对照表位置：{process_file}\n"
            f"识别街道数：{len(comparison)}\n"
            f"更新容器总数行：{updated['container_rows']}，更新纯净率行：{updated['purity_rows']}"
        )

    button.configure(command=lambda: run_async(root, button, status, action, "正在生成对照表并更新汇总表，请稍候..."))


def build_city_tab(parent, root: tk.Misc):
    parent.columnconfigure(0, weight=1)
    summary = FileField(parent, "待修改文件区：已生成的月报汇总表", [("Excel 文件", "*.xlsx")])
    resident = FileField(parent, "市级小区汇总表", [("Excel 文件", "*.xlsx")])
    social = FileField(parent, "市级社会汇总表", [("Excel 文件", "*.xlsx")])
    dates = DateRange(parent)
    output_dir = FileField(parent, "输出区：更新后汇总表保存地址", [("文件夹", "*")], save_dir=True)
    output_dir.value.set(str(backend.DEFAULT_OUTPUT_DIR))
    status = StatusBox(parent, "等待上传月汇总表和两个市级汇总表。")
    button = ttk.Button(parent, text="更新市级检查情况")

    row = 0
    for widget in (summary, resident, social, dates, output_dir, status):
        row = add_field(parent, row, widget)
    button.grid(row=row, column=0, sticky="w")

    def action() -> str:
        start, end = dates.values()
        out_dir = backend.resolve_output_dir(output_dir.value.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        output_file = out_dir / summary.path().name
        resident_stats = backend.read_city_summary_stats(resident.path(), start, end)
        social_stats = backend.read_city_summary_stats(social.path(), start, end)
        updated = backend.update_summary_city_checks(summary.path(), output_file, resident_stats, social_stats)
        open_output(output_file)
        return (
            f"更新完成：{output_file.name}\n"
            f"汇总表位置：{output_file}\n"
            f"更新市级行：{updated['city_rows']}\n"
            f"小区表街道数：{len(resident_stats)}，社会表街道数：{len(social_stats)}\n"
            f"时间范围：{start:%Y-%m-%d} 至 {end:%Y-%m-%d}"
        )

    button.configure(command=lambda: run_async(root, button, status, action, "正在读取市级汇总表并更新月汇总表，请稍候..."))


def build_transfer_tab(parent, root: tk.Misc):
    parent.columnconfigure(0, weight=1)
    ledger = FileField(parent, "上传区：中转站台账", [("Excel 文件", "*.xlsx *.xls")])
    summary = FileField(parent, "待修改文件区：已生成的月报汇总表", [("Excel 文件", "*.xlsx")])
    dates = DateRange(parent)
    output_dir = FileField(parent, "输出区：更新后汇总表保存地址", [("文件夹", "*")], save_dir=True)
    output_dir.value.set(str(backend.DEFAULT_OUTPUT_DIR))
    status = StatusBox(parent, "等待上传中转站台账和待修改汇总表。")
    hint = ttk.Label(
        parent,
        text="按所选时间段筛选中转站台账日期，并按街道更新汇总表中的“中转站”相关行。",
        wraplength=720,
        foreground="#667085",
    )
    button = ttk.Button(parent, text="更新中转站情况")

    row = 0
    for widget in (ledger, summary, dates, output_dir):
        row = add_field(parent, row, widget)
    hint.grid(row=row, column=0, sticky="w", pady=(0, 14))
    row += 1
    status.grid(row=row, column=0, sticky="ew", pady=(0, 14))
    row += 1
    button.grid(row=row, column=0, sticky="w")

    def action() -> str:
        start, end = dates.values()
        out_dir = backend.resolve_output_dir(output_dir.value.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        output_file = out_dir / summary.path().name
        transfer_stats = backend.read_transfer_station_stats(ledger.path(), start, end)
        updated = backend.update_summary_transfer_station(summary.path(), output_file, transfer_stats)
        open_output(output_file)
        return (
            f"更新完成：{output_file.name}\n"
            f"汇总表位置：{output_file}\n"
            f"更新中转站行：{updated['transfer_rows']}\n"
            f"识别街道数：{len(transfer_stats)}\n"
            f"时间范围：{start:%Y-%m-%d} 至 {end:%Y-%m-%d}"
        )

    button.configure(command=lambda: run_async(root, button, status, action, "正在读取中转站台账并更新月汇总表，请稍候..."))


def build_report_window(root: tk.Tk) -> None:
    win = tk.Toplevel(root)
    win.title("生成工作报告")
    win.geometry("820x680")
    win.transient(root)
    frame = make_scrolled_frame(win)
    frame.columnconfigure(0, weight=1)

    ttk.Label(frame, text="生成两个 Word 报告", font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 16))
    summary = FileField(frame, "各街道问题汇总表", [("Excel 文件", "*.xlsx")])
    source = FileField(frame, "源台账", [("Excel 文件", "*.xlsx")])
    residential_template = FileField(frame, "垃圾分类检查自查报告模板", [("Word 文件", "*.docx")])
    social_template = FileField(frame, "社会单位、餐饮单位检查报告模板", [("Word 文件", "*.docx")])
    dates = DateRange(frame)
    output_dir = FileField(frame, "报告保存地址", [("文件夹", "*")], save_dir=True)
    output_dir.value.set(str(backend.DEFAULT_OUTPUT_DIR))
    status = StatusBox(frame, "等待上传报告模板。")
    button = ttk.Button(frame, text="生成两个报告")

    row = 1
    for widget in (summary, source, residential_template, social_template, dates, output_dir, status):
        row = add_field(frame, row, widget)
    button.grid(row=row, column=0, sticky="w")

    def action() -> str:
        start, end = dates.values()
        out_dir = backend.resolve_output_dir(output_dir.value.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        result = backend.generate_reports(
            summary.path(),
            source.path(),
            residential_template.path(),
            social_template.path(),
            start,
            end,
            out_dir,
        )
        open_output(out_dir)
        files = "\n".join(str(path) for path in result["output_files"])
        message = f"生成完成：{len(result['output_files'])} 个报告\n保存位置：{out_dir}\n\n{files}"
        if result.get("missing"):
            message += "\n\n缺失项：\n" + "\n".join(str(item) for item in result["missing"])
        return message

    button.configure(command=lambda: run_async(root, button, status, action, "正在上传并生成两个报告，请稍候..."))


def build_summary_window(root: tk.Tk) -> None:
    win = tk.Toplevel(root)
    win.title("生成汇总表")
    win.geometry("920x740")
    win.minsize(760, 560)
    win.transient(root)

    shell = ttk.Frame(win, padding=16)
    shell.pack(fill="both", expand=True)
    shell.columnconfigure(0, weight=1)
    shell.rowconfigure(1, weight=1)

    header = ttk.Frame(shell)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
    header.columnconfigure(0, weight=1)
    ttk.Label(header, text="生成汇总表", font=("Microsoft YaHei UI", 15, "bold")).grid(row=0, column=0, sticky="w")
    ttk.Label(
        header,
        text="选择一个页签完成对应操作，原有生成、更新容器总数与纯净率、更新市级检查情况功能保持不变。",
        foreground="#667085",
        wraplength=760,
    ).grid(row=1, column=0, sticky="w", pady=(6, 0))

    notebook = ttk.Notebook(shell)
    notebook.grid(row=1, column=0, sticky="nsew")

    summary_tab = ttk.Frame(notebook)
    purity_tab = ttk.Frame(notebook)
    city_tab = ttk.Frame(notebook)
    transfer_tab = ttk.Frame(notebook)
    summary_frame = make_scrolled_frame(summary_tab)
    purity_frame = make_scrolled_frame(purity_tab)
    city_frame = make_scrolled_frame(city_tab)
    transfer_frame = make_scrolled_frame(transfer_tab)

    notebook.add(summary_tab, text="生成月报汇总表")
    notebook.add(purity_tab, text="更新容器总数与纯净率")
    notebook.add(city_tab, text="更新市级检查情况")
    notebook.add(transfer_tab, text="更新中转站情况")

    build_summary_tab(summary_frame, root)
    build_purity_tab(purity_frame, root)
    build_city_tab(city_frame, root)
    build_transfer_tab(transfer_frame, root)


def main() -> None:
    backend.WORK_DIR.mkdir(parents=True, exist_ok=True)
    backend.DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    root = tk.Tk()
    root.title("月报工具")
    root.geometry("620x360")
    root.minsize(560, 320)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")

    container = ttk.Frame(root, padding=28)
    container.pack(fill="both", expand=True)
    container.columnconfigure(0, weight=1)

    ttk.Label(container, text="月报工具", font=("Microsoft YaHei UI", 20, "bold")).grid(row=0, column=0, sticky="w")
    ttk.Label(
        container,
        text="选择要处理的任务，再上传对应模板和台账文件。",
        foreground="#667085",
    ).grid(row=1, column=0, sticky="w", pady=(8, 24))

    ttk.Button(container, text="生成汇总表", command=lambda: build_summary_window(root)).grid(row=2, column=0, sticky="ew", ipady=16, pady=(0, 14))
    ttk.Label(
        container,
        text="进入后可生成月报汇总表，也可以续写容器总数、纯净率和市级检查情况。",
        foreground="#667085",
        wraplength=540,
    ).grid(row=3, column=0, sticky="w", pady=(0, 18))

    ttk.Button(container, text="生成工作报告", command=lambda: build_report_window(root)).grid(row=4, column=0, sticky="ew", ipady=16, pady=(0, 14))
    ttk.Label(
        container,
        text="上传汇总表、源台账和两个 Word 模板，生成两个工作报告文档。",
        foreground="#667085",
        wraplength=540,
    ).grid(row=5, column=0, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    main()
