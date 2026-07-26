#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""周报合并：读取 Excel/CSV，并将各时间来源纵向追加到主表。"""

from __future__ import annotations

import csv
import queue
import re
import threading
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_BASE_FIELDS = (
    "仓店类型", "办事处", "店铺类别", "仓店名称", "仓店代码", "销售截止日期", "货号",
    "商品名称", "牌价", "大类", "系列", "大小童", "性别", "配货季",
)
DEFAULT_METRICS = ("零售数量", "结算金额", "零售吊牌额", "财务存", "财务存牌价")
SOURCE_ROLES = ("同期月", "本周", "上周", "同期周")
SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv"}


@dataclass
class TableData:
    path: Path
    rows: list[list[object]]
    header_index: int
    header_map: dict[str, list[int]]


@dataclass
class MergeResult:
    output_path: Path
    source_counts: dict[str, int]


def normalize_header(value: object) -> str:
    """统一表头比较形式，忽略空白和换行。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def parse_field_text(text: str) -> list[str]:
    """解析按行、逗号或制表符输入的动态字段。"""
    parts = re.split(r"[\n,，\t]+", text)
    fields: list[str] = []
    seen: set[str] = set()
    for part in parts:
        field = part.strip()
        normalized = normalize_header(field)
        if normalized and normalized not in seen:
            fields.append(field)
            seen.add(normalized)
    return fields


def _read_xlsx(path: Path) -> list[list[object]]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        return [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def _read_xls(path: Path) -> list[list[object]]:
    import xlrd

    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    rows: list[list[object]] = []
    for row_index in range(sheet.nrows):
        row: list[object] = []
        for column_index in range(sheet.ncols):
            cell = sheet.cell(row_index, column_index)
            if cell.ctype == xlrd.XL_CELL_DATE:
                row.append(xlrd.xldate.xldate_as_datetime(cell.value, workbook.datemode))
            else:
                row.append(cell.value)
        rows.append(row)
    return rows


def _read_csv(path: Path) -> list[list[object]]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with path.open("r", encoding=encoding, newline="") as file:
                sample = file.read(8192)
                file.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
                except csv.Error:
                    dialect = csv.excel
                return [list(row) for row in csv.reader(file, dialect)]
        except UnicodeDecodeError as error:
            last_error = error
    raise ValueError(f"无法识别 CSV 编码：{path.name}") from last_error


def read_rows(path: str | Path) -> list[list[object]]:
    file_path = Path(path)
    if not file_path.exists():
        raise ValueError(f"文件不存在：{file_path}")
    suffix = file_path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ValueError(f"不支持的文件格式：{file_path.name}")
    if suffix == ".xlsx":
        rows = _read_xlsx(file_path)
    elif suffix == ".xls":
        rows = _read_xls(file_path)
    else:
        rows = _read_csv(file_path)
    if not rows:
        raise ValueError(f"文件没有数据：{file_path.name}")
    return rows


def build_header_map(row: Iterable[object]) -> dict[str, list[int]]:
    header_map: dict[str, list[int]] = defaultdict(list)
    for index, value in enumerate(row):
        normalized = normalize_header(value)
        if normalized:
            header_map[normalized].append(index)
    return dict(header_map)


def detect_header(rows: list[list[object]], fields: list[str], file_name: str) -> tuple[int, dict[str, list[int]]]:
    required = {normalize_header(field) for field in fields}
    best_index = -1
    best_map: dict[str, list[int]] = {}
    best_score = -1
    for index, row in enumerate(rows[:100]):
        header_map = build_header_map(row)
        score = len(required.intersection(header_map))
        if score > best_score:
            best_index, best_map, best_score = index, header_map, score
        if required.issubset(header_map):
            return index, header_map
    missing = sorted(required.difference(best_map))
    matched = f"已匹配 {best_score}/{len(required)} 个字段"
    raise ValueError(f"{file_name} 无法识别完整表头（{matched}），缺少：{'、'.join(missing)}")


def load_table(path: str | Path, base_fields: list[str], metrics: list[str]) -> TableData:
    file_path = Path(path)
    rows = read_rows(file_path)
    header_index, header_map = detect_header(rows, base_fields + metrics, file_path.name)
    return TableData(file_path, rows, header_index, header_map)


def _column_index(table: TableData, field: str, use_last: bool) -> int:
    indexes = table.header_map[normalize_header(field)]
    return indexes[-1] if use_last else indexes[0]


def _has_values(row: list[object], indexes: Iterable[int]) -> bool:
    return any(index < len(row) and row[index] not in (None, "") for index in indexes)


def _cell_value(row: list[object], index: int) -> object:
    return row[index] if index < len(row) else None


def _last_content_row(rows: list[list[object]]) -> int:
    for index in range(len(rows) - 1, -1, -1):
        if any(value not in (None, "") for value in rows[index]):
            return index + 1
    return 1


def _copy_header_style(sheet, source_column: int, target_column: int, header_row: int) -> None:
    source = sheet.cell(header_row, source_column)
    target = sheet.cell(header_row, target_column)
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    if source.alignment:
        target.alignment = copy(source.alignment)
    if source.fill:
        target.fill = copy(source.fill)
    if source.font:
        target.font = copy(source.font)
    if source.border:
        target.border = copy(source.border)


def _create_output_workbook(main: TableData):
    from openpyxl import Workbook, load_workbook

    if main.path.suffix.lower() == ".xlsx":
        workbook = load_workbook(main.path, data_only=False)
        return workbook, workbook.active
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "周报合并"
    for row_index, row in enumerate(main.rows, start=1):
        for column_index, value in enumerate(row, start=1):
            sheet.cell(row_index, column_index, value)
    return workbook, sheet


def merge_reports(
    main_path: str | Path,
    source_paths: dict[str, str | Path],
    base_fields: list[str],
    metrics: list[str],
    output_path: str | Path,
    log: Callable[[str], None] | None = None,
    progress: Callable[[int, int], None] | None = None,
) -> MergeResult:
    """保留主表，并将各来源行写入带角色前缀的动态指标列。"""
    logger = log or (lambda _message: None)
    if not base_fields:
        raise ValueError("基础字段不能为空")
    if not metrics:
        raise ValueError("追加指标不能为空")
    selected_sources = [(role, Path(source_paths[role])) for role in SOURCE_ROLES if source_paths.get(role)]
    if not selected_sources:
        raise ValueError("请至少选择一个追加文件")

    output = Path(output_path)
    input_paths = [Path(main_path), *(path for _, path in selected_sources)]
    if any(output.resolve() == path.resolve() for path in input_paths):
        raise ValueError("输出文件不能覆盖任何输入文件")

    logger("正在读取主文件……")
    main = load_table(main_path, base_fields, metrics)
    sources: list[tuple[str, TableData]] = []
    for role, path in selected_sources:
        logger(f"正在读取{role}：{path.name}")
        sources.append((role, load_table(path, base_fields, metrics)))

    workbook, sheet = _create_output_workbook(main)
    try:
        original_width = max(max((len(row) for row in main.rows), default=0), sheet.max_column)
        header_row = main.header_index + 1
        source_style_column = max(1, original_width)
        output_columns: dict[tuple[str, str], int] = {}
        next_column = original_width + 1
        for role, _table in sources:
            for metric in metrics:
                output_columns[(role, normalize_header(metric))] = next_column
                cell = sheet.cell(header_row, next_column, f"{role}{metric.strip()}")
                _copy_header_style(sheet, source_style_column, next_column, header_row)
                cell.value = f"{role}{metric.strip()}"
                next_column += 1

        main_base_columns = {
            normalize_header(field): _column_index(main, field, use_last=False) + 1
            for field in base_fields
        }
        append_row = _last_content_row(main.rows) + 1
        total_rows = sum(
            1
            for _role, table in sources
            for row in table.rows[table.header_index + 1:]
            if _has_values(
                row,
                [_column_index(table, field, False) for field in base_fields]
                + [_column_index(table, metric, True) for metric in metrics],
            )
        )
        completed = 0
        counts: dict[str, int] = {}
        for role, table in sources:
            base_indexes = {normalize_header(field): _column_index(table, field, False) for field in base_fields}
            metric_indexes = {normalize_header(metric): _column_index(table, metric, True) for metric in metrics}
            source_indexes = list(base_indexes.values()) + list(metric_indexes.values())
            count = 0
            for row in table.rows[table.header_index + 1:]:
                if not _has_values(row, source_indexes):
                    continue
                for field_name, source_index in base_indexes.items():
                    sheet.cell(append_row, main_base_columns[field_name], _cell_value(row, source_index))
                for metric_name, source_index in metric_indexes.items():
                    target_column = output_columns[(role, metric_name)]
                    sheet.cell(append_row, target_column, _cell_value(row, source_index))
                append_row += 1
                count += 1
                completed += 1
                if progress:
                    progress(completed, max(total_rows, 1))
            counts[role] = count
            logger(f"{role}：已追加 {count} 行")

        output.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output)
        logger(f"合并完成：{output}")
        return MergeResult(output, counts)
    finally:
        workbook.close()


class WeeklyReportWindow:
    """独立周报合并窗口。"""

    def __init__(self, parent: tk.Misc):
        self.window = tk.Toplevel(parent)
        self.window.title("周报合并")
        self.window.geometry("920x760")
        self.window.minsize(780, 650)
        self.messages: queue.Queue[dict[str, object]] = queue.Queue()
        self.file_vars = {"周报本月": tk.StringVar(), **{role: tk.StringVar() for role in SOURCE_ROLES}}
        self._build_ui()
        self._poll_messages()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.window, padding=16)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="周报合并", font=("Arial", 18, "bold")).pack(anchor=tk.W)
        ttk.Label(
            container,
            text="主文件完整保留；追加文件按角色扩展指标列，并将数据行追加到主表下方。",
        ).pack(anchor=tk.W, pady=(2, 10))

        files_frame = ttk.LabelFrame(container, text="文件选择（主文件必选，其他文件可选）", padding=10)
        files_frame.pack(fill=tk.X)
        labels = ("周报本月",) + SOURCE_ROLES
        for row, role in enumerate(labels):
            label = "主文件（周报本月）" if role == "周报本月" else role
            ttk.Label(files_frame, text=label, width=18).grid(row=row, column=0, sticky=tk.W, pady=3)
            ttk.Entry(files_frame, textvariable=self.file_vars[role]).grid(
                row=row, column=1, sticky=tk.EW, padx=6, pady=3
            )
            ttk.Button(files_frame, text="选择…", command=lambda r=role: self._choose_file(r)).grid(
                row=row, column=2, pady=3
            )
            ttk.Button(files_frame, text="清除", command=lambda r=role: self.file_vars[r].set("")).grid(
                row=row, column=3, padx=(5, 0), pady=3
            )
        files_frame.columnconfigure(1, weight=1)

        config_frame = ttk.LabelFrame(container, text="字段配置（每行一个，也支持逗号分隔）", padding=10)
        config_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        left = ttk.Frame(config_frame)
        right = ttk.Frame(config_frame)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        ttk.Label(left, text="基础字段（默认14个）").pack(anchor=tk.W)
        self.base_text = tk.Text(left, height=14, wrap=tk.NONE)
        self.base_text.pack(fill=tk.BOTH, expand=True)
        ttk.Label(right, text="追加指标（默认5个，可动态增删）").pack(anchor=tk.W)
        self.metric_text = tk.Text(right, height=14, wrap=tk.NONE)
        self.metric_text.pack(fill=tk.BOTH, expand=True)
        self._restore_defaults()

        config_buttons = ttk.Frame(container)
        config_buttons.pack(fill=tk.X)
        ttk.Button(config_buttons, text="恢复默认字段", command=self._restore_defaults).pack(side=tk.LEFT)
        ttk.Label(config_buttons, text="同名指标重复时取最后一列（例如财务存）").pack(side=tk.LEFT, padx=12)

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill=tk.X, pady=(10, 5))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.progress_label = ttk.Label(progress_frame, text="就绪", width=12, anchor=tk.E)
        self.progress_label.pack(side=tk.RIGHT, padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="处理日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(container)
        actions.pack(fill=tk.X, before=files_frame, pady=(0, 10))
        self.merge_button = ttk.Button(actions, text="开始合并", command=self._start_merge)
        self.merge_button.pack(side=tk.RIGHT)
        ttk.Button(actions, text="关闭", command=self.window.destroy).pack(side=tk.RIGHT, padx=8)

    def _choose_file(self, role: str) -> None:
        path = filedialog.askopenfilename(
            parent=self.window,
            title=f"选择{role}文件",
            filetypes=[("周报文件", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")],
        )
        if path:
            self.file_vars[role].set(path)

    def _restore_defaults(self) -> None:
        self.base_text.delete("1.0", tk.END)
        self.base_text.insert("1.0", "\n".join(DEFAULT_BASE_FIELDS))
        self.metric_text.delete("1.0", tk.END)
        self.metric_text.insert("1.0", "\n".join(DEFAULT_METRICS))

    def _start_merge(self) -> None:
        main_path = self.file_vars["周报本月"].get().strip()
        sources = {role: self.file_vars[role].get().strip() for role in SOURCE_ROLES}
        base_fields = parse_field_text(self.base_text.get("1.0", tk.END))
        metrics = parse_field_text(self.metric_text.get("1.0", tk.END))
        if not main_path:
            messagebox.showwarning("缺少主文件", "请选择主文件（周报本月）。", parent=self.window)
            return
        if not any(sources.values()):
            messagebox.showwarning("缺少追加文件", "请至少选择一个追加文件。", parent=self.window)
            return
        if not base_fields or not metrics:
            messagebox.showwarning("字段为空", "基础字段和追加指标都不能为空。", parent=self.window)
            return

        output_path = filedialog.asksaveasfilename(
            parent=self.window,
            title="保存周报合并结果",
            initialdir=str(Path(main_path).parent),
            initialfile="周报合并结果.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not output_path:
            return
        self.merge_button.configure(state=tk.DISABLED)
        self.progress.configure(value=0, maximum=1)
        self.progress_label.configure(text="准备中")
        self._append_log("开始合并……")
        worker = threading.Thread(
            target=self._merge_worker,
            args=(main_path, sources, base_fields, metrics, output_path),
            daemon=True,
        )
        worker.start()

    def _merge_worker(
        self,
        main_path: str,
        sources: dict[str, str],
        base_fields: list[str],
        metrics: list[str],
        output_path: str,
    ) -> None:
        try:
            result = merge_reports(
                main_path,
                sources,
                base_fields,
                metrics,
                output_path,
                log=lambda text: self.messages.put({"type": "log", "text": text}),
                progress=lambda value, maximum: self.messages.put(
                    {"type": "progress", "value": value, "maximum": maximum}
                ),
            )
            self.messages.put({"type": "done", "result": result})
        except Exception as error:
            self.messages.put({"type": "error", "text": str(error)})

    def _poll_messages(self) -> None:
        if not self.window.winfo_exists():
            return
        try:
            while True:
                message = self.messages.get_nowait()
                message_type = message["type"]
                if message_type == "log":
                    self._append_log(str(message["text"]))
                elif message_type == "progress":
                    maximum = int(message["maximum"])
                    value = int(message["value"])
                    self.progress.configure(maximum=maximum, value=value)
                    self.progress_label.configure(text=f"{value}/{maximum}")
                elif message_type == "done":
                    self._handle_done(message["result"])
                elif message_type == "error":
                    self._handle_error(str(message["text"]))
        except queue.Empty:
            pass
        self.window.after(100, self._poll_messages)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _handle_done(self, result: object) -> None:
        assert isinstance(result, MergeResult)
        self.merge_button.configure(state=tk.NORMAL)
        counts = "\n".join(f"{role}：{count} 行" for role, count in result.source_counts.items())
        self.progress_label.configure(text="完成")
        messagebox.showinfo(
            "合并完成",
            f"周报合并成功。\n\n{counts}\n\n输出文件：\n{result.output_path}",
            parent=self.window,
        )

    def _handle_error(self, text: str) -> None:
        self.merge_button.configure(state=tk.NORMAL)
        self.progress_label.configure(text="失败")
        self._append_log(f"错误：{text}")
        messagebox.showerror("合并失败", text, parent=self.window)
