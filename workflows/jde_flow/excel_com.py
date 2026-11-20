#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 写入与重算。

模板包含图表、图片、数据透视表与外部链接，默认由 Excel COM 打开并保存，
以保留这些对象；COM 不可用时降级为 openpyxl 写入，并尽量回写外部链接部件。

同一份印花税计税表需要「写明细账 + 更新会计期间」和「回写批号」两次写入。
为避免 Excel 未完全退出时占用文件，默认复用同一个 Excel 会话完成两次写入。
"""
from __future__ import annotations

import shutil
import time
import zipfile
from copy import copy
from datetime import date, datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

from config.config_loader import ConfigError
from log.loggingConfig import logger

EXTERNAL_LINK_DIR = "xl/externalLinks/"
XL_PASTE_FORMATS = -4122
DATE_FORMAT = "yyyy/m/d"
# 一次清空过多行时 Excel COM 可能静默不生效，按块清空历史明细。
CLEAR_CHUNK_ROWS = 2000
EXCEL_EPOCH = datetime(1899, 12, 30)


class ExcelUnavailable(Exception):
    """本机无法使用 Excel COM，用于回退到 openpyxl 写入。"""


def external_link_parts(path):
    """读取外部链接部件；openpyxl 保存只写回部分关系，须在保存后回写。"""
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()
                if name.startswith(EXTERNAL_LINK_DIR) and not name.endswith("/")}


def restore_external_links(path, parts):
    """把原始外部链接部件写回，避免 rId 悬空导致办公软件无法打开生成文件。"""
    if not parts:
        return
    path = Path(path)
    temp = None
    try:
        with zipfile.ZipFile(path) as archive:
            entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
        with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as handle:
            temp = Path(handle.name)
        with zipfile.ZipFile(temp, "w") as archive:
            for item, data in entries:
                archive.writestr(item, parts.get(item.filename, data))
        temp.replace(path)
        temp = None
    finally:
        if temp and temp.exists():
            temp.unlink()


def _excel_serial(value):
    """Excel 日期序列号；直接传 pywintypes.Time 会按 UTC 处理、差 8 小时。"""
    moment = value if isinstance(value, datetime) else datetime(
        value.year, value.month, value.day)
    return (moment - EXCEL_EPOCH).total_seconds() / 86400


def label_value_coordinate(sheet, label):
    """返回标签右侧第一个可写单元格坐标；标签本身是合并区域时取合并区右侧。"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None or str(cell.value).strip() != label:
                continue
            for area in sheet.merged_cells.ranges:
                if (area.min_row <= cell.row <= area.max_row
                        and area.min_col <= cell.column <= area.max_col):
                    return sheet.cell(cell.row, area.max_col + 1).coordinate
            return sheet.cell(cell.row, cell.column + 1).coordinate
    raise ConfigError(f"{sheet.title} 未找到“{label}”单元格")


class ExcelSession:
    """一个 Excel 实例加一个已打开工作簿的会话，可连续写入并保存。"""

    def __init__(self, engine="Excel.Application"):
        self.engine = engine
        self.app = None
        self.book = None
        self._pythoncom = None
        self._sheets = {}

    # ---------- 打开与关闭 ----------
    def open_copy(self, template, destination):
        """复制模板到输出位置后打开，模板中的图表、透视表等对象原样保留。"""
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template, destination)
        self.open(destination)

    def open(self, path):
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelUnavailable(f"缺少 pywin32：{exc}") from exc
        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            try:
                self.app = win32com.client.DispatchEx(self.engine)
            except Exception as exc:
                raise ExcelUnavailable(f"无法启动 {self.engine}：{exc}") from exc
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.AskToUpdateLinks = False
            self.app.AutomationSecurity = 3
            self.book = self._open_book(Path(path))
            if self.book.ReadOnly:
                raise ConfigError("生成文件被占用，无法保存写入结果")
        except Exception:
            self.close()
            raise

    def _open_book(self, path):
        """新启的 Excel 实例偶发“Open 方法无效”，重试可覆盖启动竞态。"""
        last = None
        for _ in range(3):
            try:
                book = self.app.Workbooks.Open(
                    Filename=str(path),
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                    AddToMru=False,
                )
                if book is None:
                    raise RuntimeError("Open 未返回工作簿")
                return book
            except Exception as exc:
                last = exc
                time.sleep(1)
        raise ExcelUnavailable(f"Excel 无法打开 {path}：{last}")

    # ---------- 写入 ----------
    def sheet(self, name):
        """按名称取工作表并缓存；保存后 COM 动态绑定可能失效，缓存可避免再次解析。"""
        if name in self._sheets:
            return self._sheets[name]
        try:
            sheet = self.book.Worksheets(name)
        except Exception as exc:
            raise ConfigError(f"工作簿缺少工作表：{name}") from exc
        self._sheets[name] = sheet
        return sheet

    def write_ledger(self, sheet_name, header_row, column_pairs, rows):
        """按列映射写入明细账，保留模板其他列与行格式。

        模板数据区可能带自动筛选，被筛掉的行处于隐藏状态；先取消筛选并显示
        全部行，再写入，避免新数据被隐藏或清空不彻底。标题行保持不动。
        """
        if not rows:
            raise ConfigError("明细账没有数据，未生成输出")
        sheet = self.sheet(sheet_name)
        used = sheet.UsedRange
        last_row = max(used.Row + used.Rows.Count - 1, header_row)
        self._remove_filters(sheet, header_row, last_row)
        # 只清映射列的历史数据，保持模板其他列（公式等）不动。
        for target_col, _ in column_pairs:
            self._clear_column(sheet, target_col, header_row + 1, last_row)
        # 逐列写入：纯文本列先设为文本格式，避免 "00100" 被 Excel 转成 100。
        for target_col, source_col in column_pairs:
            values = [row[source_col - 1] if source_col - 1 < len(row) else None
                      for row in rows]
            target_range = sheet.Range(
                sheet.Cells(header_row + 1, target_col),
                sheet.Cells(header_row + len(rows), target_col),
            )
            text_column = all(value is None or isinstance(value, str) for value in values)
            original_format = target_range.NumberFormat if text_column else None
            if text_column:
                target_range.NumberFormat = "@"
            target_range.Value = [[value if value is not None else ""] for value in values]
            if text_column and original_format is not None:
                target_range.NumberFormat = original_format
        existing = last_row - header_row
        if len(rows) > existing:
            first_col = min(col for col, _ in column_pairs)
            last_col = max(col for col, _ in column_pairs)
            first = sheet.Range(sheet.Cells(header_row + 1, first_col),
                                sheet.Cells(header_row + 1, last_col))
            extra = sheet.Range(sheet.Cells(header_row + existing + 1, first_col),
                                sheet.Cells(header_row + len(rows), last_col))
            first.Copy()
            extra.PasteSpecial(XL_PASTE_FORMATS)
            self.app.CutCopyMode = False
            extra.EntireRow.RowHeight = first.EntireRow.RowHeight
        logger.info(f"{sheet_name} 已写入明细账 {len(rows)} 行")

    def _remove_filters(self, sheet, header_row, last_row):
        """取消工作表自动筛选并显示数据区所有行；标题行保持原样。"""
        try:
            if sheet.FilterMode:
                sheet.ShowAllData()
            if sheet.AutoFilterMode:
                sheet.AutoFilterMode = False
        except Exception as exc:
            logger.warning(f"{sheet.Name} 取消失效筛选失败：{exc}")
        if last_row > header_row:
            try:
                sheet.Rows(f"{header_row + 1}:{last_row}").Hidden = False
            except Exception as exc:
                logger.warning(f"{sheet.Name} 显示被筛选隐藏的行失败：{exc}")

    def _clear_column(self, sheet, column, first_row, last_row):
        """按块清空一列的数据区；一次性清空整列时 COM 可能静默不生效。"""
        letter = get_column_letter(int(column))
        for start in range(first_row, last_row + 1, CLEAR_CHUNK_ROWS):
            end = min(start + CLEAR_CHUNK_ROWS - 1, last_row)
            sheet.Range(f"{letter}{start}:{letter}{end}").ClearContents()

    def label_cell(self, sheet_name, label):
        """定位标签右侧的写入单元格；标签在合并区域内时取合并区右侧单元格。"""
        sheet = self.sheet(sheet_name)
        used = sheet.UsedRange
        first_row = used.Row
        last_row = used.Row + used.Rows.Count - 1
        first_col = used.Column
        last_col = used.Column + used.Columns.Count - 1
        for row in range(first_row, last_row + 1):
            for col in range(first_col, last_col + 1):
                value = sheet.Cells(row, col).Value
                if value is None or str(value).strip() != label:
                    continue
                area = sheet.Cells(row, col).MergeArea
                return sheet.Cells(row, area.Column + area.Columns.Count)
        raise ConfigError(f"{sheet_name} 未找到“{label}”单元格")

    def write_label_value(self, sheet_name, label, value, number_format=None):
        cell = self.label_cell(sheet_name, label)
        if value is None or (isinstance(value, str) and not value.strip()):
            cell.ClearContents()
        elif isinstance(value, (datetime, date)):
            cell.Value2 = _excel_serial(value)
        else:
            cell.Value = value
        if number_format:
            cell.NumberFormat = number_format
        address = f"{get_column_letter(int(cell.Column))}{int(cell.Row)}"
        logger.info(f"{sheet_name}!{address} 已更新“{label}”为 {value}")
        return address

    def calculate(self):
        """常规重算即可：本流程只改单元格值，不改公式。

        模板含透视表/外部链接时，CalculateFullRebuild 可能抛出 OLE 错误并让后续
        COM 调用异常，因此不使用全量重建。
        """
        self.app.Calculate()

    def save(self):
        self.book.Save()

    def close(self):
        """关闭工作簿并退出 Excel；关闭失败只记录日志，不覆盖业务异常。"""
        book, app = self.book, self.app
        self.book = self.app = None
        self._sheets.clear()
        try:
            if book is not None:
                try:
                    book.Close(SaveChanges=False)
                except Exception as exc:
                    logger.warning(f"关闭工作簿失败：{exc}")
        finally:
            try:
                if app is not None:
                    try:
                        app.Quit()
                    except Exception as exc:
                        logger.warning(f"退出 Excel 失败：{exc}")
            finally:
                if self._pythoncom is not None:
                    try:
                        self._pythoncom.CoUninitialize()
                    except Exception:
                        pass
                    self._pythoncom = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()


def _write_rows_with_openpyxl(sheet, header_row, column_pairs, rows):
    columns = sorted({col for col, _ in column_pairs})
    # 去掉模板数据区的自动筛选，并显示被筛掉的行，与 Excel 写入保持一致。
    sheet.auto_filter.ref = None
    for row in range(header_row + 1, sheet.max_row + 1):
        sheet.row_dimensions[row].hidden = False
    # 保存首条明细格式，扩展新行时延续日期、金额等显示格式。
    styles = {col: copy(sheet.cell(header_row + 1, col)._style) for col in columns}
    for col in columns:
        for row in range(header_row + 1, sheet.max_row + 1):
            sheet.cell(row, col).value = None
    for offset, row in enumerate(rows, 1):
        for col, source_col in column_pairs:
            cell = sheet.cell(header_row + offset, col)
            cell._style = copy(styles[col])
            value = row[source_col - 1] if source_col - 1 < len(row) else None
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"


def _write_labels_with_openpyxl(book, labels):
    for sheet_name, label, value, number_format in labels:
        if sheet_name not in book.sheetnames:
            raise ConfigError(f"模板缺少工作表：{sheet_name}")
        sheet = book[sheet_name]
        cell = sheet[label_value_coordinate(sheet, label)]
        cell.value = value
        if number_format:
            cell.number_format = number_format
        logger.info(f"{sheet_name}!{cell.coordinate} 已更新“{label}”为 {value}")


def _save_with_openpyxl(book, destination, links):
    temp_path = None
    try:
        with NamedTemporaryFile(dir=Path(destination).parent, suffix=".xlsx",
                                delete=False) as tmp:
            temp_path = Path(tmp.name)
        book.save(temp_path)
        restore_external_links(temp_path, links)
        temp_path.replace(destination)
        temp_path = None
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


def write_ledger_with_openpyxl(template, destination, sheet_name, header_row,
                               column_pairs, rows, labels=()):
    """无 Excel 时的降级写入：读取模板后另存，保留其他工作表与单元格样式。"""
    links = external_link_parts(template)
    book = openpyxl.load_workbook(template)
    try:
        if sheet_name not in book.sheetnames:
            raise ConfigError(f"模板缺少工作表：{sheet_name}")
        _write_rows_with_openpyxl(book[sheet_name], header_row, column_pairs, rows)
        _write_labels_with_openpyxl(book, labels)
        book.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        _save_with_openpyxl(book, destination, links)
    finally:
        book.close()


def write_ledger(template, destination, sheet_name, header_row, column_pairs, rows,
                 labels=(), engine="Excel.Application"):
    """复制模板、写入明细账与指定标签单元格并重算保存，返回仍打开的 Excel 会话。

    会话由调用方在完成后续写入（如回写批号）后关闭；使用 openpyxl 降级时返回 None。
    """
    engine = str(engine or "Excel.Application").strip()
    destination = Path(destination)
    if engine.lower() == "openpyxl":
        write_ledger_with_openpyxl(template, destination, sheet_name, header_row,
                                   column_pairs, rows, labels)
        return None
    session = None
    try:
        session = ExcelSession(engine)
        session.open_copy(template, destination)
        session.write_ledger(sheet_name, header_row, column_pairs, rows)
        for sheet_name_, label, value, number_format in labels:
            session.write_label_value(sheet_name_, label, value, number_format)
        session.calculate()
        session.save()
        if not destination.is_file() or destination.stat().st_size == 0:
            raise ConfigError(f"生成的输出文件无效：{destination}")
        logger.info(f"已复制模板并由 {engine} 写入明细并保存：{destination}")
    except ExcelUnavailable as exc:
        logger.warning(f"Excel 不可用，改用 openpyxl 写入模板：{exc}")
        if session is not None:
            session.close()
        _remove_file(destination)
        write_ledger_with_openpyxl(template, destination, sheet_name, header_row,
                                   column_pairs, rows, labels)
        return None
    except Exception:
        if session is not None:
            session.close()
        _remove_file(destination)
        raise
    kept, session = session, None
    return kept


def update_cells_with_openpyxl(path, writes):
    """无 Excel 时的降级写入：只改单元格值，保存后回写外部链接部件。"""
    links = external_link_parts(path)
    book = openpyxl.load_workbook(path)
    try:
        for sheet_name, label, value in writes:
            if sheet_name not in book.sheetnames:
                raise ConfigError(f"工作簿缺少工作表：{sheet_name}")
            sheet = book[sheet_name]
            sheet[label_value_coordinate(sheet, label)].value = value
            logger.info(f"{sheet_name} 已更新“{label}”为 {value}")
        book.calculation = CalcProperties(fullCalcOnLoad=True, forceFullCalc=True)
        _save_with_openpyxl(book, path, links)
    finally:
        book.close()


def update_cells(path, writes, engine="Excel.Application"):
    """单独更新若干标签单元格：优先 Excel，失败时降级 openpyxl。"""
    path = Path(path)
    if not path.is_file():
        raise ConfigError(f"待更新文件不存在：{path}")
    if str(engine).strip().lower() == "openpyxl":
        update_cells_with_openpyxl(path, writes)
        return
    session = None
    try:
        session = ExcelSession(engine)
        session.open(path)
        for sheet_name, label, value in writes:
            session.write_label_value(sheet_name, label, value)
        session.calculate()
        session.save()
    except ExcelUnavailable as exc:
        logger.warning(f"Excel 不可用，改用 openpyxl 更新单元格：{exc}")
        if session is not None:
            session.close()
            session = None
        update_cells_with_openpyxl(path, writes)
    finally:
        if session is not None:
            session.close()


def _remove_file(path):
    try:
        if Path(path).exists():
            Path(path).unlink()
    except OSError:
        pass
