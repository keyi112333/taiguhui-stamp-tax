#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""印花税流程：导出明细账、写入模板并更新会计期间、录入分录并回写批号。"""
import calendar
import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl

from config.config_loader import ConfigError
from log.loggingConfig import logger
from utils.helpers import get_timestamp
from web.JDE.jde_function import create_journal_entry, download_gl_movement
from workflows.jde_flow import excel_com

LEDGER_SHEET = "JDE明细账"
CALC_SHEET = "印花税计算底稿"
# 需要更新“会计期间：”与“批号：”的页签：计算底稿 + 租赁 + 收入类。
PERIOD_SHEETS = ("租赁", "收入类")
PERIOD_LABEL = "会计期间："
BATCH_LABEL = "批号："
JOURNAL_TITLES = ("账号", "借项金额", "贷项金额", "备注")


def normalize_header_name(value):
    """标题规范化：忽略空白、下划线、连字符、加号、斜杠与括号。"""
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[\s_\-+/()（）]+", "", text)


def build_header_pairs(source_titles, target_titles, explicit_mapping=None):
    """按规范化标题构建 源标题 -> 模板标题 映射。"""
    source_by_norm = {normalize_header_name(title): title for title in source_titles}
    target_by_norm = {normalize_header_name(title): title for title in target_titles}
    pairs = {}
    mapping = {} if explicit_mapping is None else dict(explicit_mapping)

    def resolve_source(title):
        if title in source_by_norm:
            return source_by_norm[title]
        if title in source_titles:
            return title
        return source_by_norm.get(normalize_header_name(title))

    def resolve_target(title):
        if title in target_by_norm:
            return target_by_norm[title]
        if title in target_titles:
            return title
        return target_by_norm.get(normalize_header_name(title))

    for source_title, target_title in mapping.items():
        resolved_source = resolve_source(source_title)
        resolved_target = resolve_target(target_title)
        if resolved_source and resolved_target:
            pairs[resolved_source] = resolved_target

    for source_title in source_titles:
        if source_title in pairs:
            continue
        target_title = target_by_norm.get(normalize_header_name(source_title))
        if target_title and target_title not in pairs.values():
            pairs[source_title] = target_title

    for target_title in target_titles:
        if target_title in pairs.values():
            continue
        source_title = source_by_norm.get(normalize_header_name(target_title))
        if source_title and source_title not in pairs:
            pairs[source_title] = target_title

    return pairs


def format_jde_date(value, label):
    """格式化为 JDE 日期 MM/DD/YYYY。"""
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).strftime("%m/%d/%Y")
        except ValueError:
            continue
    raise ConfigError(f"{label} 日期格式不正确，应为 YYYY-MM-DD 或 MM/DD/YYYY 等可识别格式；当前值：{text}")


def configured_period(value):
    """会计期间配置值；空白（含只有空格）视为未配置，按自然月返回今天。"""
    text = "" if value is None else str(value).strip()
    return text or date.today()


def month_bounds(period):
    """返回会计期间自然月的第一天和最后一天。"""
    if isinstance(period, (datetime, date)):
        year, month = period.year, period.month
    else:
        text = "" if period is None else str(period).strip()
        match = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not match:
            raise ConfigError(f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 或 YYYYMM 等可识别格式；当前值：{text}")
        year, month = int(match.group(1)), int(match.group(2))
    if not 1 <= month <= 12:
        raise ConfigError(f"会计期间月份不正确，应为 1-12；当前值：{month}")
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def _journal_title(value):
    """分录区标题规范化：忽略空白并统一“帐/账”，不依赖标题所在行。"""
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).replace("帐", "账")


def journal_region(sheet):
    """按账号/借项金额/贷项金额/备注四个标题动态定位分录区。

    标题行可以不在第一行，四列也可以不相邻；数据区从表头下一行开始，
    到最后一个有账号的行结束（不按账号后缀筛选）。
    返回 (表头行, {标题: 列号}, 数据行号列表)。
    """
    expected = {_journal_title(title) for title in JOURNAL_TITLES}
    matches = []
    for row in sheet.iter_rows():
        found = {}
        for cell in row:
            title = _journal_title(cell.value)
            if title in expected:
                found[title] = cell.column
        if len(found) == len(expected):
            matches.append((row[0].row, found))
    if not matches:
        raise ConfigError("未找到 账号/借项金额/贷项金额/备注 标题行")
    if len(matches) > 1:
        raise ConfigError(f"无法唯一定位分录区表头：找到 {len(matches)} 处")
    header_row, columns = matches[0]
    account_column = columns["账号"]
    last_row = None
    for row in range(header_row + 1, sheet.max_row + 1):
        value = sheet.cell(row, account_column).value
        if value is not None and str(value).strip():
            last_row = row
    if last_row is None:
        return header_row, columns, []
    return header_row, columns, list(range(header_row + 1, last_row + 1))


def _zero_amount(value):
    """金额为空、空串或数值 0 时视为零，用于跳过借贷皆为 0 的行。"""
    if value is None:
        return True
    text = str(value).strip()
    if not text:
        return True
    try:
        return float(text) == 0
    except ValueError:
        return False


def header_columns(sheet, row):
    """读取指定行的标题与列号；标题重复时报错。"""
    result = {}
    for cell in sheet[row]:
        if cell.value is None or not str(cell.value).strip():
            continue
        title = str(cell.value).strip()
        if title in result:
            raise ConfigError(f"{sheet.title} 标题重复：{title}")
        result[title] = cell.column
    return result


class StampTaxFlow:
    """印花税单行流程。"""

    def __init__(self, context):
        self.context = context
        self.config = context.config
        self.session = None

    def get(self, key, default=None):
        return self.config.get(f"印花税配置.{key}", default)

    def calc_sheet(self):
        name = str(self.get("计算底稿工作表", CALC_SHEET) or "").strip()
        return name or CALC_SHEET

    def period_sheets(self):
        """需要更新会计期间与批号的页签：计算底稿 + 租赁 + 收入类。"""
        names = [self.calc_sheet()]
        for name in PERIOD_SHEETS:
            if name not in names:
                names.append(name)
        return names

    def accounting_period(self):
        """会计期间首日：显式配置则用配置值，空白按自然月（当月）。"""
        first, _ = month_bounds(configured_period(self.get("会计期间")))
        return first

    def ledger_month(self):
        """明细账共用的会计期间：上下文优先，空白按自然月。"""
        value = self.context.get("jde_ledger_period")
        if not str(value or "").strip():
            value = self.get("会计期间")
        return month_bounds(configured_period(value))

    def ledger_dates(self):
        """明细账起始与至日期固定跟随会计期间的自然月首末日。"""
        first, last = self.ledger_month()
        return (format_jde_date(first, "会计期间起始日期"),
                format_jde_date(last, "会计期间结束日期"))

    def export_ledger(self):
        """从 JDE 导出明细账，返回下载文件路径。"""
        required = ("明细账自公司", "明细账至公司")
        missing = [name for name in required
                   if self.get(name) is None or not str(self.get(name)).strip()]
        if missing:
            raise ConfigError(f"缺少导出参数：{', '.join(missing)}")
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        start_date, end_date = self.ledger_dates()

        destination = self.get("明细账文件")
        if destination:
            destination = Path(destination)
            if not destination.is_absolute():
                destination = self.config.get_path("output") / destination
        else:
            destination = self.config.get_path("output") / f"JDE明细账 {get_timestamp()}"
        path = download_gl_movement(
            page,
            ledger_type=self.get("明细账分类帐类型", "AA"),
            from_company=self.get("明细账自公司"),
            to_company=self.get("明细账至公司"),
            start_date=start_date,
            end_date=end_date,
            destination=destination,
        )
        self.context.set("jde_ledger_path", path)
        return path

    def create_template(self, ledger_path):
        """复制模板生成输出文件，按标题把明细账写入模板的 JDE明细账 页。"""
        input_dir = self.config.get_path("input")
        template_name = self.get("模板文件")
        if template_name:
            template = Path(template_name)
            if not template.is_absolute():
                template = input_dir / template
        else:
            candidates = [p for p in input_dir.glob("*印花税计税表*.xlsx")
                          if not p.name.startswith("~$")
                          and p.resolve() != Path(ledger_path).resolve()]
            if len(candidates) != 1:
                raise ConfigError(
                    "输入目录需有唯一包含“印花税计税表”的 xlsx 模板，"
                    "或配置 印花税配置.模板文件"
                )
            template = candidates[0]
        if not template.is_file():
            raise ConfigError(f"印花税模板不存在：{template}")

        company = self.get("公司")
        if company is None or not str(company).strip() or re.search(r'[<>:"/\\|?*\x00-\x1f]', str(company)):
            raise ConfigError("印花税配置.公司 必填且不能包含文件名非法字符")
        company = str(company).strip()
        # 会计期间未配置时按自然月（当月），文件名统一用 yyyymm。
        period = f"{self.accounting_period():%Y%m}"
        output_dir = Path(self.get("输出目录") or self.config.get_path("output"))
        if not output_dir.is_absolute():
            output_dir = self.config.get_path("output") / output_dir
        destination = output_dir / f"印花税计税表 {company} {period}.xlsx"
        if destination.resolve() in (template.resolve(), Path(ledger_path).resolve()):
            raise ConfigError(f"输出文件不能与输入重合：{destination}")
        if destination.exists():
            try:
                destination.unlink()
            except OSError as exc:
                raise ConfigError(
                    f"无法覆盖同名更新表 {destination}："
                    "文件可能仍被 Excel 占用，请关闭后重试"
                ) from exc
            logger.info(f"已覆盖同名更新表：{destination}")

        mapping = self.get("标题映射", "")
        try:
            mapping = json.loads(mapping) if mapping else {}
        except (TypeError, ValueError) as exc:
            raise ConfigError("标题映射须为 JSON 对象：源标题 -> 模板标题") from exc
        if not isinstance(mapping, dict) or not all(
                isinstance(k, str) and isinstance(v, str) for k, v in mapping.items()):
            raise ConfigError("标题映射须为字符串到字符串的 JSON 对象")

        source = openpyxl.load_workbook(ledger_path, data_only=True)
        target = None
        try:
            sheet_name = self.get("明细账工作表") or source.sheetnames[0]
            if sheet_name not in source.sheetnames:
                raise ConfigError(f"明细账缺少工作表：{sheet_name}")
            source_sheet = source[sheet_name]
            target = openpyxl.load_workbook(template)
            if LEDGER_SHEET not in target.sheetnames:
                raise ConfigError(f"模板缺少工作表：{LEDGER_SHEET}")
            target_sheet = target[LEDGER_SHEET]
            source_row = int(self.get("明细账标题行", 1))
            target_row = int(self.get("模板标题行", 1))
            if min(source_row, target_row) < 1:
                raise ConfigError("标题行必须为正整数")
            src = header_columns(source_sheet, source_row)
            dst = header_columns(target_sheet, target_row)
            pairs = build_header_pairs(list(src), list(dst), mapping)
            if not pairs:
                raise ConfigError("没有匹配的标题，请补充 印花税配置.标题映射")
            if len(set(pairs.values())) != len(pairs):
                raise ConfigError("多个源标题不能写入同一个模板标题")
            unmatched = sorted(set(src) - set(pairs))
            if unmatched:
                raise ConfigError(f"明细账标题未映射到模板：{unmatched}")
            missing = sorted(set(dst) - set(pairs.values()))
            if missing:
                raise ConfigError(f"明细账缺少模板所需标题：{missing}")
            rows = [row for row in source_sheet.iter_rows(min_row=source_row + 1, values_only=True)
                    if not all(value is None or (isinstance(value, str) and not value.strip())
                               for value in row)]
            if not rows:
                raise ConfigError("明细账没有数据，未生成输出")
            column_pairs = [(dst[name], src[origin]) for origin, name in pairs.items()]
        finally:
            source.close()
            if target is not None:
                target.close()

        labels = [(name, PERIOD_LABEL, self.accounting_period(), excel_com.DATE_FORMAT)
                  for name in self.period_sheets()]
        engine = str(self.get("重算引擎", "Excel.Application")).strip()
        self.session = excel_com.write_ledger(
            template, destination, LEDGER_SHEET, target_row, column_pairs, rows,
            labels=labels, engine=engine,
        )
        self.context.set("stamp_tax_path", destination)
        self.context.set("jde_ledger_rows", len(rows))
        self.context.set("jde_ledger_mapping", pairs)
        return destination

    def journal_rows(self, path):
        """按标题动态定位计算底稿分录区，读取账号/借项/贷项/备注。"""
        book = openpyxl.load_workbook(path, data_only=True)
        try:
            name = self.calc_sheet()
            if name not in book.sheetnames:
                raise ConfigError(f"计算底稿缺少工作表：{name}")
            sheet = book[name]
            _, columns, region = journal_region(sheet)
            rows = []
            for row in region:
                code = sheet.cell(row, columns["账号"]).value
                if code is None or not str(code).strip():
                    continue
                debit = sheet.cell(row, columns["借项金额"]).value
                credit = sheet.cell(row, columns["贷项金额"]).value
                if debit is None and credit is None:
                    coordinate = sheet.cell(row, columns["借项金额"]).coordinate
                    raise ConfigError(
                        f"{name}!{coordinate} 金额为空，公式可能未重算；"
                        "请确认 Excel 可正常重算后重试"
                    )
                if _zero_amount(debit) and _zero_amount(credit):
                    logger.info(f"{name} 第 {row} 行借贷金额均为 0，跳过")
                    continue
                remark = sheet.cell(row, columns["备注"]).value
                rows.append((str(code).strip(), debit, credit, remark))
            if not rows:
                raise ConfigError("计算底稿没有可录入的分录行")
            description = str(rows[0][3] or "").strip()
            return rows, description
        finally:
            book.close()

    def write_batch_number(self, path, batch):
        """把 JDE 批号回写到计算底稿、租赁、收入类“批号：”右侧单元格。"""
        if not batch:
            return
        names = self.period_sheets()
        if self.session is not None:
            for name in names:
                self.session.write_label_value(name, BATCH_LABEL, batch)
            self.session.save()
            logger.info(f"JDE 批号 {batch} 已回写：{'、'.join(names)}")
            return
        excel_com.update_cells(
            path,
            [(name, BATCH_LABEL, batch) for name in names],
            engine=self.get("重算引擎", "Excel.Application"),
        )

    def write_journal(self, path):
        """把计算底稿分录区整块录入 JDE 日记帐分录，并回写批号。"""
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        rows, description = self.journal_rows(path)
        batch = create_journal_entry(
            page,
            company=self.get("分录公司", "02501"),
            description=description,
            gl_date=format_jde_date(date.today(), "总帐日期"),
            rows=rows,
        )
        self.context.set("jde_journal_batch", batch)
        self.context.set("jde_journal_rows", len(rows))
        self.write_batch_number(path, batch)
        return batch

    def run(self):
        self.context.set("stamp_tax_status", "running")
        try:
            ledger_path = self.export_ledger()
            path = self.create_template(ledger_path)
            try:
                self.write_journal(path)
            finally:
                self.close()
            self.context.set("stamp_tax_status", "completed")
        except Exception:
            self.context.set("stamp_tax_status", "failed")
            self.close()
            raise

    def close(self):
        """关闭生成计税表时保留的 Excel 会话。"""
        if self.session is not None:
            self.session.close()
            self.session = None
