# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""报表输出：xlsx（五个 sheet）+ csv（发票清单主表）+ json（全量结构化）。

Excel 面向人看：冻结首行、列宽、金额 ``#,##0.00``、合计行、普票绿/专票红、
异常红底、折扣行红字、明细自动筛选。CSV 面向二次加工（``utf-8-sig``，Excel
双击不乱码）。JSON 面向程序对接（保留全部字段，含明细行原文）。
"""

from __future__ import annotations

import csv
import json
import os
from dataclasses import asdict
from typing import Any, Optional, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from invoice_cn_ce.batch.aggregate import (
    BatchInvoice,
    DedupInfo,
    ProductStat,
    Totals,
    round2,
    to_float,
)

# ---------------------------------------------------------------- 样式

FONT_NAME = "微软雅黑"
HDR_FILL = PatternFill("solid", fgColor="DCE9F7")
SUM_FILL = PatternFill("solid", fgColor="EDF2F7")
WARN_FILL = PatternFill("solid", fgColor="FDECEA")
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="1F3552")
SUB_FONT = Font(name=FONT_NAME, size=9, color="5A6B7B")
HDR_FONT = Font(name=FONT_NAME, size=10, bold=True, color="1F3552")
BODY_FONT = Font(name=FONT_NAME, size=10, color="22303C")
SUM_FONT = Font(name=FONT_NAME, size=10, bold=True, color="1F3552")
GREEN_FONT = Font(name=FONT_NAME, size=10, bold=True, color="3B6D11")   # 普票
RED_FONT = Font(name=FONT_NAME, size=10, bold=True, color="B23A2E")     # 专票 / 异常
DISC_FONT = Font(name=FONT_NAME, size=10, color="B23A2E")               # 折扣行
MUTED_FONT = Font(name=FONT_NAME, size=9, color="7A8899")
THIN = Side(style="thin", color="C9D3DD")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0.00"
PCT_FMT = "0.0%"

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

ROW_TYPE_DISCOUNT = "折扣"


# ---------------------------------------------------------------- 通用工具


def _header_row(ws, row: int, headers: Sequence[str], height: int = 28) -> None:
    for idx, text in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=idx, value=text)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def _set_widths(ws, widths: Sequence[float]) -> None:
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def _cell(ws, row: int, col: int, value: Any, font=None, fmt: Optional[str] = None,
          align=None, fill=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font or BODY_FONT
    cell.border = BOX
    if fmt:
        cell.number_format = fmt
    if align is not None:
        cell.alignment = align
    if fill is not None:
        cell.fill = fill
    return cell


def _kind_font(inv: BatchInvoice) -> Font:
    return RED_FONT if inv.deductible else GREEN_FONT


def check_text(inv: BatchInvoice) -> str:
    """校验结果列：通过 / 具体原因。"""
    return "通过" if not inv.has_issue else "；".join(inv.issues)


def _fmt_money(value: Optional[float]) -> str:
    return "" if value is None else f"{value:.2f}"


def strip_category_prefix(name: str) -> str:
    """``*类别*品名`` -> ``品名``；没有前缀时原样返回。"""
    if name.startswith("*"):
        end = name.find("*", 1)
        if 1 < end < 24:
            return name[end + 1:]
    return name


def display_name(item) -> str:
    """明细表里展示的品名：去掉 ``*类别*`` 前缀（类别已单独占一列）。"""
    return strip_category_prefix(item.name or "")


# ---------------------------------------------------------------- Sheet 1


def _sheet_invoices(ws, invoices: Sequence[BatchInvoice], grand: Totals) -> None:
    ws.title = "发票清单"
    headers = ["序号", "票种", "可抵扣", "开票日期", "发票号码", "购买方", "销售方",
               "金额(不含税)", "税额", "价税合计", "税率", "明细行数", "校验结果", "来源文件"]
    last_col = len(headers)

    ws.cell(row=1, column=1, value=f"发票归集汇总（共 {len(invoices)} 张）").font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    subtitle = (
        f"不含税金额 {grand.amount:,.2f} 元　税额 {grand.tax:,.2f} 元　"
        f"价税合计 {grand.total:,.2f} 元　"
        f"勾稽：{'通过' if grand.reconciled else '不通过'}　"
        f"异常 {sum(1 for i in invoices if i.has_issue)} 张"
    )
    ws.cell(row=2, column=1, value=subtitle).font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)

    _header_row(ws, 3, headers)
    _set_widths(ws, [5, 24, 8, 12, 22, 30, 30, 14, 12, 13, 8, 9, 52, 46])

    row = 4
    for idx, inv in enumerate(invoices, start=1):
        values = [
            idx, inv.kind or "未识别", "是" if inv.deductible else "否", inv.invoice_date,
            inv.invoice_no, inv.buyer_name, inv.seller_name, inv.amount, inv.tax, inv.total,
            inv.rate, len(inv.items), check_text(inv), inv.file_label,
        ]
        for col, value in enumerate(values, start=1):
            fmt = MONEY_FMT if col in (8, 9, 10) else None
            align = RIGHT if col in (8, 9, 10) else (
                LEFT_WRAP if col in (6, 7, 13, 14) else CENTER
            )
            cell = _cell(ws, row, col, value, fmt=fmt, align=align)
            if col == 2:
                cell.font = _kind_font(inv)
            if col == 10 and inv.has_issue:
                cell.fill = WARN_FILL
            if col == 13 and inv.has_issue:
                cell.font = Font(name=FONT_NAME, size=9, color="B23A2E")
            if col == 14:
                cell.font = MUTED_FONT
        row += 1

    total_values = {1: "合计", 2: f"{len(invoices)} 张", 8: grand.amount, 9: grand.tax,
                    10: grand.total, 12: grand.item_count}
    for col in range(1, last_col + 1):
        _cell(ws, row, col, total_values.get(col, ""), font=SUM_FONT, fill=SUM_FILL,
              fmt=MONEY_FMT if col in (8, 9, 10) else None,
              align=RIGHT if col in (8, 9, 10) else CENTER)
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(last_col)}{row - 1}"
    return row


def _sheet_invoice_footer(ws, row: int, last_col: int, lines: Sequence[str]) -> None:
    """在「发票清单」底部写运行信息 / 去重信息（不额外占一个 sheet）。"""
    row += 1
    for text in lines:
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = Font(name=FONT_NAME, size=9, color="5A6B7B")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
        row += 1


# ---------------------------------------------------------------- Sheet 2


def _sheet_items(ws, invoices: Sequence[BatchInvoice]) -> None:
    headers = ["序号", "发票号码", "开票日期", "票种", "销售方", "行类型", "商品类别",
               "商品名称", "规格型号", "单位", "数量", "单价", "金额(不含税)", "税率",
               "税额", "原始行"]
    last_col = len(headers)
    _header_row(ws, 1, headers)
    _set_widths(ws, [5, 22, 12, 24, 28, 8, 16, 46, 26, 6, 8, 11, 13, 8, 11, 62])

    row = 2
    seq = 0
    for inv in invoices:
        for item in inv.items:
            seq += 1
            row_type = item.row_type or ("折扣" if item.is_reversal else "商品")
            font = DISC_FONT if row_type == ROW_TYPE_DISCOUNT else BODY_FONT
            values = [
                seq, inv.invoice_no, inv.invoice_date, inv.kind, inv.seller_name, row_type,
                item.category, display_name(item), item.spec, item.unit, item.quantity,
                item.unit_price,
                to_float(item.amount), item.tax_rate, to_float(item.tax_amount), item.raw_line,
            ]
            for col, value in enumerate(values, start=1):
                fmt = MONEY_FMT if col in (12, 13, 15) else None
                align = RIGHT if col in (11, 12, 13, 15) else (
                    LEFT_WRAP if col in (7, 8, 9, 16) else CENTER
                )
                _cell(ws, row, col, value, font=font, fmt=fmt, align=align)
            row += 1
    if row > 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{row - 1}"
    else:
        _cell(ws, row, 1, "未解析到任何明细行", font=SUM_FONT, align=CENTER)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------- Sheet 3


def _sheet_groups(ws, groups: dict, grand: Totals) -> None:
    _set_widths(ws, [40, 10, 16, 14, 16, 11])
    row = 1
    for dimension, stats in groups.items():
        ws.cell(row=row, column=1, value=f"按「{dimension}」汇总").font = TITLE_FONT
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        row += 1
        _header_row(ws, row, [dimension, "张数", "金额(不含税)", "税额", "价税合计", "金额占比"],
                    height=24)
        row += 1
        for stat in stats:
            _cell(ws, row, 1, stat.key, align=LEFT_WRAP)
            _cell(ws, row, 2, stat.count, align=CENTER)
            _cell(ws, row, 3, round2(stat.amount), fmt=MONEY_FMT, align=RIGHT)
            _cell(ws, row, 4, round2(stat.tax), fmt=MONEY_FMT, align=RIGHT)
            _cell(ws, row, 5, round2(stat.total), fmt=MONEY_FMT, align=RIGHT)
            _cell(ws, row, 6, stat.share, fmt=PCT_FMT, align=RIGHT)
            row += 1
        _cell(ws, row, 1, "合计", font=SUM_FONT, fill=SUM_FILL, align=CENTER)
        _cell(ws, row, 2, sum(s.count for s in stats), font=SUM_FONT, fill=SUM_FILL, align=CENTER)
        _cell(ws, row, 3, round2(sum(s.amount for s in stats)) or 0.0, font=SUM_FONT,
              fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 4, round2(sum(s.tax for s in stats)) or 0.0, font=SUM_FONT,
              fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 5, round2(sum(s.total for s in stats)) or 0.0, font=SUM_FONT,
              fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 6, 1.0 if stats else 0.0, font=SUM_FONT, fill=SUM_FILL,
              fmt=PCT_FMT, align=RIGHT)
        row += 2

    verdict = "一致" if grand.reconciled else "不一致：" + "；".join(grand.notes)
    ws.cell(row=row, column=1,
            value=(f"勾稽总校验：{grand.count} 张票金额合计 {grand.amount:,.2f} 元、"
                   f"税额 {grand.tax:,.2f} 元、价税合计 {grand.total:,.2f} 元；"
                   f"与各票逐张相加（{grand.sum_amount:,.2f} / {grand.sum_tax:,.2f} / "
                   f"{grand.sum_total:,.2f}）{verdict}")).font = (
        Font(name=FONT_NAME, size=9, color="1F3552" if grand.reconciled else "B23A2E"))
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)


# ---------------------------------------------------------------- Sheet 4


def _sheet_products(ws, products: Sequence[ProductStat]) -> None:
    headers = ["品目（归一化）", "类别", "涉及票数", "明细行数", "金额(不含税)", "税额",
               "价税合计", "金额占比"]
    _header_row(ws, 1, headers)
    _set_widths(ws, [46, 18, 10, 10, 16, 14, 16, 11])
    row = 2
    for stat in products:
        _cell(ws, row, 1, stat.key, align=LEFT_WRAP)
        _cell(ws, row, 2, stat.category, align=LEFT_WRAP)
        _cell(ws, row, 3, stat.invoice_count, align=CENTER)
        _cell(ws, row, 4, stat.line_count, align=CENTER)
        _cell(ws, row, 5, round2(stat.amount), fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 6, round2(stat.tax), fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 7, round2(stat.total), fmt=MONEY_FMT, align=RIGHT)
        _cell(ws, row, 8, stat.share, fmt=PCT_FMT, align=RIGHT)
        row += 1
    _cell(ws, row, 1, "合计", font=SUM_FONT, fill=SUM_FILL, align=CENTER)
    _cell(ws, row, 2, "", font=SUM_FONT, fill=SUM_FILL)
    _cell(ws, row, 3, "", font=SUM_FONT, fill=SUM_FILL)
    _cell(ws, row, 4, sum(s.line_count for s in products), font=SUM_FONT, fill=SUM_FILL,
          align=CENTER)
    _cell(ws, row, 5, round2(sum(s.amount for s in products)) or 0.0, font=SUM_FONT,
          fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
    _cell(ws, row, 6, round2(sum(s.tax for s in products)) or 0.0, font=SUM_FONT,
          fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
    _cell(ws, row, 7, round2(sum(s.total for s in products)) or 0.0, font=SUM_FONT,
          fill=SUM_FILL, fmt=MONEY_FMT, align=RIGHT)
    _cell(ws, row, 8, 1.0 if products else 0.0, font=SUM_FONT, fill=SUM_FILL,
          fmt=PCT_FMT, align=RIGHT)
    row += 2
    ws.cell(row=row, column=1,
            value=("口径：折扣行按位置摊回其上一条商品行；品名先去掉 *类别* 前缀与括号说明，"
                   "再按规格噪声（支/装/瓶装/适用于…）截断，最后把「某品名是另一品名前缀」"
                   "的同物异名合并。品目合计应与发票清单金额合计一致。")).font = SUB_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------- Sheet 5


def _sheet_issues(ws, issues: Sequence[dict]) -> None:
    headers = ["序号", "来源文件", "等级", "发票号码", "原因", "原始行"]
    last_col = len(headers)
    _header_row(ws, 1, headers)
    _set_widths(ws, [5, 46, 8, 22, 66, 60])
    row = 2
    for idx, issue in enumerate(issues, start=1):
        _cell(ws, row, 1, idx, align=CENTER)
        _cell(ws, row, 2, issue.get("file", ""), font=MUTED_FONT, align=LEFT_WRAP)
        _cell(ws, row, 3, issue.get("level", ""), align=CENTER)
        _cell(ws, row, 4, issue.get("invoice_number", ""), align=CENTER)
        _cell(ws, row, 5, issue.get("reason", ""),
              font=Font(name=FONT_NAME, size=10, color="B23A2E"), align=LEFT_WRAP)
        _cell(ws, row, 6, issue.get("raw", ""), font=MUTED_FONT, align=LEFT_WRAP)
        row += 1
    if not issues:
        _cell(ws, row, 1, "无异常", font=SUM_FONT, align=CENTER)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    else:
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{row - 1}"
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------- 落盘


def write_outputs(out_dir: str, invoices: Sequence[BatchInvoice], issues: Sequence[dict],
                  dedup: DedupInfo, grand: Totals, groups: dict,
                  products: Sequence[ProductStat], meta: dict, stamp: str) -> dict:
    """写 xlsx / csv / json，返回三个文件路径。"""
    os.makedirs(out_dir, exist_ok=True)
    base = f"发票汇总-{stamp}"
    xlsx_path = os.path.join(out_dir, f"{base}.xlsx")
    csv_path = os.path.join(out_dir, f"{base}.csv")
    json_path = os.path.join(out_dir, "数据.json")

    wb = Workbook()
    ws = wb.active
    last_row = _sheet_invoices(ws, invoices, grand)
    _sheet_invoice_footer(
        ws, last_row, 14,
        [
            f"运行时间：{meta.get('run_at', '')}　"
            f"耗时 {meta.get('elapsed', 0):.2f} 秒　"
            f"输入 {len(meta.get('inputs', []))} 个路径 / {meta.get('file_count', 0)} 个文件　"
            f"工具：invoice_cn_ce {meta.get('version', '')} batch",
            f"去重：重复 {dedup.duplicate_count} 处已忽略"
            f"（内容重复 {len(dedup.by_hash)}，同号 {len(dedup.by_number)}）"
            + (f"；同号重复："
               + "、".join(f"{a} ↔ {b}" for a, b in dedup.by_number[:5])
               if dedup.by_number else ""),
            f"分组维度：{'、'.join(meta.get('group_dims', []))}　"
            f"校验容差：±{meta.get('tolerance', 0.02)} 元　"
            f"异常条目：{len(issues)}",
        ],
    )
    _sheet_items(wb.create_sheet("发票明细"), invoices)
    _sheet_groups(wb.create_sheet("分组汇总"), groups, grand)
    _sheet_products(wb.create_sheet("品目汇总"), products)
    _sheet_issues(wb.create_sheet("异常清单"), issues)
    wb.save(xlsx_path)

    _write_csv(csv_path, invoices, grand)
    _write_json(json_path, invoices, issues, dedup, grand, groups, products, meta)
    return {"xlsx": xlsx_path, "csv": csv_path, "json": json_path}


def _write_csv(path: str, invoices: Sequence[BatchInvoice], grand: Totals) -> None:
    header = ["序号", "票种", "可抵扣", "开票日期", "发票号码", "购买方", "销售方",
              "金额(不含税)", "税额", "价税合计", "税率", "明细行数", "校验结果", "来源文件"]
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fp:
        writer = csv.writer(fp)
        writer.writerow(header)
        for idx, inv in enumerate(invoices, start=1):
            writer.writerow([
                idx, inv.kind, "是" if inv.deductible else "否", inv.invoice_date,
                inv.invoice_no, inv.buyer_name, inv.seller_name,
                _fmt_money(inv.amount), _fmt_money(inv.tax), _fmt_money(inv.total),
                inv.rate, len(inv.items), check_text(inv), inv.file_label,
            ])
        writer.writerow(["合计", f"{len(invoices)} 张", "", "", "", "", "",
                         _fmt_money(grand.amount), _fmt_money(grand.tax),
                         _fmt_money(grand.total), "", grand.item_count, "", ""])
    os.replace(tmp, path)


def _write_json(path: str, invoices: Sequence[BatchInvoice], issues: Sequence[dict],
                dedup: DedupInfo, grand: Totals, groups: dict,
                products: Sequence[ProductStat], meta: dict) -> None:
    payload = {
        "meta": meta,
        "totals": {
            "count": grand.count,
            "amount_without_tax": grand.amount,
            "tax_amount": grand.tax,
            "total_amount": grand.total,
            "item_count": grand.item_count,
            "discount_amount": grand.discount,
            "reconciled": grand.reconciled,
            "reconcile_notes": grand.notes,
        },
        "dedup": {
            "duplicate_count": dedup.duplicate_count,
            "by_hash": [{"duplicate": a, "kept": b} for a, b in dedup.by_hash],
            "by_number": [{"duplicate": a, "kept": b} for a, b in dedup.by_number],
        },
        "invoices": [inv.to_dict() for inv in invoices],
        "issues": list(issues),
        "groups": {dim: [asdict(stat) for stat in stats] for dim, stats in groups.items()},
        "products": [asdict(stat) for stat in products],
    }
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=1, default=str)
    os.replace(tmp, path)


__all__ = ["check_text", "display_name", "strip_category_prefix", "write_outputs"]
