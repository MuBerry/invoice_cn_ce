# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""报表产物：五个 sheet、数值型金额、冻结窗格、异常清单、CSV 编码、JSON 结构。

直接用内存构造的 :class:`BatchInvoice` 调 ``write_outputs``，因此本文件不依赖
任何真实票据，公开 CI 可跑。
"""

from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook

from invoice_cn_ce.batch.aggregate import DedupInfo, group_stats, product_stats, totals
from invoice_cn_ce.batch.report import write_outputs
from tests.conftest import make_invoice, make_item

EXPECTED_SHEETS = ["发票清单", "发票明细", "分组汇总", "品目汇总", "异常清单"]


@pytest.fixture()
def dataset():
    """两张票：一张带折扣两行，一张单行；外加一条异常与一条去重记录。"""
    items_a = [
        make_item("393.63", "51.17", name="*测试类别*甲商品"),
        make_item("-180.55", "-23.47", name="*测试类别*甲商品", item_no=2),
    ]
    items_b = [make_item("100.00", "13.00", name="*测试类别*乙商品")]
    invoices = [
        make_invoice("N0001", 213.08, 27.70, items=items_a, display="a.pdf", date="2026-01-02"),
        make_invoice("N0002", 100.00, 13.00, items=items_b, display="b.pdf", date="2026-02-03"),
    ]
    issues = [
        {"file": "坏.pdf", "level": "文件级", "invoice_number": "",
         "reason": "PDF 无文本层（扫描件 / 纯图片）", "raw": ""},
        {"file": "a.pdf", "level": "票级", "invoice_number": "N0001",
         "reason": "测试用票级提示", "raw": "原始行"},
    ]
    dedup = DedupInfo(by_hash=[("dup.pdf", "a.pdf")], by_number=[("c.pdf", "a.pdf")])
    grand = totals(invoices)
    groups = {dim: group_stats(invoices, dim) for dim in ("销售方", "月份")}
    products = product_stats(invoices)
    meta = {
        "run_at": "2026-02-03 10:00:00",
        "elapsed": 0.01,
        "inputs": ["invoices"],
        "file_count": 2,
        "group_dims": ["销售方", "月份"],
        "tolerance": 0.02,
        "version": "1.0.0",
    }
    return invoices, issues, dedup, grand, groups, products, meta


@pytest.fixture()
def outputs(tmp_path: Path, dataset):
    invoices, issues, dedup, grand, groups, products, meta = dataset
    return write_outputs(str(tmp_path / "子目录" / "报表"), invoices, issues, dedup, grand,
                         groups, products, meta, "20260203-100000")


def test_writes_three_files_creating_out_dir(outputs) -> None:
    for key in ("xlsx", "csv", "json"):
        assert Path(outputs[key]).exists(), key
    assert Path(outputs["xlsx"]).name == "发票汇总-20260203-100000.xlsx"
    assert Path(outputs["json"]).name == "数据.json"


def test_workbook_has_exactly_five_sheets(outputs) -> None:
    wb = load_workbook(outputs["xlsx"])
    assert wb.sheetnames == EXPECTED_SHEETS


def test_invoice_sheet_amounts_are_numeric_and_sumable(outputs) -> None:
    ws = load_workbook(outputs["xlsx"])["发票清单"]
    headers = [c.value for c in ws[3]]
    assert headers[:10] == ["序号", "票种", "可抵扣", "开票日期", "发票号码", "购买方", "销售方",
                            "金额(不含税)", "税额", "价税合计"]
    amount_col = headers.index("金额(不含税)") + 1
    tax_col = headers.index("税额") + 1
    total_col = headers.index("价税合计") + 1
    values = [ws.cell(row=4, column=c).value for c in (amount_col, tax_col, total_col)]
    assert all(isinstance(v, float) for v in values), values
    assert ws.cell(row=4, column=amount_col).number_format == "#,##0.00"

    # 合计行存在且等于逐票相加
    last = ws.max_row
    while ws.cell(row=last, column=1).value != "合计":
        last -= 1
    assert ws.cell(row=last, column=amount_col).value == pytest.approx(313.08)
    assert ws.cell(row=last, column=tax_col).value == pytest.approx(40.70)
    assert ws.cell(row=last, column=total_col).value == pytest.approx(353.78)
    assert isinstance(ws.cell(row=last, column=amount_col).value, float)


def test_freeze_panes_set_on_main_sheets(outputs) -> None:
    wb = load_workbook(outputs["xlsx"])
    assert wb["发票清单"].freeze_panes == "A4"
    assert wb["发票明细"].freeze_panes == "A2"
    assert wb["品目汇总"].freeze_panes == "A2"
    assert wb["异常清单"].freeze_panes == "A2"


def test_items_sheet_has_one_row_per_item(outputs, dataset) -> None:
    invoices = dataset[0]
    ws = load_workbook(outputs["xlsx"])["发票明细"]
    assert ws.max_row == 1 + sum(len(i.items) for i in invoices)
    amounts = [ws.cell(row=r, column=13).value for r in range(2, ws.max_row + 1)]
    # 必须是数值（int/float 都算：100.00 落盘后被读回 int 100），不能是文本
    assert all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in amounts), amounts
    # 折扣行为负值，且用红色标注
    assert any(v < 0 for v in amounts)
    discount_rows = [r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=6).value == "折扣"]
    assert discount_rows
    assert ws.cell(row=discount_rows[0], column=6).font.color.rgb.endswith("B23A2E")


def test_products_sheet_allocates_discount_back(outputs) -> None:
    ws = load_workbook(outputs["xlsx"])["品目汇总"]
    rows = {
        ws.cell(row=r, column=1).value: ws.cell(row=r, column=5).value
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=1).value not in (None, "合计")
    }
    assert rows["甲商品"] == pytest.approx(213.08)  # 393.63 - 180.55
    assert rows["乙商品"] == pytest.approx(100.00)


def test_issue_sheet_lists_every_issue(outputs) -> None:
    ws = load_workbook(outputs["xlsx"])["异常清单"]
    files = [ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)]
    assert "坏.pdf" in files and "a.pdf" in files
    reasons = [ws.cell(row=r, column=5).value for r in range(2, ws.max_row + 1)]
    assert any("无文本层" in (r or "") for r in reasons)


def test_issue_sheet_has_placeholder_when_clean(outputs, tmp_path, dataset) -> None:
    invoices, _issues, dedup, grand, groups, products, meta = dataset
    clean = write_outputs(str(tmp_path / "clean"), invoices, [], dedup, grand, groups, products,
                          meta, "20260203-100001")
    ws = load_workbook(clean["xlsx"])["异常清单"]
    assert ws.cell(row=2, column=1).value == "无异常"


def test_csv_is_utf8_sig_with_total_row(outputs, dataset) -> None:
    invoices = dataset[0]
    raw = Path(outputs["csv"]).read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf"), "Excel 双击需要 BOM"
    text = raw.decode("utf-8-sig")
    rows = list(csv.reader(text.splitlines()))
    assert rows[0][0] == "序号"
    assert len(rows) == len(invoices) + 2  # 表头 + 每票一行 + 合计
    assert rows[-1][0] == "合计"
    assert rows[-1][7] == "313.08"
    # 金额固定两位小数，便于二次加工
    assert rows[1][7] == "213.08"
    assert rows[1][8] == "27.70"


def test_json_structure_is_complete_and_reproducible(outputs) -> None:
    payload = json.loads(Path(outputs["json"]).read_text(encoding="utf-8"))
    assert set(payload) == {"meta", "totals", "dedup", "invoices", "issues", "groups", "products"}
    assert payload["totals"]["count"] == 2
    assert payload["totals"]["amount_without_tax"] == pytest.approx(313.08)
    assert payload["totals"]["reconciled"] is True
    assert payload["dedup"]["duplicate_count"] == 2
    assert len(payload["invoices"]) == 2
    first = payload["invoices"][0]
    assert first["invoice_number"] == "N0001"
    assert first["item_count"] == 2
    assert len(first["items"]) == 2
    assert {"name", "amount", "tax_amount", "row_type", "raw_line"} <= set(first["items"][0])
    assert {g["key"] for g in payload["groups"]["月份"]} == {"2026-01", "2026-02"}
    assert payload["products"][0]["amount"] > 0


def test_json_references_no_real_party_names(outputs) -> None:
    """报表只回显输入里的字段，测试素材是虚构名，这里顺带守住「测试不掺真实数据」。"""
    text = Path(outputs["json"]).read_text(encoding="utf-8")
    assert "测试销售方" in text


def test_group_sheet_contains_total_row_per_dimension(outputs) -> None:
    ws = load_workbook(outputs["xlsx"])["分组汇总"]
    titles = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    assert any(isinstance(t, str) and t.startswith("按「销售方」汇总") for t in titles)
    assert any(isinstance(t, str) and t.startswith("按「月份」汇总") for t in titles)
    totals_rows = [r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "合计"]
    assert len(totals_rows) == 2
    for row in totals_rows:
        assert isinstance(ws.cell(row=row, column=3).value, float)
        assert ws.cell(row=row, column=6).value == pytest.approx(1.0)


def test_empty_dataset_still_produces_valid_workbook(tmp_path: Path) -> None:
    out = write_outputs(str(tmp_path / "空"), [], [], DedupInfo(), totals([]), {}, [], {}, "S")
    wb = load_workbook(out["xlsx"])
    assert wb.sheetnames == EXPECTED_SHEETS
    ws = wb["发票明细"]
    assert ws.cell(row=2, column=1).value == "未解析到任何明细行"
    payload = json.loads(Path(out["json"]).read_text(encoding="utf-8"))
    assert payload["invoices"] == []
    assert payload["totals"]["count"] == 0
