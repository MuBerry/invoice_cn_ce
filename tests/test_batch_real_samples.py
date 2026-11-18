# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""真实票据的对账测试 —— 需要 ``INVOICE_FIXTURES`` 环境变量。

``INVOICE_FIXTURES`` 指向一个**本地**目录（真实票据不入库，公开仓库取不到）。
未设置时全部 ``skip``，公开 CI 只会看到 skip，不会失败也不会泄数据。

判据（不依赖工具自报的合计，逐票独立核对）：
1. 每张票 ``Σitems.amount == 票面不含税金额``，``Σitems.tax == 票面税额``；
2. 票面 ``金额 + 税额 == 价税合计``；
3. 明细行金额加起来与逐票相加的合计一致（勾稽）；
4. 折扣/冲销行（负金额）被保留，且同一张票里净额为正；
5. 折让摊回后，品目汇总的金额合计 == 发票清单金额合计。
"""

from __future__ import annotations

import json
import json
from pathlib import Path

import pytest

from invoice_cn_ce.batch.aggregate import (
    TOLERANCE,
    product_stats,
    to_float,
    totals,
)
from invoice_cn_ce.batch.aggregate import build_invoice
from invoice_cn_ce.batch.ingest import Ingestor
from invoice_cn_ce.formats import parse_pdf
from tests.conftest import real_pdf_dir, run_cli


@pytest.fixture(scope="module")
def parsed():
    """把样本目录里的每张票解析一次，供本模块所有用例复用。"""
    folder = real_pdf_dir()
    result = Ingestor().collect([str(folder)])
    invoices = []
    failures = []
    try:
        for source in result.files:
            resp = parse_pdf(Path(source.path).read_bytes(), filename=Path(source.path).name)
            if getattr(resp, "message", "success") != "success":
                failures.append((source.display, resp.message))
                continue
            invoices.append(build_invoice(resp, source))
    finally:
        result.cleanup()
    return invoices, failures


def test_every_sample_parsed(parsed) -> None:
    invoices, failures = parsed
    assert invoices, "样本目录里没有解析成功的发票"
    assert not failures, f"以下样本解析失败：{failures}"


def test_invoice_header_totals_reconcile(parsed) -> None:
    invoices, _failures = parsed
    bad = [
        (inv.file_label, inv.amount, inv.tax, inv.total)
        for inv in invoices
        if inv.amount is not None
        and inv.tax is not None
        and inv.total is not None
        and abs(inv.amount + inv.tax - inv.total) > TOLERANCE
    ]
    assert bad == []


def test_item_amounts_reconcile_with_header(parsed) -> None:
    invoices, _failures = parsed
    bad = []
    for inv in invoices:
        assert inv.items, f"{inv.file_label} 没解析到明细行"
        item_amount = round(sum(to_float(it.amount) or 0.0 for it in inv.items), 2)
        item_tax = round(sum(to_float(it.tax_amount) or 0.0 for it in inv.items), 2)
        if inv.amount is not None and abs(item_amount - inv.amount) > TOLERANCE:
            bad.append((inv.file_label, "amount", inv.amount, item_amount))
        if inv.tax is not None and abs(item_tax - inv.tax) > TOLERANCE:
            bad.append((inv.file_label, "tax", inv.tax, item_tax))
    assert bad == []


def test_item_fields_are_populated(parsed) -> None:
    invoices, _failures = parsed
    empty_name = [
        (inv.file_label, it.item_no) for inv in invoices for it in inv.items if not it.name.strip()
    ]
    assert empty_name == []
    no_rate = [
        (inv.file_label, it.item_no)
        for inv in invoices
        for it in inv.items
        if not it.tax_rate
    ]
    assert no_rate == []


def test_discount_rows_are_preserved(parsed) -> None:
    """负金额行不能丢：丢了会导致 Σitems 大于票面金额。"""
    invoices, _failures = parsed
    for inv in invoices:
        negatives = [it for it in inv.items if (to_float(it.amount) or 0) < 0]
        for item in negatives:
            assert item.is_reversal is True
            assert item.row_type == "折扣"
        if negatives:
            assert inv.amount is not None and inv.amount > 0


def test_totals_reconcile_across_invoices(parsed) -> None:
    invoices, _failures = parsed
    grand = totals(invoices)
    assert grand.reconciled is True, grand.notes
    assert grand.amount == pytest.approx(grand.sum_amount, abs=TOLERANCE)
    assert grand.tax == pytest.approx(grand.sum_tax, abs=TOLERANCE)
    assert grand.total == pytest.approx(grand.sum_total, abs=TOLERANCE)


def test_product_stats_total_equals_invoice_total(parsed) -> None:
    invoices, _failures = parsed
    stats = product_stats(invoices)
    grand = totals(invoices)
    assert sum(s.amount for s in stats) == pytest.approx(grand.amount, abs=TOLERANCE)
    assert sum(s.tax for s in stats) == pytest.approx(grand.tax, abs=TOLERANCE)


def test_cli_batch_on_fixture_dir_is_stable(tmp_path: Path, parsed) -> None:
    """CLI 端到端跑一遍：张数、合计、勾稽与直接解析结果一致。"""
    folder = real_pdf_dir()
    proc = run_cli([str(folder), "-o", str(tmp_path), "--quiet"])
    assert proc.returncode == 0, proc.stderr

    payload = json.loads((tmp_path / "数据.json").read_text(encoding="utf-8"))
    invoices, _failures = parsed
    assert payload["totals"]["count"] == len(invoices)
    assert payload["totals"]["reconciled"] is True
    assert payload["meta"]["issue_count"] == 0
