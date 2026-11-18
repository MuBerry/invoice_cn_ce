# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""汇总层：金额解析、票级勾稽、按票号去重、合计、分组、品目折扣摊回。"""

from __future__ import annotations

import pytest

from invoice_cn_ce.batch.aggregate import (
    GROUP_CHOICES,
    TOLERANCE,
    DedupInfo,
    dedupe,
    dominant_rate,
    group_stats,
    money,
    normalize_product_name,
    parse_group_option,
    product_stats,
    round2,
    sort_invoices,
    to_float,
    totals,
    validate,
)
from tests.conftest import make_invoice, make_item


# ---------------------------------------------------------------- 数值工具


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("213.08", 213.08),
        ("¥213.08", 213.08),
        ("￥1,234.56", 1234.56),
        (" 27.70 ", 27.70),
        ("-180.55", -180.55),
        ("2 13.08", 213.08),
        (213.08, 213.08),
        (213, 213.0),
    ],
)
def test_to_float_parses_invoice_numbers(raw, expected) -> None:
    assert to_float(raw) == pytest.approx(expected)


@pytest.mark.parametrize("raw", [None, "", "  ", "一百", "¥", "abc.def"])
def test_to_float_returns_none_instead_of_zero(raw) -> None:
    """解析不出来必须是 None —— 静默当 0 会把错算藏起来。"""
    assert to_float(raw) is None


def test_round2_tames_float_dust() -> None:
    assert round2(0.1 + 0.2) == 0.3
    assert round2(None) is None
    assert round2(-2030.2449999999999) == -2030.24  # 负数的加 1e-9 修正


def test_money_treats_none_as_zero_for_totals_only() -> None:
    assert money(None) == 0.0
    assert money(12.34) == 12.34


# ---------------------------------------------------------------- 分组参数


def test_parse_group_option_defaults_to_seller() -> None:
    assert parse_group_option("") == ["销售方"]
    assert parse_group_option("   ") == ["销售方"]


def test_parse_group_option_supports_multiple_dims_and_cn_comma() -> None:
    assert parse_group_option("购买方,月份") == ["购买方", "月份"]
    assert parse_group_option("购买方，月份 税率") == ["购买方", "月份", "税率"]


def test_parse_group_option_dedupes_and_keeps_order() -> None:
    assert parse_group_option("月份,销售方,月份") == ["月份", "销售方"]


@pytest.mark.parametrize("bad", ["颜色", "销售方,颜色", "buyer"])
def test_parse_group_option_rejects_unknown_dimension(bad: str) -> None:
    with pytest.raises(ValueError) as excinfo:
        parse_group_option(bad)
    assert "不支持的分组维度" in str(excinfo.value)


def test_group_choices_are_stable() -> None:
    assert GROUP_CHOICES == ("销售方", "购买方", "月份", "税率", "品目")


# ---------------------------------------------------------------- 票级勾稽


def test_validate_passes_when_items_reconcile() -> None:
    items = [make_item("393.63", "51.17"), make_item("-180.55", "-23.47", item_no=2)]
    inv = make_invoice("X1", 213.08, 27.70, items=items)
    validate(inv)
    assert inv.has_issue is False
    assert inv.issues == []


def test_validate_flags_item_sum_mismatch() -> None:
    items = [make_item("100.00", "13.00")]
    inv = make_invoice("X2", 200.00, 26.00, items=items)
    validate(inv)
    assert inv.has_issue is True
    assert any("明细金额合计" in issue for issue in inv.issues)


def test_validate_flags_tax_mismatch_separately() -> None:
    items = [make_item("100.00", "5.00")]
    inv = make_invoice("X3", 100.00, 26.00, items=items)
    validate(inv)
    assert any("明细税额合计" in issue for issue in inv.issues)


def test_validate_tolerates_two_cent_rounding() -> None:
    items = [make_item("100.00", "13.00")]
    inv = make_invoice("X4", 100.01, 13.01, items=items)
    validate(inv)
    assert inv.has_issue is False  # 差额 0.01 <= TOLERANCE


def test_validate_flags_total_mismatch() -> None:
    inv = make_invoice("X5", 100.00, 13.00, items=[make_item("100.00", "13.00")])
    inv.total = 999.00  # 人为制造「金额+税额 ≠ 价税合计」
    validate(inv)
    assert any("价税合计" in issue for issue in inv.issues)


def test_validate_flags_missing_items() -> None:
    inv = make_invoice("X6", 100.00, 13.00)
    validate(inv)
    assert any("未解析到明细行" in issue for issue in inv.issues)


def test_validate_flags_missing_amounts() -> None:
    items = [make_item("", "")]
    inv = make_invoice("X7", 100.00, 13.00, items=items)
    validate(inv)
    assert any("未取到金额" in issue for issue in inv.issues)


def test_validate_flags_parser_error_message() -> None:
    inv = make_invoice("X8", 100.00, 13.00, items=[make_item("100.00", "13.00")])
    inv.message = "error: parse_failed: broken"
    validate(inv)
    assert any("解析器返回" in issue for issue in inv.issues)


def test_validate_dedupes_repeated_issue_text() -> None:
    inv = make_invoice("", 100.00, 13.00)
    validate(inv)
    assert inv.issues.count("未识别到发票号码") == 1


def test_deductible_only_for_special_invoices() -> None:
    general = make_invoice("A", 1, 0.13, items=[make_item("1.00")])
    special = make_invoice("B", 1, 0.13, items=[make_item("1.00")], kind="电子发票（增值税专用发票）")
    assert general.deductible is False
    assert special.deductible is True


# ---------------------------------------------------------------- 按票号去重


def test_dedupe_by_invoice_number_keeps_first() -> None:
    first = make_invoice("N1", 100, 13, items=[make_item("100.00", "13.00")], display="a.pdf")
    second = make_invoice("N2", 200, 26, items=[make_item("200.00", "26.00")], display="b.pdf")
    dup = make_invoice("N1", 100, 13, items=[make_item("100.00", "13.00")], display="c.pdf")
    kept, info = dedupe([first, second, dup])
    assert [i.invoice_no for i in kept] == ["N1", "N2"]
    assert info.by_number == [("c.pdf", "a.pdf")]
    assert info.total == 3 and info.kept == 2


def test_dedupe_disabled_keeps_every_invoice_number() -> None:
    invoices = [make_invoice("N1", 100, 13), make_invoice("N1", 100, 13)]
    kept, info = dedupe(invoices, enabled=False)
    assert len(kept) == 2
    assert info.by_number == []
    assert info.duplicate_count == 0


def test_dedupe_records_hash_duplicates_from_ingest() -> None:
    invoices = [make_invoice("N1", 100, 13)]
    kept, info = dedupe(invoices, [("dup.pdf", "kept.pdf")])
    assert len(kept) == 1
    assert info.by_hash == [("dup.pdf", "kept.pdf")]
    assert info.duplicate_count == 1


def test_dedupe_does_not_collapse_invoices_without_number() -> None:
    invoices = [make_invoice("", 100, 13), make_invoice("", 100, 13)]
    kept, _info = dedupe(invoices)
    assert len(kept) == 2


def test_dedup_info_defaults() -> None:
    info = DedupInfo()
    assert info.duplicate_count == 0


# ---------------------------------------------------------------- 合计


def test_totals_sum_and_reconcile() -> None:
    invoices = [
        make_invoice("N1", 213.08, 27.70, items=[make_item("393.63", "51.17")]),
        make_invoice("N2", 200.27, 26.03, items=[make_item("200.27", "26.03")]),
    ]
    grand = totals(invoices)
    assert grand.count == 2
    assert grand.amount == pytest.approx(413.35)
    assert grand.tax == pytest.approx(53.73)
    assert grand.total == pytest.approx(467.08)
    assert grand.item_count == 2
    assert grand.reconciled is True


def test_totals_ignore_none_amounts() -> None:
    a = make_invoice("N1", 100.00, 13.00)
    b = make_invoice("N2", 0, 0)
    b.amount = None
    b.total = None
    grand = totals([a, b])
    assert grand.amount == pytest.approx(100.00)
    assert grand.total == pytest.approx(113.00)


def test_totals_discount_aggregates_negative_rows() -> None:
    items = [make_item("393.63", "51.17"), make_item("-180.55", "-23.47", item_no=2)]
    grand = totals([make_invoice("N1", 213.08, 27.70, items=items)])
    assert grand.discount == pytest.approx(-180.55)


def test_totals_empty_input() -> None:
    grand = totals([])
    assert (grand.count, grand.amount, grand.total) == (0, 0.0, 0.0)
    assert grand.reconciled is True


def test_totals_detects_inconsistency() -> None:
    inv = make_invoice("N1", 100.00, 13.00)
    inv.total = 500.00  # 票内不勾稽
    grand = totals([inv])
    assert grand.reconciled is False
    assert grand.notes


# ---------------------------------------------------------------- 分组


def test_group_stats_by_seller_and_month() -> None:
    invoices = [
        make_invoice("N1", 100.00, 13.00, seller="甲", date="2026-01-05"),
        make_invoice("N2", 50.00, 6.50, seller="甲", date="2026-02-05"),
        make_invoice("N3", 30.00, 3.90, seller="乙", date="2026-02-06"),
    ]
    by_seller = group_stats(invoices, "销售方")
    assert [(s.key, s.count, round(s.amount, 2)) for s in by_seller] == [
        ("甲", 2, 150.00),
        ("乙", 1, 30.00),
    ]
    assert by_seller[0].share == pytest.approx(150 / 180)

    # 月份分组同样按金额降序：2026-01 只有一张但金额更大，排在前面
    by_month = group_stats(invoices, "月份")
    assert [(s.key, s.count) for s in by_month] == [("2026-01", 1), ("2026-02", 2)]


def test_group_stats_fallbacks_for_missing_values() -> None:
    inv = make_invoice("", 10.00, 1.30, seller="", buyer="", date="")
    assert group_stats([inv], "销售方")[0].key == "(未识别销售方)"
    assert group_stats([inv], "购买方")[0].key == "(未识别购买方)"
    assert group_stats([inv], "月份")[0].key == "(无开票日期)"
    assert group_stats([inv], "税率")[0].key == "(无税率)"


def test_group_stats_handles_empty_input() -> None:
    assert group_stats([], "销售方") == []


def test_group_stats_share_sums_to_one() -> None:
    invoices = [make_invoice(f"N{i}", 10.0 * i, 0, seller=f"卖方{i % 3}") for i in range(1, 7)]
    stats = group_stats(invoices, "销售方")
    assert sum(s.share for s in stats) == pytest.approx(1.0)


def test_dominant_rate_picks_most_common() -> None:
    items = [
        make_item("1.00", rate="13%"),
        make_item("1.00", rate="13%"),
        make_item("1.00", rate="6%"),
    ]
    assert dominant_rate(items) == "13%"
    assert dominant_rate([]) == ""


# ---------------------------------------------------------------- 品目与折扣摊回


def test_product_stats_allocates_discount_back_to_previous_row() -> None:
    items = [
        make_item("393.63", "51.17", name="*医疗*测试品目甲"),
        make_item("-180.55", "-23.47", name="*医疗*测试品目甲", item_no=2),
        make_item("179.58", "23.34", name="*医疗*测试品目乙", item_no=3),
        make_item("-79.16", "-10.29", name="*医疗*测试品目乙", item_no=4),
    ]
    inv = make_invoice("N1", 313.50, 40.75, items=items)
    stats = {s.key: s for s in product_stats([inv])}
    assert stats["测试品目甲"].amount == pytest.approx(213.08)
    assert stats["测试品目甲"].tax == pytest.approx(27.70)
    assert stats["测试品目甲"].line_count == 2
    assert stats["测试品目乙"].amount == pytest.approx(100.42)
    assert stats["测试品目乙"].tax == pytest.approx(13.05)


def test_product_stats_total_equals_invoice_amount() -> None:
    items = [
        make_item("393.63", "51.17", name="*医疗*测试品目甲"),
        make_item("-180.55", "-23.47", name="*医疗*测试品目甲", item_no=2),
        make_item("100.00", "13.00", name="*办公*测试品目乙", item_no=3),
    ]
    inv = make_invoice("N1", 313.08, 40.70, items=items)
    stats = product_stats([inv])
    assert sum(s.amount for s in stats) == pytest.approx(313.08)
    assert sum(s.tax for s in stats) == pytest.approx(40.70)
    for stat in stats:
        assert stat.total == pytest.approx(round(stat.amount + stat.tax, 2))


def test_product_stats_orphan_discount_does_not_crash() -> None:
    items = [make_item("-10.00", "-1.30", name="*医疗*孤立折扣")]
    stats = product_stats([make_invoice("N1", -10.00, -1.30, items=items)])
    assert len(stats) == 1
    assert stats[0].amount == pytest.approx(-10.00)


def test_product_stats_merges_prefix_variants() -> None:
    items = [
        make_item("100.00", "13.00", name="*测试类别*测试品目甲"),
        make_item("50.00", "6.50", name="*测试类别*测试品目甲套装", item_no=2),
    ]
    stats = product_stats([make_invoice("N1", 150.00, 19.50, items=items)])
    # 「套装」属于规格噪声，归一后两者都收敛成「测试品目甲」，合并为一条
    assert len(stats) == 1
    assert stats[0].key == "测试品目甲"
    assert stats[0].invoice_count == 1


def test_product_stats_counts_distinct_invoices() -> None:
    a = make_invoice("N1", 100.00, 13.00, items=[make_item("100.00", "13.00", name="*x*品目A")])
    b = make_invoice("N2", 100.00, 13.00, items=[make_item("100.00", "13.00", name="*x*品目A")])
    stats = product_stats([a, b])
    assert len(stats) == 1
    assert stats[0].invoice_count == 2
    assert stats[0].line_count == 2


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("*测试类别*某商品", "某商品"),
        ("*类别*商品（说明文字）", "商品"),
        ("商品 规格 适用于某型号", "商品规格"),
        ("普通品名", "普通品名"),
        ("", ""),
    ],
)
def test_normalize_product_name(raw: str, expected: str) -> None:
    """归一化只影响展示与聚合口径，不能改动金额；这里锁定当前行为。"""
    assert normalize_product_name(raw) == expected


# ---------------------------------------------------------------- 排序


def test_sort_invoices_by_date_then_number() -> None:
    a = make_invoice("9", 1, 0, date="2026-02-01")
    b = make_invoice("1", 1, 0, date="2026-01-01")
    c = make_invoice("5", 1, 0, date="2026-01-01")
    d = make_invoice("", 1, 0, date="")
    ordered = sort_invoices([a, b, c, d])
    assert [i.invoice_no for i in ordered] == ["1", "5", "9", ""]


def test_tolerance_is_two_cents() -> None:
    assert TOLERANCE == 0.02
