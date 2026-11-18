# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""PDF 明细行抽取：版面重建、通用列定位、合计反推、兜底不丢行。

**全部素材是内存里造的词坐标**，不含任何真实票据。这里刻意用「与真实票样不同」
的坐标与列宽，验证生产代码确实在做通用定位，而不是写死某一种版式。
"""

from __future__ import annotations

import pytest

from invoice_cn_ce.parsers.items import (
    COLUMN_ORDER,
    MAX_TRUSTED_COL_WIDTH,
    _normalize_rate,
    assign_line,
    build_lines,
    detect_header_anchors,
    extract_items,
    extract_page_items,
    find_totals_line,
    join_cell,
    parse_totals_line,
    split_word,
    to_words,
)
from tests.conftest import standard_invoice_words, word


class FakePage:
    """最小 PageView 替身：只需 ``words()`` 与 ``_page.width``。"""

    class _Inner:
        width = 600.0

    def __init__(self, words: list) -> None:
        self._words = words
        self._page = self._Inner()

    def words(self) -> list:
        return self._words


# ---------------------------------------------------------------- 基本抽取


def test_extracts_single_item_from_standard_layout() -> None:
    items, note = extract_page_items(standard_invoice_words(), 600.0)
    assert note == ""
    assert len(items) == 1
    item = items[0]
    assert item.amount == "393.63"
    assert item.tax_amount == "51.17"
    assert item.tax_rate == "13%"
    assert item.unit == "盒"
    assert item.quantity == "2"
    assert item.unit_price == "196.815"  # 三位小数原样保留
    assert "测试商品甲" in item.name
    assert item.row_type == "商品"
    assert item.is_reversal is False
    assert item.raw_line


def test_same_table_at_shifted_coordinates_gives_same_values() -> None:
    """整体平移 / 拉伸坐标后结果不变 —— 证明没有写死票样坐标。"""
    base = standard_invoice_words()
    for dx in (0.0, 17.0, -11.0):
        shifted = [dict(w, x0=w["x0"] + dx, x1=w["x1"] + dx) for w in base]
        items, _note = extract_page_items(shifted, 600.0)
        assert len(items) == 1
        assert (items[0].amount, items[0].tax_amount, items[0].tax_rate) == (
            "393.63",
            "51.17",
            "13%",
        )


def test_header_split_per_character_is_still_matched() -> None:
    """全电发票的表头常被切成单字（``金`` + ``额``），必须合并后再匹配。"""
    words = [
        word("项", 59, 68, 146),
        word("目", 68, 77, 146),
        word("名", 77, 86, 146),
        word("称", 86, 95, 146),
        word("数", 308, 317, 146),
        word("量", 319, 328, 146),
        word("单", 379, 388, 147),  # 表头跨两行，靠「表头带」聚合
        word("价", 390, 399, 147),
        word("金", 450, 459, 147),
        word("额", 464, 473, 147),
        word("税", 568, 577, 147),
        word("额", 582, 591, 147),
        word("测试商品甲", 26, 149, 155),
        word("2", 324, 328, 155),
        word("196.815", 371, 399, 155),
        word("393.63", 449, 473, 155),
        word("13%", 492, 502, 155),
        word("51.17", 572, 592, 155),
    ]
    items, _note = extract_page_items(words, 600.0)
    assert len(items) == 1
    assert items[0].amount == "393.63"
    assert items[0].tax_amount == "51.17"


def test_detect_header_anchors_returns_expected_columns() -> None:
    lines = build_lines(to_words(standard_invoice_words()))
    anchors = detect_header_anchors(lines, 250.0)
    assert {"name", "spec", "unit", "qty", "price", "amount", "rate", "tax"} <= set(anchors)


# ---------------------------------------------------------------- 折扣行 / 续行


def test_discount_row_is_kept_as_negative_row() -> None:
    words = standard_invoice_words()
    words += [
        word("测试商品甲", 26, 149, 180),
        word("-180.55", 445, 473, 180),
        word("13%", 492, 502, 180),
        word("-23.47", 568, 592, 180),
    ]
    items, _note = extract_page_items(words, 600.0)
    assert len(items) == 2
    discount = items[1]
    assert discount.amount == "-180.55"
    assert discount.tax_amount == "-23.47"
    assert discount.is_reversal is True
    assert discount.row_type == "折扣"
    assert float(items[0].amount) + float(discount.amount) == pytest.approx(213.08)


def test_wrapped_name_is_merged_into_the_same_row() -> None:
    words = standard_invoice_words()
    words += [
        word("（续行说明）", 26, 120, 163),
        word("第二段规格", 159, 221, 163),
    ]
    items, _note = extract_page_items(words, 600.0)
    assert len(items) == 1
    assert "续行说明" in items[0].name
    assert "第二段规格" in items[0].spec


def test_wrapped_text_after_discount_row_does_not_invent_new_row() -> None:
    words = standard_invoice_words()
    words += [
        word("测试商品甲", 26, 149, 180),
        word("-180.55", 445, 473, 180),
        word("13%", 492, 502, 180),
        word("-23.47", 568, 592, 180),
        word("（折扣说明）", 26, 120, 188),
    ]
    items, _note = extract_page_items(words, 600.0)
    assert len(items) == 2  # 续行不产生第三条明细


def test_stop_words_are_not_treated_as_items() -> None:
    words = standard_invoice_words()
    words += [word("备注", 26, 60, 300), word("收款人", 26, 60, 310)]
    items, _note = extract_page_items(words, 600.0)
    assert len(items) == 1


# ---------------------------------------------------------------- 兜底通路


def test_page_without_any_table_returns_no_items() -> None:
    words = [word("这是一段普通文字", 26, 200, 100), word("没有表格", 26, 120, 120)]
    items, note = extract_page_items(words, 600.0)
    assert items == []
    assert note == "no_table"


def test_page_with_no_words_returns_no_table() -> None:
    items, note = extract_page_items([], 600.0)
    assert (items, note) == ([], "no_table")


def test_totals_only_layout_falls_back_to_boundary_split() -> None:
    """没有表头锚点、只有「合计」行时：金额/税额仍要取到，并标记低置信度。"""
    words = [
        word("测试商品甲", 26, 300, 155),
        word("393.63", 445, 473, 155),
        word("13%", 492, 502, 155),
        word("51.17", 568, 592, 155),
        word("合计", 61, 111, 256),
        word("¥393.63", 413, 473, 256),
        word("¥51.17", 565, 592, 256),
    ]
    items, note = extract_page_items(words, 600.0)
    assert len(items) == 1
    assert items[0].amount == "393.63"
    assert items[0].tax_amount == "51.17"
    assert note == "low_confidence"


def test_allow_fallback_false_returns_nothing_without_trusted_columns() -> None:
    words = [
        word("测试商品甲", 26, 300, 155),
        word("393.63", 445, 473, 155),
        word("13%", 492, 502, 155),
        word("51.17", 568, 592, 155),
    ]
    items, note = extract_page_items(words, 600.0, allow_fallback=False)
    assert items == []
    assert note == "no_table"


def test_too_wide_reverse_column_is_not_trusted() -> None:
    """反推出来的列过宽（区间不可信）时不许当作高置信度使用。"""
    words = [
        word("测试商品甲", 26, 300, 155),
        word("393.63", 340, 360, 155),
        word("13%", 492, 502, 155),
        word("51.17", 568, 592, 155),
        word("合计", 61, 111, 256),
        word("¥393.63", 320, 360, 256),
        word("¥51.17", 565, 592, 256),
    ]
    _items, note = extract_page_items(words, 600.0)
    assert note in {"", "low_confidence"}
    assert MAX_TRUSTED_COL_WIDTH == 150.0


# ---------------------------------------------------------------- 合计行


def test_find_totals_line_skips_price_total_line() -> None:
    lines = build_lines(
        to_words(
            [
                word("价税合计(大写)", 56, 116, 277),
                word("合计", 61, 111, 258),
                word("¥240.78", 456, 487, 276),
            ]
        )
    )
    found = find_totals_line(lines)
    assert found is not None
    assert "价税合计" not in found.flat
    assert "合计" in found.flat


def test_find_totals_line_merges_split_characters() -> None:
    lines = build_lines(to_words([word("合", 61, 70, 258), word("计", 102, 111, 258)]))
    assert find_totals_line(lines) is not None


def test_parse_totals_line_reads_amount_and_tax() -> None:
    lines = build_lines(
        to_words(
            [
                word("合计", 61, 111, 258),
                word("¥213.08", 413, 473, 258),
                word("¥27.70", 565, 592, 258),
            ]
        )
    )
    pick = parse_totals_line(find_totals_line(lines))
    assert pick.found is True
    assert (pick.amount, pick.tax) == (213.08, 27.70)
    assert pick.amount_right and pick.tax_right


def test_parse_totals_line_without_line() -> None:
    pick = parse_totals_line(None)
    assert pick.found is False
    assert (pick.amount, pick.tax) == (None, None)


# ---------------------------------------------------------------- 单元格工具


def test_join_cell_inserts_space_on_large_gap() -> None:
    words = to_words([word("甲", 10, 20, 1), word("乙", 21, 30, 1), word("丙", 60, 70, 1)])
    assert join_cell(words) == "甲乙 丙"


def test_split_word_splits_multi_cell_text_by_char_ratio() -> None:
    pieces = split_word(to_words([word("甲乙 丙丁", 100, 200, 1)])[0])
    assert [p.text for p in pieces] == ["甲乙", "丙丁"]
    assert pieces[0].x0 == pytest.approx(100.0)
    assert pieces[1].x0 > pieces[0].x1


def test_split_word_keeps_single_chunk_intact() -> None:
    single = to_words([word("甲乙丙", 100, 200, 1)])[0]
    assert [p.text for p in split_word(single)] == ["甲乙丙"]


def test_assign_line_puts_words_in_columns() -> None:
    lines = build_lines(to_words(standard_invoice_words()))
    anchors = detect_header_anchors(lines, 250.0)
    from invoice_cn_ce.parsers.items import build_columns

    cols = build_columns(anchors, 600.0)
    detail = next(line for line in lines if abs(line.top - 155) < 1)
    cells = assign_line(detail, cols)
    assert set(cells) == set(COLUMN_ORDER)
    assert "13%" in "".join(w.text for w in cells["rate"])
    assert cols, "列模型不能为空"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [("13%", "13%"), ("13 %", "13%"), ("免税", "免税"), ("不征税", "不征税"), ("*", "*"), ("", "")],
)
def test_normalize_rate(raw: str, expected: str) -> None:
    assert _normalize_rate(raw) == expected


# ---------------------------------------------------------------- 跨页


def test_extract_items_concatenates_pages_and_renumbers() -> None:
    page1 = FakePage(standard_invoice_words())
    page2 = FakePage(
        standard_invoice_words(name="测试商品乙", amount="100.00", tax="13.00", with_header=False)
    )
    items, note = extract_items([page1, page2])
    assert [i.item_no for i in items] == list(range(1, len(items) + 1))
    assert len(items) >= 2
    assert note in {"", "low_confidence"}


def test_extract_items_survives_broken_page() -> None:
    class Boom:
        class _Inner:
            width = 600.0

        _page = _Inner()

        def words(self):
            raise RuntimeError("pdfplumber exploded")

    items, note = extract_items([Boom(), FakePage(standard_invoice_words())])
    assert len(items) == 1
    assert note in {"", "low_confidence"}


def test_extract_items_with_no_pages() -> None:
    items, note = extract_items([])
    assert (items, note) == ([], "no_table")
