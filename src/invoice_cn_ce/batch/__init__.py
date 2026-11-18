# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""批量导入与汇总（``python -m invoice_cn_ce.batch``）。

与上游单文件端点并存：单文件解析仍走 :func:`invoice_cn_ce.formats.parse_pdf`，
本子包只增加「批量输入 / 校验 / 去重 / 汇总 / 报表」这几层，不改既有入口行为。

模块划分：

===================  ================================================
``ingest``           输入收集：文件 / 目录递归 / ZIP（含嵌套）、过滤、解包、hash 去重
``aggregate``        逐票校验、按发票号去重、合计、分组、品目归一化
``report``           xlsx（五个 sheet）/ csv / json 输出
``cli``              命令行入口、终端摘要
===================  ================================================
"""

from __future__ import annotations

from invoice_cn_ce.batch.aggregate import (
    BatchInvoice,
    DedupInfo,
    GroupStat,
    ProductStat,
    Totals,
    build_invoice,
    dedupe,
    group_stats,
    parse_group_option,
    product_stats,
    sort_invoices,
    totals,
    validate,
)
from invoice_cn_ce.batch.ingest import Ingestor, IngestResult, SourceFile

__all__ = [
    "BatchInvoice",
    "DedupInfo",
    "GroupStat",
    "IngestResult",
    "Ingestor",
    "ProductStat",
    "SourceFile",
    "Totals",
    "build_invoice",
    "dedupe",
    "group_stats",
    "parse_group_option",
    "product_stats",
    "sort_invoices",
    "totals",
    "validate",
]
