# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""批量汇总功能的公共测试装置。

**只使用动态生成的最小素材**：本文件不包含任何真实发票、公司名或税号。
需要真实票据的用例通过 ``INVOICE_FIXTURES`` 环境变量指向本地样本目录，
变量缺省或目录为空时自动 ``skip``，公开 CI 只跑合成素材部分。
"""

from __future__ import annotations

import io
import os
import sys
import zipfile
from pathlib import Path
from typing import Iterable, Sequence, Tuple

import pytest

# 允许不安装（未 ``pip install -e .``）时直接 ``python -m pytest tests``：
# 把仓库的 ``src/`` 前置到导入路径。已安装时这一行是无副作用的兜底。
_SRC = str(Path(__file__).resolve().parents[1] / "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from invoice_cn_ce.batch.aggregate import BatchInvoice  # noqa: E402
from invoice_cn_ce.batch.ingest import SourceFile  # noqa: E402
from invoice_cn_ce.schemas import InvoiceItem  # noqa: E402

#: 真实发票目录（本地验证用，不入库）。指向包含若干电子发票 PDF 的目录。
FIXTURES_ENV = "INVOICE_FIXTURES"


def real_pdf_dir() -> Path:
    """返回真实样本目录；不可用时 skip 当前用例。"""
    raw = os.environ.get(FIXTURES_ENV, "").strip()
    if not raw:
        pytest.skip(f"未设置 {FIXTURES_ENV}，跳过依赖真实票据的用例")
    path = Path(raw)
    if not path.is_dir():
        pytest.skip(f"{FIXTURES_ENV} 指向的目录不存在：{path}")
    if not any(path.glob("*.pdf")):
        pytest.skip(f"{FIXTURES_ENV} 目录里没有 PDF：{path}")
    return path


# ---------------------------------------------------------------- ZIP 素材


def zip_bytes(entries: Sequence[Tuple[str, bytes]]) -> bytes:
    """把 ``[(包内路径, 内容)]`` 打成内存 ZIP。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, payload in entries:
            zf.writestr(name, payload)
    return buf.getvalue()


def write_zip(path: Path, entries: Sequence[Tuple[str, bytes]]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(zip_bytes(entries))
    return path


def nested_zip_bytes(depth: int, leaf_name: str, payload: bytes) -> bytes:
    """生成 ``depth`` 层嵌套的 ZIP，最内层放 ``leaf_name``。"""
    data = payload
    name = leaf_name
    for _ in range(depth):
        data = zip_bytes([(name, data)])
        name = "inner.zip"
    return data


#: 不是 PDF 的最小内容（用于验证「绝不产生假票」）。
NOT_A_PDF = b"this is not a pdf at all"


# ---------------------------------------------------------------- 数据模型工厂


def make_item(
    amount: str,
    tax: str = "",
    *,
    name: str = "测试品目",
    rate: str = "13%",
    item_no: int = 1,
) -> InvoiceItem:
    """构造一条明细行；``amount`` 为负即折扣/冲销行。"""
    negative = amount.startswith("-")
    return InvoiceItem(
        item_no=item_no,
        name=name,
        category="",
        amount=amount,
        tax_rate=rate,
        tax_amount=tax,
        is_reversal=negative,
        row_type="折扣" if negative else "商品",
    )


def make_invoice(
    number: str,
    amount: float,
    tax: float,
    *,
    items: Iterable[InvoiceItem] = (),
    seller: str = "测试销售方",
    buyer: str = "测试购买方",
    date: str = "2026-01-15",
    kind: str = "电子发票（普通发票）",
    display: str = "sample.pdf",
) -> BatchInvoice:
    """构造一张 :class:`BatchInvoice`，不经过 PDF 解析。"""
    return BatchInvoice(
        source=SourceFile(path=display, display=display, sha1=number, size=1),
        kind=kind,
        invoice_no=number,
        invoice_date=date,
        buyer_name=buyer,
        buyer_tax_id="91000000000000000X",
        seller_name=seller,
        seller_tax_id="91000000000000000Y",
        amount=round(amount, 2),
        tax=round(tax, 2),
        total=round(amount + tax, 2),
        items=list(items),
    )


# ---------------------------------------------------------------- PDF 版面素材


def word(text: str, x0: float, x1: float, top: float) -> dict:
    """构造一个 pdfplumber 风格的词字典（``extract_words()`` 的元素形态）。"""
    return {"text": text, "x0": x0, "x1": x1, "top": top, "bottom": top + 10.0}


# ---------------------------------------------------------------- CLI 子进程


def run_cli(args: Sequence[str]) -> "subprocess.CompletedProcess[str]":
    """以子进程方式跑 ``python -m invoice_cn_ce.batch``，返回完整结果。

    用子进程而不是直接调 ``run()``，这样退出码、stdout/stderr 与真实用法一致
    （包括「不 traceback」「中文不乱码」这类只能从进程外看到的性质）。
    """
    import subprocess

    env = dict(os.environ)
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = _SRC + (os.pathsep + existing if existing else "")
    env["PYTHONIOENCODING"] = "utf-8"
    env["NO_COLOR"] = "1"
    return subprocess.run(
        [sys.executable, "-m", "invoice_cn_ce.batch", *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        timeout=180,
    )


def standard_invoice_words(
    *,
    name: str = "测试商品甲",
    amount: str = "393.63",
    tax: str = "51.17",
    rate: str = "13%",
    with_header: bool = True,
    with_totals: bool = True,
) -> list:
    """模拟一张「表头 + 一行明细 + 合计」的电子发票版面。

    坐标是**测试自己的版面**，与任何真实票样无关；列位置显式写在这里，正是为了
    证明生产代码不是靠这些坐标吃饭的（换一套坐标仍应解出同样的值）。
    """
    rows: list = []
    if with_header:
        rows += [
            word("项目名称", 59, 95, 146),
            word("规格型号", 160, 196, 146),
            word("单位", 234, 252, 146),
            word("数量", 308, 328, 146),
            word("单价", 379, 399, 146),
            word("金额", 450, 473, 146),
            word("税率/征收率", 478, 528, 146),
            word("税额", 568, 592, 146),
        ]
    rows += [
        word(name, 26, 149, 155),
        word("规格A", 159, 221, 155),
        word("盒", 239, 246, 155),
        word("2", 324, 328, 155),
        word("196.815", 371, 399, 155),
        word(amount, 449, 473, 155),
        word(rate, 492, 502, 155),
        word(tax, 572, 592, 155),
    ]
    if with_totals:
        rows += [
            word("合计", 61, 111, 256),
            word(f"¥{amount}", 413, 473, 256),
            word(f"¥{tax}", 565, 592, 256),
        ]
    return rows
