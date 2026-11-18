# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""批量汇总层：逐票校验、按发票号去重、分组统计、品目归一化。

这一层只做**数据加工**，不碰 PDF：输入是 ``formats.parse_pdf`` 返回的
:class:`invoice_cn_ce.schemas.ExtractResponse` 加上来源信息，输出是可直接喂给
报表的数据结构。

原则与解析层一致：**宁可标记，不可错算**。金额字符串解析不出来就留 ``None``，
校验差额超过容差就标 ``has_issue`` 并把原因写清楚，不静默当 0。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Tuple

from invoice_cn_ce.batch.ingest import SourceFile
from invoice_cn_ce.parsers.items import UNIT_WORDS
from invoice_cn_ce.schemas import InvoiceItem

#: 金额校验容差（元）。票面两位小数，累加误差不应超过 2 分。
TOLERANCE = 0.02

#: 可用的分组维度。
GROUP_CHOICES = ("销售方", "购买方", "月份", "税率", "品目")
DEFAULT_GROUP = "销售方"

#: 品名里的「规格噪声」标记，从第一个标记处截断。这些是**版式用语**（说明规格的
#: 起头词），不是某类商品的专名，换一张票也成立。
_SPEC_MARKERS = ("适用于", "适用", "无仪器", "不含仪器", "瓶装", "盒装", "袋装",
                 "套装", "简装", "标准装", "零售装")
_PAREN_RE = re.compile(r"[（(][^（()）]*[）)]")
_CATEGORY_RE = re.compile(r"^\*([^*]{1,20})\*(.*)$", re.S)

#: 尾缀「数量 + 计量单位」的通用形态（``200支`` / ``300片`` / ``10盒`` / ``1.5千克``）。
#: 单位表复用解析层的 :data:`invoice_cn_ce.parsers.items.UNIT_WORDS` —— 只认计量
#: 单位，**不含任何商品名**，避免把「试纸 / 试条 / 家用」这类业务词写死进正则。
_UNIT_ALT = "|".join(
    sorted((re.escape(u) for u in UNIT_WORDS), key=len, reverse=True)
)
#: ``200支试纸`` 这种「数量 + 单位 + 尾巴」整体切掉（单位后面还跟文字的形态）。
_QTY_TAIL_RE = re.compile(rf"\s*\d+(?:\.\d+)?(?:{_UNIT_ALT})[\s\S]*$")
#: 纯尾缀：``200支`` / ``1.5千克``，或紧跟在汉字后的裸数量 ``某商品300``。
#: 裸数量必须紧跟汉字，避免把型号后缀（``打印纸A4``）当数量切掉。
_TRAILING_QTY_RE = re.compile(
    rf"(?:\s*\d+(?:\.\d+)?(?:{_UNIT_ALT})|(?<=[\u4e00-\u9fff])\d+(?:\.\d+)?)$"
)


# ---------------------------------------------------------------- 数值工具


def to_float(value: object) -> Optional[float]:
    """把票面字符串金额转成 float；空串 / 非数字返回 ``None``。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    text = str(value).replace("¥", "").replace("￥", "").replace(",", "").strip()
    text = re.sub(r"\s+", "", text)
    if not text:
        return None
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def round2(value: Optional[float]) -> Optional[float]:
    return None if value is None else round(float(value) + 1e-9, 2)


def money(value: Optional[float]) -> float:
    """求合计时把 ``None`` 当 0（只用于汇总，不用于单票字段）。"""
    return 0.0 if value is None else float(value)


# ---------------------------------------------------------------- 行模型


@dataclass
class BatchInvoice:
    """一张票在批量视图下的样子。"""

    source: SourceFile = field(default_factory=SourceFile)
    kind: str = ""
    invoice_no: str = ""
    invoice_date: str = ""
    buyer_name: str = ""
    buyer_tax_id: str = ""
    seller_name: str = ""
    seller_tax_id: str = ""
    amount: Optional[float] = None
    tax: Optional[float] = None
    total: Optional[float] = None
    rate: str = ""
    items: List[InvoiceItem] = field(default_factory=list)
    message: str = ""
    has_issue: bool = False
    issues: List[str] = field(default_factory=list)
    #: 解析器给出的票种键（new / old / xml / ofd …），仅用于排查。
    format_key: str = ""

    @property
    def deductible(self) -> bool:
        """可抵扣进项税：专票（含全电专票）可抵扣，普票不可。"""
        return "专用" in (self.kind or "")

    @property
    def file_label(self) -> str:
        return self.source.display or self.source.path

    def item_rows(self) -> List[InvoiceItem]:
        return list(self.items)

    def to_dict(self) -> dict:
        return {
            "file": self.file_label,
            "archive": self.source.archive,
            "inner_path": self.source.inner,
            "format": self.format_key,
            "invoice_type": self.kind,
            "deductible": self.deductible,
            "invoice_number": self.invoice_no,
            "issue_date": self.invoice_date,
            "buyer_name": self.buyer_name,
            "buyer_tax_id": self.buyer_tax_id,
            "seller_name": self.seller_name,
            "seller_tax_id": self.seller_tax_id,
            "amount_without_tax": self.amount,
            "tax_amount": self.tax,
            "total_amount": self.total,
            "tax_rate": self.rate,
            "item_count": len(self.items),
            "has_issue": self.has_issue,
            "issues": list(self.issues),
            "message": self.message,
            "items": [it.model_dump() for it in self.items],
        }


def build_invoice(resp: object, source: SourceFile) -> BatchInvoice:
    """``ExtractResponse`` -> :class:`BatchInvoice`（并做票级校验）。"""
    inv = BatchInvoice(
        source=source,
        kind=getattr(resp, "invoice_type", "") or "",
        invoice_no=getattr(resp, "invoice_number", "") or "",
        invoice_date=getattr(resp, "issue_date", "") or "",
        buyer_name=getattr(resp, "buyer_name", "") or "",
        buyer_tax_id=getattr(resp, "buyer_tax_id", "") or "",
        seller_name=getattr(resp, "seller_name", "") or "",
        seller_tax_id=getattr(resp, "seller_tax_id", "") or "",
        amount=to_float(getattr(resp, "amount_without_tax", "")),
        tax=to_float(getattr(resp, "tax_amount", "")),
        total=to_float(getattr(resp, "total_amount", "")),
        items=list(getattr(resp, "items", []) or []),
        message=getattr(resp, "message", "") or "",
    )
    inv.rate = dominant_rate(inv.items)
    return validate(inv)


def validate(inv: BatchInvoice, tol: float = TOLERANCE) -> BatchInvoice:
    """票级勾稽：Σ明细 vs 票面金额 / 税额；金额+税额 vs 价税合计。"""
    issues: List[str] = []
    items = inv.items
    if not items:
        issues.append("未解析到明细行（PDF 明细依赖坐标重建，扫描件/特殊版式会缺失）")
    else:
        missing = [it for it in items if to_float(it.amount) is None]
        if missing:
            issues.append(f"有 {len(missing)} 行明细未取到金额")
        sum_amount = round2(sum(money(to_float(it.amount)) for it in items)) or 0.0
        sum_tax = round2(sum(money(to_float(it.tax_amount)) for it in items)) or 0.0
        if inv.amount is not None:
            diff = round2(sum_amount - inv.amount) or 0.0
            if abs(diff) > tol:
                issues.append(
                    f"明细金额合计 {sum_amount:.2f} 与票面金额 {inv.amount:.2f} 差 {diff:+.2f}"
                )
        if inv.tax is not None:
            diff = round2(sum_tax - inv.tax) or 0.0
            if abs(diff) > tol:
                issues.append(
                    f"明细税额合计 {sum_tax:.2f} 与票面税额 {inv.tax:.2f} 差 {diff:+.2f}"
                )

    if inv.amount is None:
        issues.append("未取到票面金额（不含税）")
    if inv.tax is None:
        issues.append("未取到票面税额")
    if inv.total is None:
        issues.append("未取到价税合计")
    if inv.amount is not None and inv.tax is not None and inv.total is not None:
        diff = round2(inv.amount + inv.tax - inv.total) or 0.0
        if abs(diff) > tol:
            issues.append(
                f"金额+税额 {inv.amount + inv.tax:.2f} 与价税合计 {inv.total:.2f} 差 {diff:+.2f}"
            )
    if inv.message and inv.message != "success":
        issues.append(f"解析器返回：{inv.message}")
    if not inv.invoice_no:
        issues.append("未识别到发票号码")

    inv.issues = _dedup_keep_order(issues)
    inv.has_issue = bool(inv.issues)
    return inv


def _dedup_keep_order(values: Iterable[str]) -> List[str]:
    seen: set = set()
    out: List[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


# ---------------------------------------------------------------- 去重


@dataclass
class DedupInfo:
    """去重结果。"""

    by_hash: List[Tuple[str, str]] = field(default_factory=list)
    by_number: List[Tuple[str, str]] = field(default_factory=list)
    total: int = 0
    kept: int = 0

    @property
    def duplicate_count(self) -> int:
        return len(self.by_hash) + len(self.by_number)


def dedupe(invoices: Sequence[BatchInvoice], hash_dups: Sequence[Tuple[str, str]] = (),
           enabled: bool = True) -> Tuple[List[BatchInvoice], DedupInfo]:
    """按发票号码去重（同号不同文件视为重复票）。"""
    info = DedupInfo(by_hash=list(hash_dups), total=len(invoices))
    if not enabled:
        info.kept = len(invoices)
        return list(invoices), info
    by_no: dict = {}
    kept: List[BatchInvoice] = []
    for inv in invoices:
        key = (inv.invoice_no or "").strip()
        if key:
            first = by_no.get(key)
            if first is not None:
                info.by_number.append((inv.file_label, first.file_label))
                continue
            by_no[key] = inv
        kept.append(inv)
    info.kept = len(kept)
    return kept, info


# ---------------------------------------------------------------- 合计


@dataclass
class Totals:
    """全量合计 + 勾稽结论。"""

    count: int = 0
    amount: float = 0.0
    tax: float = 0.0
    total: float = 0.0
    item_count: int = 0
    discount: float = 0.0
    sum_amount: float = 0.0
    sum_tax: float = 0.0
    sum_total: float = 0.0
    reconciled: bool = True
    notes: List[str] = field(default_factory=list)


def totals(invoices: Sequence[BatchInvoice]) -> Totals:
    t = Totals(count=len(invoices))
    t.amount = round2(sum(money(i.amount) for i in invoices)) or 0.0
    t.tax = round2(sum(money(i.tax) for i in invoices)) or 0.0
    t.total = round2(sum(money(i.total) for i in invoices)) or 0.0
    t.item_count = sum(len(i.items) for i in invoices)
    t.discount = round2(
        sum(money(to_float(it.amount)) for i in invoices for it in i.items
            if (to_float(it.amount) or 0) < 0)
    ) or 0.0
    # 逐票先各自取两位，再累加 —— 与「按票相加」的口径完全一致。
    t.sum_amount = round2(sum(round2(money(i.amount)) or 0.0 for i in invoices)) or 0.0
    t.sum_tax = round2(sum(round2(money(i.tax)) or 0.0 for i in invoices)) or 0.0
    t.sum_total = round2(sum(round2(money(i.total)) or 0.0 for i in invoices)) or 0.0

    notes: List[str] = []
    if abs(t.amount - t.sum_amount) > TOLERANCE:
        notes.append(f"金额合计 {t.amount:.2f} ≠ 各票金额之和 {t.sum_amount:.2f}")
    if abs(t.tax - t.sum_tax) > TOLERANCE:
        notes.append(f"税额合计 {t.tax:.2f} ≠ 各票税额之和 {t.sum_tax:.2f}")
    if abs(t.total - t.sum_total) > TOLERANCE:
        notes.append(f"价税合计 {t.total:.2f} ≠ 各票价税合计之和 {t.sum_total:.2f}")
    if invoices and abs(t.amount + t.tax - t.total) > TOLERANCE * len(invoices):
        notes.append(f"金额+税额 {t.amount + t.tax:.2f} ≠ 价税合计 {t.total:.2f}")
    t.notes = notes
    t.reconciled = not notes
    return t


# ---------------------------------------------------------------- 分组


@dataclass
class GroupStat:
    key: str = ""
    count: int = 0
    amount: float = 0.0
    tax: float = 0.0
    total: float = 0.0
    share: float = 0.0


def dominant_rate(items: Sequence[InvoiceItem]) -> str:
    """主要税率：商品行里出现次数最多的税率。"""
    counter: dict = {}
    for item in items:
        rate = (item.tax_rate or "").strip()
        if not rate:
            continue
        counter[rate] = counter.get(rate, 0) + 1
    if not counter:
        return ""
    return max(counter.items(), key=lambda kv: (kv[1], kv[0]))[0]


def group_key(inv: BatchInvoice, dimension: str) -> str:
    if dimension == "销售方":
        return inv.seller_name or "(未识别销售方)"
    if dimension == "购买方":
        return inv.buyer_name or "(未识别购买方)"
    if dimension == "月份":
        return inv.invoice_date[:7] if inv.invoice_date else "(无开票日期)"
    if dimension == "税率":
        return inv.rate or "(无税率)"
    if dimension == "品目":
        return dominant_product(inv) or "(未识别品目)"
    return "(全部)"


def group_stats(invoices: Sequence[BatchInvoice], dimension: str) -> List[GroupStat]:
    bucket: dict = {}
    for inv in invoices:
        key = group_key(inv, dimension)
        stat = bucket.get(key)
        if stat is None:
            stat = GroupStat(key=key)
            bucket[key] = stat
        stat.count += 1
        stat.amount = round2(stat.amount + money(inv.amount)) or 0.0
        stat.tax = round2(stat.tax + money(inv.tax)) or 0.0
        stat.total = round2(stat.total + money(inv.total)) or 0.0
    total_amount = sum(s.amount for s in bucket.values()) or 0.0
    for stat in bucket.values():
        stat.share = (stat.amount / total_amount) if total_amount else 0.0
    return sorted(bucket.values(), key=lambda s: (-s.amount, s.key))


def dominant_product(inv: BatchInvoice) -> str:
    """一张票的主品目：净金额绝对值最大的商品行。"""
    best_key, best_value = "", -1.0
    for item in inv.items:
        amount = to_float(item.amount)
        if amount is None or amount < 0:
            continue
        key = product_key(item)
        if abs(amount) > best_value:
            best_key, best_value = key, abs(amount)
    return best_key


# ---------------------------------------------------------------- 品目


def normalize_product_name(name: str) -> str:
    """品目归一化：去 ``*类别*`` 前缀、去括号说明、去规格噪声、去尾部数量描述。

    规则全部是**通用形态**，不含任何商品专名：

    1. 去掉 ``*类别*`` 前缀；
    2. 去掉圆括号里的说明（``（不含仪器）`` 之类）；
    3. 从第一个「版式用语」处截断（``适用于…``、``瓶装…``、``套装`` …）；
    4. 去掉尾部的「数量 + 计量单位」（``200支``、``1.5千克``），以及紧跟在汉字
       后面的裸数量（``某商品300``）；型号后缀（``打印纸A4``）不动；
    5. 最后才合并空白 —— 先合并会把 ``A4 500张`` 粘成 ``A4500张``，把型号一起吃掉。

    改规则只影响展示与聚合口径，**不动金额**；跨票的同物异名（一个品名是另一个的
    前缀）由 :func:`_build_prefix_map` 再合并一次。
    """
    text = name or ""
    m = _CATEGORY_RE.match(text)
    if m:
        text = m.group(2)
    text = _PAREN_RE.sub("", text)
    positions = [text.find(marker) for marker in _SPEC_MARKERS]
    positions = [p for p in positions if 4 <= p]
    if positions:
        text = text[: min(positions)]
    text = _QTY_TAIL_RE.sub("", text)
    text = _TRAILING_QTY_RE.sub("", text)
    text = re.sub(r"\s+", "", text)
    return text.strip(" -*·、,，.。")


def product_key(item: InvoiceItem) -> str:
    core = normalize_product_name(item.name)
    if core:
        return core
    if item.category:
        return f"[{item.category}]"
    return "(未识别品目)"


def _build_prefix_map(keys: Sequence[str]) -> dict:
    """把「某品名是另一品名前缀」的键合并到最长那个（同物异名归一）。"""
    uniq = sorted(set(keys), key=len, reverse=True)
    mapping: dict = {}
    for key in keys:
        target = key
        for longer in uniq:  # 已按长度降序，首个命中即最长
            if len(longer) > len(key) and longer.startswith(key):
                target = longer
                break
        mapping[key] = target
    return mapping


@dataclass
class ProductStat:
    key: str = ""
    category: str = ""
    invoice_count: int = 0
    line_count: int = 0
    amount: float = 0.0
    tax: float = 0.0
    total: float = 0.0
    share: float = 0.0


def product_stats(invoices: Sequence[BatchInvoice]) -> List[ProductStat]:
    """品目汇总：折扣行摊回其上一条商品行，跨票按归一键聚合。"""
    raw_keys: List[str] = []
    rows: List[Tuple[str, InvoiceItem, str]] = []
    for inv in invoices:
        last_key = ""
        for item in inv.items:
            amount = to_float(item.amount)
            if amount is not None and amount < 0:
                key = last_key or f"(折扣){product_key(item)}"
            else:
                key = product_key(item)
                last_key = key
            raw_keys.append(key)
            rows.append((key, item, inv.invoice_no))
    mapping = _build_prefix_map(raw_keys)

    bucket: dict = {}
    invoices_of: dict = {}
    for key, item, invoice_no in rows:
        final = mapping.get(key, key)
        stat = bucket.get(final)
        if stat is None:
            stat = ProductStat(key=final)
            bucket[final] = stat
            invoices_of[final] = set()
        stat.line_count += 1
        stat.amount = round2(stat.amount + money(to_float(item.amount))) or 0.0
        stat.tax = round2(stat.tax + money(to_float(item.tax_amount))) or 0.0
        stat.category = item.category or stat.category
        invoices_of[final].add(invoice_no)

    total_amount = sum(s.amount for s in bucket.values()) or 0.0
    for key, stat in bucket.items():
        stat.share = (stat.amount / total_amount) if total_amount else 0.0
        stat.invoice_count = len(invoices_of[key])
        stat.total = round2(stat.amount + stat.tax) or 0.0
    return sorted(bucket.values(), key=lambda s: (-s.amount, s.key))


# ---------------------------------------------------------------- 其他


def sort_invoices(invoices: Sequence[BatchInvoice]) -> List[BatchInvoice]:
    """按开票日期 + 发票号码排序（缺号的排最后）。"""
    return sorted(invoices, key=lambda i: (i.invoice_date or "9999-99-99", i.invoice_no or "~"))


def parse_group_option(raw: str) -> List[str]:
    """解析 ``--group 购买方,月份``；非法维度抛 :class:`ValueError`。"""
    out: List[str] = []
    for part in re.split(r"[,，\s]+", raw or ""):
        part = part.strip()
        if not part:
            continue
        if part not in GROUP_CHOICES:
            raise ValueError(
                f"不支持的分组维度：{part}（可选：{'、'.join(GROUP_CHOICES)}）"
            )
        if part not in out:
            out.append(part)
    return out or [DEFAULT_GROUP]


__all__ = [
    "BatchInvoice",
    "DedupInfo",
    "GROUP_CHOICES",
    "DEFAULT_GROUP",
    "GroupStat",
    "ProductStat",
    "TOLERANCE",
    "Totals",
    "build_invoice",
    "dedupe",
    "dominant_rate",
    "group_stats",
    "money",
    "normalize_product_name",
    "parse_group_option",
    "product_stats",
    "round2",
    "sort_invoices",
    "to_float",
    "totals",
    "validate",
]
