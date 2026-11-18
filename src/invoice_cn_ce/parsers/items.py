# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""商品明细行抽取（PDF）—— 按词坐标重建版面，动态定位列。

上游 1.0.0 的 PDF 路径恒返回 ``items=[]``：票头字段用正则从清洗后的文本里取，
而明细是**表格**，正则拿不到「哪一列是哪一列」。本模块补上这块能力，且**不写
死任何票样的坐标**：

1. **版面重建**：``extract_words()`` 的词按 ``top`` 容差（:data:`LINE_TOL_Y`）聚成行，
   行内按 ``x0`` 升序 —— 全电发票的字符间距大，同一行的词上下会抖 2pt 左右。
2. **通用列定位**（三级，逐级降级）：

   a. **表头锚点**：在明细区上方找列标题（项目名称/规格型号/单位/数量/单价/
      金额/税率/征收率/税额），支持被切碎的标题（``'金'`` + ``'额'`` 两个字各占一个词）
      按相邻合并后再匹配；用锚点中心生成列边界。
   b. **数值右对齐反推**：用「合计」行金额/税额的右边界反推这两列；税率列再用
      明细行里 ``13%`` 这类 token 的 x 中心补。
   c. **兜底**：以上都不成立时，以「税率 token」为分界，左边最近的数字当金额、
      右边最近的数字当税额，整行原文存进 ``raw_line``，并标记低置信度 ——
      **不丢行、不算错**。

3. **单元格填充**：词按位置归列；跨列的长词（没被 pdfplumber 切开的相邻单元格）
   按空白切块后按字符比例估算位置。

输出沿用 :class:`invoice_cn_ce.schemas.InvoiceItem`：所有数值字段保持 ``str`` 原值
（``196.815`` 这类三位小数单价不丢精度），折扣行金额为负并带 ``is_reversal=True``。

设计约束：本模块只读页面视图（``PageView.words()``），不碰票头字段；票头仍由
:mod:`invoice_cn_ce.parsers.invoice` 负责。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Tuple

from invoice_cn_ce.schemas import InvoiceItem

# ---------------------------------------------------------------- 常量

#: 同一行的词 ``top`` 容差（pt）。全电发票同一行的词上下抖动约 2pt。
LINE_TOL_Y = 3.0
#: 表头锚点聚成「表头带」的 top 容差 —— 容许表头被拆成多行（竖排分散）。
HEADER_BAND_TOL_Y = 30.0
#: 拆分跨列长词时，两个词块之间视为「有空格」的最小间隙（pt）。
GAP_SPACE_PT = 3.0
#: 反推列的宽度超过这个值就认为区间不可信，改走「税率分界」兜底。
MAX_TRUSTED_COL_WIDTH = 150.0

COLUMN_ORDER = ("name", "spec", "unit", "qty", "price", "amount", "rate", "tax")
TEXT_KEYS = ("name", "spec", "unit")
NUM_KEYS = ("qty", "price", "amount", "tax")

#: 列标题候选（去空白后精确匹配）。长候选在前，保证合并匹配时不误吞。
LABEL_ALTS: dict = {
    "name": (
        "货物或应税劳务、服务名称",
        "货物或应税劳务名称",
        "项目名称及规格型号",
        "货物或应税劳务",
        "项目名称",
        "服务名称",
        "商品名称",
        "货物名称",
        "劳务名称",
    ),
    "spec": ("规格型号", "规格", "型号"),
    "unit": ("单位",),
    "qty": ("数量",),
    "price": ("单价",),
    "amount": ("金额",),
    "rate": ("税率/征收率", "税率征收率", "税率", "征收率"),
    "tax": ("税额",),
}
LABEL_INDEX: dict = {alt: key for key, alts in LABEL_ALTS.items() for alt in alts}
MAX_LABEL_LEN = max(len(a) for a in LABEL_INDEX)

#: 计量单位白名单（内容规则；单位列缺失时也能认出来）。
UNIT_WORDS = frozenset(
    """
    盒 瓶 个 台 件 支 套 张 把 只 箱 包 袋 双 卷 根 块 片 粒 枚 本 份 组 条 米 千克 克 吨
    kg g mg t ml l m m2 m3 cm mm km 升 毫升 千米 平方米 立方米 吨公里 天 次 人 月 年 小时
    项 批 辆 艘 架 座 间 层 门 扇 株 棵 亩 头 羽 对 副 打 令 筒 罐 桶 筐 提 捆 扎 页 册 台套
    """.split()
)

#: 税率 / 征收率 token。
RATE_RE = re.compile(r"^(?:\d{1,2}(?:\.\d+)?%|免税|不征税|零税率|\*+|不征收)$")
#: 数字（允许内部空格断字，如 ``¥213. 08``）。
NUM_RE = re.compile(r"-?\d[\d,]*(?:\.\s?\d+)?")
_PURE_NUM_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
_SPACE_RE = re.compile(r"[ \t\u3000\xa0]+")

#: 明细区里出现这些词的行不参与明细识别（票头 / 页脚残留）。
ROW_STOP_WORDS = (
    "合计",
    "价税合计",
    "小写",
    "备注",
    "开票人",
    "收款人",
    "复核",
    "销售方",
    "购买方",
    "开票日期",
    "发票号码",
)

#: 商品名称里的规格尾巴标记（仅用于把「规格」从品名里挪出来，不删信息）。
SPEC_TAIL_MARKERS = (
    "适用于",
    "适用型号",
    "无仪器",
    "不含仪器",
    "瓶装",
    "盒装",
    "袋装",
    "净含量",
    "规格型号",
)
_MODEL_TAIL_RE = re.compile(r"\s(?=[A-Za-z]{1,6}[-－]?\d)")
_CATEGORY_RE = re.compile(r"^\*([^*]{1,20})\*(.*)$", re.S)


# ---------------------------------------------------------------- 基础工具


def _norm(text: str) -> str:
    """去掉所有空白（含全角空格），便于标签匹配。"""
    return _SPACE_RE.sub("", text or "")


def _parse_num(text: str) -> Optional[float]:
    """把单元格文本解析为数字；失败返回 ``None``（不做 0 兜底）。"""
    if not text:
        return None
    cleaned = _norm(text)
    if not cleaned or not _PURE_NUM_RE.match(cleaned):
        return None
    try:
        return round(float(cleaned), 4)
    except ValueError:
        return None


def _is_rate_text(text: str) -> bool:
    return bool(RATE_RE.match(_norm(text)))


def _is_numeric_like(text: str) -> bool:
    """数字形态（允许货币符号、千分位、百分号、内部空格）。"""
    t = _norm(text).replace(",", "").replace("%", "")
    if not t or not any(ch.isdigit() for ch in t):
        return False
    return all(ch.isdigit() or ch in ".-" for ch in t)


def _clean_text(text: str) -> str:
    """合并空白、去掉汉字之间由排版造成的空格。"""
    if not text:
        return ""
    t = _SPACE_RE.sub(" ", text or "").strip()
    t = re.sub(r"(?<=[\u3000-\u9fff\uff00-\uffef])\s+(?=[\u3000-\u9fff\uff00-\uffef])", "", t)
    return t.strip()


def _fmt_num(value: Optional[float]) -> str:
    """数字 -> 稳定的十进制字符串（去掉浮点尾巴）。"""
    if value is None:
        return ""
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.4f}".rstrip("0").rstrip(".")


# ---------------------------------------------------------------- 版面结构


@dataclass
class Word:
    """一个词（pdfplumber ``extract_words()`` 的元素）。"""

    x0: float
    x1: float
    top: float
    text: str

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def norm(self) -> str:
        return _norm(self.text)

    def moved(self, x0: float, x1: float, text: str) -> "Word":
        return Word(x0=x0, x1=x1, top=self.top, text=text)


@dataclass
class Line:
    """版面行：同一 ``top`` 上的一组词。"""

    top: float
    page: int = 1
    words: List[Word] = field(default_factory=list)

    @property
    def flat(self) -> str:
        """去空白后的整行文本（标签/关键词匹配用）。"""
        return _norm("".join(w.text for w in self.words))

    @property
    def text(self) -> str:
        return join_cell(self.words)

    def sorted_words(self) -> List[Word]:
        return sorted(self.words, key=lambda w: w.x0)


def join_cell(words: Sequence[Word]) -> str:
    """按 x 顺序拼接词，间隙大于阈值时补一个空格。"""
    if not words:
        return ""
    ordered = sorted(words, key=lambda w: w.x0)
    out = [ordered[0].text]
    prev = ordered[0]
    for word in ordered[1:]:
        if word.x0 - prev.x1 > GAP_SPACE_PT:
            out.append(" ")
        out.append(word.text)
        prev = word
    return "".join(out)


def join_raw(words: Sequence[Word]) -> str:
    """原始行文本（按视觉顺序用空格分隔，便于人工核对）。"""
    return " ".join(w.text.strip() for w in sorted(words, key=lambda w: w.x0) if w.text.strip())


def to_words(raw_words: Iterable[dict]) -> List[Word]:
    """pdfplumber 词字典 -> :class:`Word`（跳过空词）。"""
    out: List[Word] = []
    for item in raw_words:
        text = (item.get("text") or "").strip()
        if not text:
            continue
        x0 = float(item.get("x0") or 0.0)
        x1 = float(item.get("x1") or x0)
        out.append(Word(x0=x0, x1=x1, top=float(item.get("top") or 0.0), text=text))
    return out


def build_lines(words: Iterable[Word], page: int = 1, tol: float = LINE_TOL_Y) -> List[Line]:
    """按 ``top`` 聚类成行，行内按 ``x0`` 升序。"""
    lines: List[Line] = []
    for word in sorted(words, key=lambda w: (w.top, w.x0)):
        if lines and abs(word.top - lines[-1].top) <= tol:
            lines[-1].words.append(word)
        else:
            lines.append(Line(top=word.top, page=page, words=[word]))
    for line in lines:
        line.words.sort(key=lambda w: w.x0)
    return lines


# ---------------------------------------------------------------- 列模型


@dataclass
class Column:
    """一列。``start``/``end`` 是横向边界，``right`` 是数值列的期望右边界。"""

    key: str
    start: float = 0.0
    end: float = 0.0
    center: float = 0.0
    right: float = 0.0
    source: str = "header"  # header / reverse / content / fallback

    @property
    def is_numeric(self) -> bool:
        return self.key in NUM_KEYS


@dataclass
class HeaderAnchor:
    """一个表头锚点。"""

    key: str
    x0: float
    x1: float
    top: float

    @property
    def center(self) -> float:
        return (self.x0 + self.x1) / 2.0


def _rep(col: Column) -> float:
    """列的代表 x：数值列按右边界对齐，文本列按中心对齐。"""
    return col.right if col.is_numeric else col.center


def relink(cols: List[Column], page_width: float) -> List[Column]:
    """按代表 x 排序后，用相邻列代表点的中点重新划分边界。"""
    if not cols:
        return cols
    cols.sort(key=_rep)
    for i, col in enumerate(cols):
        col.start = 0.0 if i == 0 else (_rep(cols[i - 1]) + _rep(col)) / 2.0
        col.end = page_width if i == len(cols) - 1 else (_rep(col) + _rep(cols[i + 1])) / 2.0
    return cols


def detect_header_anchors(lines: Sequence[Line], y_max: float) -> dict:
    """在 ``top < y_max`` 的行里找表头锚点（支持逐字切碎后合并匹配）。"""
    cands: List[HeaderAnchor] = []
    for line in lines:
        if line.top >= y_max:
            continue
        words = line.sorted_words()
        for i, word in enumerate(words):
            acc = ""
            for j in range(i, min(len(words), i + 7)):
                acc += words[j].norm
                if not acc:
                    break
                key = LABEL_INDEX.get(acc)
                if key is not None:
                    cands.append(HeaderAnchor(key=key, x0=word.x0, x1=words[j].x1, top=line.top))
                    break
                if len(acc) > MAX_LABEL_LEN:
                    break
    if not cands:
        return {}
    return _best_header_band(cands)


def _best_header_band(cands: Sequence[HeaderAnchor]) -> dict:
    """把锚点按 top 聚成若干「表头带」，取列数最多的那一带。"""
    bands: List[List[HeaderAnchor]] = []
    for cand in sorted(cands, key=lambda c: c.top):
        if bands and abs(cand.top - bands[-1][0].top) <= HEADER_BAND_TOL_Y:
            bands[-1].append(cand)
        else:
            bands.append([cand])

    def score(band: List[HeaderAnchor]) -> tuple:
        return (len({c.key for c in band}), -band[0].top)

    best = max(bands, key=score)
    anchors: dict = {}
    for cand in best:
        cur = anchors.get(cand.key)
        if cur is None or (cand.top, cand.x0) < (cur.top, cur.x0):
            anchors[cand.key] = cand
    return anchors


def build_columns(anchors: dict, page_width: float) -> List[Column]:
    """由表头锚点生成列（代表 x = 锚点中心 / 数值列取右边界）。"""
    if not anchors:
        return []
    cols = [
        Column(key=key, center=anchor.center, right=anchor.x1, source="header")
        for key, anchor in anchors.items()
    ]
    return relink(cols, page_width)


def column_keys(cols: Sequence[Column]) -> set:
    return {c.key for c in cols}


# ---------------------------------------------------------------- 合计行


@dataclass
class TotalsPick:
    """「合计」行的解析结果。"""

    amount: Optional[float] = None
    tax: Optional[float] = None
    amount_right: Optional[float] = None
    tax_right: Optional[float] = None
    line: Optional[Line] = None

    @property
    def found(self) -> bool:
        return self.line is not None


def find_totals_line(lines: Sequence[Line]) -> Optional[Line]:
    """找「合计」行：同一行同时含 ``合`` 与 ``计``（允许两字分开）。"""
    for line in lines:
        flat = line.flat
        if "合计" in flat and "价税合计" not in flat:
            return line
    for line in lines:
        flat = line.flat
        if "合" in flat and "计" in flat:
            return line
    return None


def parse_totals_line(line: Optional[Line]) -> TotalsPick:
    """解析合计行，返回金额/税额及其右边界（用于反推列）。"""
    pick = TotalsPick(line=line)
    if line is None:
        return pick
    words = line.sorted_words()
    start_idx = 0
    for i, word in enumerate(words):
        if "合计" in word.norm or ("合" in word.norm and "计" in word.norm):
            start_idx = i + 1
            break
        if (
            word.norm.startswith("合")
            and i + 1 < len(words)
            and words[i + 1].norm.startswith("计")
        ):
            start_idx = i + 2
            break

    located: List[Tuple[float, float]] = []
    for word in words[start_idx:]:
        text = word.text
        length = max(len(text), 1)
        for m in NUM_RE.finditer(text):
            value = _parse_num(m.group())
            if value is None:
                continue
            located.append((word.x0 + word.width * (m.end() / length), value))
    if len(located) >= 2:
        pick.amount, pick.amount_right = located[-2][1], located[-2][0]
        pick.tax, pick.tax_right = located[-1][1], located[-1][0]
    elif len(located) == 1:
        pick.amount, pick.amount_right = located[0][1], located[0][0]
    return pick


def complete_columns_from_totals(
    cols: List[Column], pick: TotalsPick, page_width: float
) -> List[Column]:
    """策略 b：用合计行的右边界反推「金额」「税额」列。"""
    if not pick.found:
        return cols
    keys = column_keys(cols)
    new_cols = list(cols)
    if "tax" not in keys and pick.tax_right is not None:
        tax_right = min(pick.tax_right, page_width)
        new_cols.append(Column("tax", right=tax_right, center=tax_right, source="reverse"))
        keys.add("tax")
    if "amount" not in keys and pick.amount_right is not None:
        amount_right = min(pick.amount_right, page_width)
        new_cols.append(Column("amount", right=amount_right, center=amount_right, source="reverse"))
    return relink(new_cols, page_width)


def add_content_columns(
    cols: List[Column], detail_words: Sequence[Word], page_width: float
) -> List[Column]:
    """按内容补列：明细行里税率 token 的 x 中心 -> 税率列。"""
    if "rate" in column_keys(cols):
        return cols
    rate_words = [w for w in detail_words if _is_rate_text(w.text) and "%" in w.text]
    if not rate_words:
        return cols
    centers = sorted(w.cx for w in rate_words)
    center = centers[len(centers) // 2]
    new_cols = list(cols)
    new_cols.append(Column("rate", center=center, right=center + 10.0, source="content"))
    return relink(new_cols, page_width)


# ---------------------------------------------------------------- 单元格填充


def split_word(word: Word) -> List[Word]:
    """把跨列的长词按空白切块，按字符比例估算每块的位置。"""
    text = word.text
    if not text:
        return [word]
    chunks = [(m.start(), m.end(), m.group()) for m in re.finditer(r"\S+", text)]
    if len(chunks) <= 1:
        return [word]
    per = word.width / max(len(text), 1)
    return [
        word.moved(word.x0 + per * start, word.x0 + per * end, chunk)
        for start, end, chunk in chunks
    ]


def _spans_columns(word: Word, cols: Sequence[Column]) -> int:
    return sum(1 for col in cols if word.x0 < col.end and word.x1 > col.start)


def _interval_column(word: Word, cols: Sequence[Column]) -> Optional[Column]:
    for col in cols:
        if col.start <= word.x0 < col.end:
            return col
    if cols and word.x0 >= cols[-1].end:
        return cols[-1]
    return cols[0] if cols else None


def assign_word(word: Word, cols: Sequence[Column]) -> Optional[str]:
    """把一个词归到某一列。"""
    if not cols:
        return None
    norm = word.norm
    if not norm:
        return None
    keys = column_keys(cols)

    # 内容规则优先：税率 / 单位（文字形态特殊，最不容易误判）
    if "rate" in keys and _is_rate_text(norm):
        return "rate"
    if "unit" in keys and norm in UNIT_WORDS and len(norm) <= 4:
        return "unit"

    numeric_cols = [c for c in cols if c.is_numeric]
    if numeric_cols and _is_numeric_like(norm):
        best = min(numeric_cols, key=lambda c: abs(word.x1 - c.right))
        if abs(word.x1 - best.right) <= 80.0:
            return best.key

    col = _interval_column(word, cols)
    if col is None:
        return None
    if col.is_numeric and numeric_cols and not _is_numeric_like(norm):
        text_cols = [c for c in cols if c.key in TEXT_KEYS]
        if text_cols:
            return min(text_cols, key=lambda c: abs(word.x0 - c.start)).key
    return col.key


def assign_line(line: Line, cols: Sequence[Column]) -> dict:
    """把一行填进各列；跨列长词先切块。"""
    cells: dict = {key: [] for key in COLUMN_ORDER}
    for word in line.sorted_words():
        pieces = split_word(word) if _spans_columns(word, cols) >= 2 else [word]
        for piece in pieces:
            key = assign_word(piece, cols)
            if key:
                cells.setdefault(key, []).append(piece)
    for key in cells:
        cells[key].sort(key=lambda w: w.x0)
    return cells


def cell_text(cells: dict, key: str) -> str:
    return join_cell(cells.get(key) or [])


def cell_raw(cells: dict, key: str) -> str:
    """数值单元格的原始文本（去掉货币符号与空格，保留原小数位）。"""
    words = cells.get(key) or []
    text = "".join(w.text for w in words)
    return _norm(text).replace("¥", "").replace("￥", "").replace(",", "")


def cell_num(cells: dict, key: str) -> Optional[float]:
    raw = cell_raw(cells, key)
    return _parse_num(raw) if raw else None


# ---------------------------------------------------------------- 明细行


@dataclass
class _Row:
    """明细行的中间态（续行会往里追加内容）。"""

    name_parts: List[str] = field(default_factory=list)
    spec_parts: List[str] = field(default_factory=list)
    unit: str = ""
    qty: str = ""
    price: str = ""
    amount: str = ""
    rate: str = ""
    tax: str = ""
    raws: List[str] = field(default_factory=list)
    page: int = 1
    low_confidence: bool = False


def _split_category(name: str) -> Tuple[str, str]:
    """拆出 ``*类别*`` 前缀，返回 ``(category, item_name_without_prefix)``。"""
    m = _CATEGORY_RE.match(name or "")
    if not m:
        return "", (name or "").strip()
    return m.group(1).strip(), m.group(2).strip()


def _split_spec_tail(name: str) -> Tuple[str, str]:
    """把品名里夹带的规格描述拆出来（``适用于…`` / `` GA-3型 300支…``）。"""
    if not name:
        return "", ""
    cut = -1
    for marker in SPEC_TAIL_MARKERS:
        idx = name.find(marker)
        if idx >= 4 and (cut < 0 or idx < cut):
            cut = idx
    m = _MODEL_TAIL_RE.search(name)
    if m and m.start() >= 4 and (cut < 0 or m.start() < cut):
        cut = m.start()
    if cut < 0:
        return name, ""
    return name[:cut].strip(" -,，、"), name[cut:].strip()


def _join_parts(parts: Sequence[str]) -> str:
    return _clean_text(" ".join(p for p in parts if p))


def _finish_row(row: _Row, item_no: int, spec_column: bool = True) -> Optional[InvoiceItem]:
    """中间态 -> :class:`InvoiceItem`。

    ``spec_column`` 为真表示这一版票面**确实有**「规格型号」列，此时品名里
    夹带的文字就属于品名本身（可能只是单元格换行），不做拆分 —— 忠于票面。
    只有版式没有规格列时，才把 ``适用于…`` / `` GA-3型 …`` 这类尾巴挪到
    ``spec``，避免品名被规格噪声污染。
    """
    raw_name = _join_parts(row.name_parts)
    spec = _join_parts(row.spec_parts)
    category, body = _split_category(raw_name)
    if not spec and not spec_column:
        body, tail = _split_spec_tail(body)
        spec = tail
    name = body
    if category:
        # 架构约定：name 保留 ``*类别*`` 前缀，category 是拆出来的那一段。
        name = f"*{category}*{body}" if body else f"*{category}*"

    amount = _parse_num(row.amount)
    tax = _parse_num(row.tax)
    row_type = "商品"
    if amount is not None and amount < 0:
        row_type = "折扣"
    elif amount is None:
        row_type = "其他"

    total = ""
    if amount is not None and tax is not None:
        total = f"{amount + tax:.2f}"

    return InvoiceItem(
        item_no=item_no,
        name=name or "",
        category=category,
        spec=spec,
        unit=row.unit,
        quantity=row.qty,
        unit_price=row.price,
        amount=row.amount,
        tax_rate=_normalize_rate(row.rate),
        tax_amount=row.tax,
        total_amount=total,
        is_reversal=(amount is not None and amount < 0),
        row_type=row_type,
        raw_line=" ⏎ ".join(row.raws)[:400],
    )


def _normalize_rate(text: str) -> str:
    t = _norm(text)
    if not t:
        return ""
    m = re.match(r"^(\d{1,2}(?:\.\d+)?)%$", t)
    if m:
        return f"{m.group(1)}%"
    if t in ("免税", "不征税", "零税率", "不征收"):
        return t
    if set(t) == {"*"}:
        return "*"
    return t[:10]


def _boundary_split(line: Line) -> Optional[Tuple[str, str, str, str]]:
    """以「税率 token」为分界：左边最近数字=金额，右边最近数字=税额。

    返回 ``(amount, rate, tax, name_text)``。
    """
    words = line.sorted_words()
    rate_idx = next((i for i, w in enumerate(words) if _is_rate_text(w.text)), None)
    if rate_idx is None:
        return None
    amount = ""
    for word in reversed(words[:rate_idx]):
        for m in reversed(list(NUM_RE.finditer(word.text))):
            if _parse_num(m.group()) is not None:
                amount = _norm(m.group())
                break
        if amount:
            break
    if not amount:
        return None
    tax = ""
    for word in words[rate_idx + 1 :]:
        hit = next((m.group() for m in NUM_RE.finditer(word.text)), None)
        if hit is not None and _parse_num(hit) is not None:
            tax = _norm(hit)
            break
    name_parts: List[str] = []
    for word in words[:rate_idx]:
        if _is_numeric_like(word.text) or word.norm in UNIT_WORDS:
            continue
        name_parts.append(word.text.strip())
    return amount, words[rate_idx].text, tax, _clean_text(" ".join(p for p in name_parts if p))


def extract_page_items(
    page_words: Sequence[dict],
    page_width: float,
    page_no: int = 1,
    *,
    allow_fallback: bool = True,
) -> Tuple[List[InvoiceItem], str]:
    """抽取一页的明细行，返回 ``(items, note)``。

    ``note`` 为 ``""`` 表示正常；``"low_confidence"`` 表示走了兜底通路（金额/税额
    可靠但列定位靠推断）；``"no_table"`` 表示这一页没有明细表。
    """
    words = to_words(page_words)
    if not words:
        return [], "no_table"
    lines = build_lines(words, page_no)

    totals_line = find_totals_line(lines)
    hj_y = totals_line.top if totals_line is not None else 900.0
    anchors = detect_header_anchors(lines, hj_y - 3.0)
    cols = build_columns(anchors, page_width)

    pick = parse_totals_line(totals_line)
    if pick.found:
        cols = complete_columns_from_totals(cols, pick, page_width)

    header_bottom = max((a.top for a in anchors.values()), default=-1.0) + 4.0 if anchors else -1.0
    detail_words = [
        w
        for line in lines
        if line.top > header_bottom and (totals_line is None or line.top < hj_y - 0.5)
        for w in line.words
    ]
    if cols:
        cols = add_content_columns(cols, detail_words, page_width)

    amount_col = next((c for c in cols if c.key == "amount"), None)
    trust_amount = amount_col is not None and (
        amount_col.source == "header"
        or (amount_col.end - amount_col.start) <= MAX_TRUSTED_COL_WIDTH
    )
    if not cols or not trust_amount:
        if not allow_fallback:
            return [], "no_table"
        if totals_line is None:
            return [], "no_table"
    note = "" if trust_amount else "low_confidence"

    rows: List[_Row] = []
    cur: Optional[_Row] = None
    for line in lines:
        if line.top <= header_bottom:
            continue
        if totals_line is not None and line.top >= hj_y - 0.5:
            continue
        if not line.flat:
            continue
        cells = assign_line(line, cols)
        amount_text = cell_raw(cells, "amount") if trust_amount else ""
        amount = _parse_num(amount_text) if amount_text else None
        rate_text = cell_text(cells, "rate")

        if amount is None and (not trust_amount or rate_text):
            fallback = _boundary_split(line)
            if fallback is not None:
                fb_amount, fb_rate, fb_tax, fb_name = fallback
                cur = _Row(page=page_no, low_confidence=not trust_amount)
                cur.amount = fb_amount
                cur.tax = fb_tax
                cur.rate = fb_rate
                cur.name_parts.append(cell_text(cells, "name") or fb_name)
                cur.spec_parts.append(cell_text(cells, "spec"))
                cur.raws.append(join_raw(line.words))
                rows.append(cur)
                continue

        if amount is not None:
            cur = _Row(page=page_no)
            cur.amount = amount_text
            cur.tax = cell_raw(cells, "tax")
            cur.rate = rate_text
            cur.unit = cell_text(cells, "unit")
            cur.qty = cell_raw(cells, "qty")
            cur.price = cell_raw(cells, "price")
            cur.name_parts.append(cell_text(cells, "name"))
            cur.spec_parts.append(cell_text(cells, "spec"))
            cur.raws.append(join_raw(line.words))
            rows.append(cur)
            continue

        # 续行：品名 / 规格续到上一条，缺的数值补上
        if cur is None:
            continue
        if any(word in line.flat for word in ROW_STOP_WORDS) and not cell_text(cells, "name"):
            continue
        cur.name_parts.append(cell_text(cells, "name"))
        cur.spec_parts.append(cell_text(cells, "spec"))
        if not cur.unit:
            cur.unit = cell_text(cells, "unit")
        if not cur.qty:
            cur.qty = cell_raw(cells, "qty")
        if not cur.price:
            cur.price = cell_raw(cells, "price")
        if not cur.rate:
            cur.rate = rate_text
        if not cur.tax:
            cur.tax = cell_raw(cells, "tax")
        cur.raws.append(join_raw(line.words))

    spec_column = "spec" in column_keys(cols)
    items: List[InvoiceItem] = []
    for row in rows:
        item = _finish_row(row, len(items) + 1, spec_column)
        if item is None:
            continue
        if row.low_confidence:
            note = "low_confidence"
        items.append(item)
    return items, (note if items else "no_table")


def extract_items(
    pages: Sequence[object], *, allow_fallback: bool = True
) -> Tuple[List[InvoiceItem], str]:
    """跨页抽取一张票的全部明细行。

    ``pages`` 是 :class:`invoice_cn_ce.pipeline.PageView` 列表（只需 ``.words()``
    与 ``._page.width``）；返回 ``(items, note)``，``note`` 同
    :func:`extract_page_items`。
    """
    all_items: List[InvoiceItem] = []
    notes: List[str] = []
    for idx, page in enumerate(pages, start=1):
        try:
            raw_words = page.words()
        except Exception:  # pdfplumber 在坏页面上可能抛任意异常，明细可缺席
            notes.append("no_table")
            continue
        width = float(getattr(getattr(page, "_page", None), "width", 0.0) or 600.0)
        items, note = extract_page_items(
            raw_words, width, page_no=idx, allow_fallback=allow_fallback
        )
        notes.append(note)
        all_items.extend(items)

    for i, item in enumerate(all_items, start=1):
        item.item_no = i
    if all_items:
        note = "low_confidence" if "low_confidence" in notes else ""
    else:
        note = "no_table"
    return all_items, note


__all__ = [
    "COLUMN_ORDER",
    "LINE_TOL_Y",
    "Column",
    "HeaderAnchor",
    "Line",
    "TotalsPick",
    "Word",
    "add_content_columns",
    "assign_line",
    "build_columns",
    "build_lines",
    "complete_columns_from_totals",
    "detect_header_anchors",
    "extract_items",
    "extract_page_items",
    "find_totals_line",
    "join_cell",
    "parse_totals_line",
    "to_words",
]
