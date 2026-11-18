# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""``python -m invoice_cn_ce.batch`` —— 批量导入与汇总报表。

    python -m invoice_cn_ce.batch <路径...> [-o out] [--group 销售方,月份] [--no-dedup]

路径可以是 PDF / XML / OFD 文件、目录（递归）或 ZIP（含嵌套 ZIP），可混用。
解析复用 :func:`invoice_cn_ce.formats.parse_pdf`（票头字段用上游解析器），
本模块只负责「批量输入 → 校验 → 去重 → 汇总 → 报表」。

上游的 ``invoice_cn_ce`` 控制脚本与 HTTP 服务不受影响：这里是独立子命令，
不改任何既有入口的默认行为。
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
import traceback
from datetime import datetime
from typing import List, Optional, Sequence

from invoice_cn_ce.batch import aggregate, ingest, report
from invoice_cn_ce.batch.aggregate import TOLERANCE, BatchInvoice, DedupInfo, money
from invoice_cn_ce.formats import parse_pdf

APP_NAME = "invoice_cn_ce 批量归集"
UNREACHABLE = "—"

# ---------------------------------------------------------------- 终端着色

_COLOR = {"on": False}


def _enable_color() -> bool:
    """尽量打开 Windows 控制台的 ANSI 支持；失败就退化纯文本。"""
    if os.environ.get("NO_COLOR") or not sys.stdout.isatty():
        return False
    if os.name == "nt":
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
            handle = kernel32.GetStdHandle(-11)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                kernel32.SetConsoleMode(handle, mode.value | 0x0004)
                return True
        except Exception:  # noqa: BLE001
            return False
    return True


def _paint(text: str, code: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _COLOR["on"] else text


def _red(text: str) -> str:
    return _paint(text, "31")


def _green(text: str) -> str:
    return _paint(text, "32")


def _yellow(text: str) -> str:
    return _paint(text, "33")


def _bold(text: str) -> str:
    return _paint(text, "1")


def _set_console_cp(codepage: int) -> None:
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleCP(codepage)  # type: ignore[attr-defined]
        ctypes.windll.kernel32.SetConsoleOutputCP(codepage)  # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        pass


def _setup_console() -> None:
    """中文输出兜底。

    * 控制台本来就支持中文（cp936 / GBK / UTF-8）-> 保持原编码，只兜底 errors。
    * 编码表示不了中文（cp437 / cp1252 等）-> 切 UTF-8 并把控制台代码页设为 65001。
    """
    cjk_safe = {"utf-8", "utf8", "cp936", "gbk", "gb2312", "gb18030", "cp950", "big5"}
    encoding = (getattr(sys.stdout, "encoding", "") or "").lower()
    if encoding in cjk_safe:
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
            except (AttributeError, ValueError, OSError):
                pass
        return
    if os.name == "nt":
        _set_console_cp(65001)
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
        except (AttributeError, ValueError, OSError):
            pass


# ---------------------------------------------------------------- 参数


def _package_version() -> str:
    try:
        from importlib.metadata import version

        return version("invoice-cn-ce")
    except Exception:  # noqa: BLE001 - 未安装（PYTHONPATH 开发态）时退回常量
        return "1.0.0"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m invoice_cn_ce.batch",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
        description=f"{APP_NAME}：把一堆发票（PDF/XML/OFD）归集成一张可追溯的汇总表。",
        epilog="""\
示例：
  python -m invoice_cn_ce.batch invoices
  python -m invoice_cn_ce.batch 发票.pdf -o D:\\报表
  python -m invoice_cn_ce.batch 发票包.zip "补充\\*.pdf" --group 销售方,月份
  python -m invoice_cn_ce.batch . -o out --group 品目,税率
  python -m invoice_cn_ce.batch invoices --no-dedup      # 同号票也要各算一张

说明：
  * 路径可以是单个文件、目录（递归扫描）或 ZIP（支持嵌套 ZIP），可混用；
  * **支持通配符**：`*` / `?` / `[..]`，`**` 可递归（Windows cmd 自己不展开，
    由本工具展开；含空格时用引号包住）；一个都没匹配上仍报「路径不存在」；
  * 支持 .pdf / .xml / .ofd；票头字段由上游 invoice_cn_ce 解析器提取；
  * PDF 的商品明细按词坐标重建版面（表头锚点 + 数值右对齐 + 税率反推），
    不依赖任何票样写死的坐标；
  * 自动跳过 __MACOSX、._*、~$*、隐藏文件等垃圾条目；
  * 同一文件重复出现按内容 hash 去重，同一张票重复出现按发票号码去重；
  * 扫描件（无文本层）、非发票文件进「异常清单」，不会静默算 0；
  * 「品目」分组是按**整票主品目**归属（一张票计入其净金额最大的品目），
    与 xlsx 里逐行口径的「品目汇总」sheet 不是一回事，两者页脚都有注明；
  * 把文件 / 文件夹 / ZIP 直接拖到仓库根目录的 drop.bat 上也能跑。

退出码：
  0  成功（至少 1 张有效发票）
  2  参数或路径错误（没有可解析的文件、--group 维度非法、无参数）
  3  跑完了但没有一张有效发票（全部解析失败 / 全是扫描件…），报表仍会输出

输出：
  <输出目录>/发票汇总-YYYYMMDD-HHMMSS.xlsx   五个 sheet
  <输出目录>/发票汇总-YYYYMMDD-HHMMSS.csv    发票清单主表（utf-8-sig）
  <输出目录>/数据.json                       全量结构化数据
""",
    )
    parser.add_argument("-h", "--help", action="help", help="显示本帮助并退出")
    parser.add_argument("paths", nargs="*", metavar="路径",
                        help="PDF / XML / OFD 文件、目录或 ZIP，可写多个")
    parser.add_argument("-o", "--out", default="./out", metavar="目录",
                        help="输出目录（默认 ./out）")
    parser.add_argument("--group", default=aggregate.DEFAULT_GROUP, metavar="维度",
                        help=("分组维度，默认「销售方」；可选："
                              + "、".join(aggregate.GROUP_CHOICES)
                              + "，多个用逗号分隔，如「购买方,月份」"
                              + "；其中「品目」按整票主品目归属（逐行口径见 xlsx 的「品目汇总」）"))
    parser.add_argument("--no-dedup", action="store_true",
                        help="关闭「按发票号码去重」（内容相同的文件仍只解析一次）")
    parser.add_argument("--keep-temp", action="store_true",
                        help="保留 ZIP 解包临时目录（排查问题时用）")
    parser.add_argument("--quiet", action="store_true", help="只输出最终结果，不打印逐文件进度")
    parser.add_argument("--no-color", action="store_true", help="禁用彩色输出")
    parser.add_argument("-V", "--version", action="version",
                        version=f"{APP_NAME} {_package_version()}")
    return parser


# ---------------------------------------------------------------- 主流程


def run(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    _COLOR["on"] = (not args.no_color) and _enable_color()

    if not args.paths:
        parser.print_help()
        return 2

    try:
        group_dims = aggregate.parse_group_option(args.group)
    except ValueError as exc:
        print(_red(f"参数错误：{exc}"), file=sys.stderr)
        return 2

    started = datetime.now()
    clock = time.time()
    issues: List[dict] = []
    invoices: List[BatchInvoice] = []

    ingestor = ingest.Ingestor()
    try:
        result = ingestor.collect(args.paths)
        for path, reason in result.skipped:
            issues.append({"file": path, "level": "文件级", "invoice_number": "",
                           "reason": reason, "raw": ""})
        if not result.files:
            print(_red("没有找到任何可解析的文件（.pdf / .xml / .ofd）。"), file=sys.stderr)
            for path, reason in result.skipped[:10]:
                print(f"  - {path}（{reason}）", file=sys.stderr)
            return 2

        for idx, source in enumerate(result.files, start=1):
            if not args.quiet:
                print(f"[{idx}/{len(result.files)}] {os.path.basename(source.display)} ... ",
                      end="", flush=True)
            try:
                with open(source.path, "rb") as fp:
                    contents = fp.read()
                resp = parse_pdf(contents, filename=os.path.basename(source.path))

                # 解析没成功（无文本层 / 文件损坏 / 非发票）-> 只进异常清单，
                # 绝不能变成一张金额为 0 的「发票」并进合计与张数。
                failure = _failure_reason(resp)
                if failure:
                    issues.append({
                        "file": source.display, "level": "文件级", "invoice_number": "",
                        "reason": failure, "raw": "",
                    })
                    if not args.quiet:
                        print(_red("跳过"))
                    continue

                inv = aggregate.build_invoice(resp, source)
                if not inv.invoice_no and not inv.items and inv.amount is None:
                    issues.append({
                        "file": source.display, "level": "文件级", "invoice_number": "",
                        "reason": "未识别到发票内容（缺发票号码、金额与明细）", "raw": "",
                    })
                    if not args.quiet:
                        print(_red("跳过"))
                    continue

                invoices.append(inv)
                if not args.quiet:
                    note = f"{len(inv.items)} 行明细"
                    if inv.has_issue:
                        note += _yellow(f"（{len(inv.issues)} 处待核对）")
                    print(note)
            except Exception as exc:  # noqa: BLE001 - 单个文件失败不能中断整批
                issues.append({
                    "file": source.display, "level": "文件级", "invoice_number": "",
                    "reason": f"解析异常：{type(exc).__name__}: {exc}", "raw": "",
                })
                if not args.quiet:
                    print(_red("异常"))
    finally:
        if not args.keep_temp:
            ingestor.result.cleanup()

    kept, dedup = aggregate.dedupe(invoices, result.hash_duplicates, enabled=not args.no_dedup)
    kept = aggregate.sort_invoices(kept)
    grand = aggregate.totals(kept)
    groups = {dim: aggregate.group_stats(kept, dim) for dim in group_dims}
    products = aggregate.product_stats(kept)

    # 票级问题进异常清单（文件级已在解析阶段收集）
    seen = {(i["file"], i["invoice_number"], i["reason"]) for i in issues}
    for inv in kept:
        if not inv.has_issue:
            continue
        reason = "；".join(inv.issues)
        key = (inv.file_label, inv.invoice_no, reason)
        if key in seen:
            continue
        seen.add(key)
        raw = inv.items[0].raw_line if inv.items else ""
        issues.append({"file": inv.file_label, "level": "票级",
                       "invoice_number": inv.invoice_no, "reason": reason, "raw": raw})

    elapsed = time.time() - clock
    stamp = started.strftime("%Y%m%d-%H%M%S")
    meta = {
        "tool": "invoice_cn_ce.batch",
        "version": _package_version(),
        "run_at": started.strftime("%Y-%m-%d %H:%M:%S"),
        "elapsed": round(elapsed, 2),
        "inputs": result.inputs,
        "file_count": len(result.files),
        "group_dims": group_dims,
        "tolerance": TOLERANCE,
        "dedup_enabled": not args.no_dedup,
        "invoice_count": len(kept),
        "issue_count": len(issues),
        "skipped": [{"path": p, "reason": r} for p, r in result.skipped],
    }
    paths = report.write_outputs(args.out, kept, issues, dedup, grand, groups, products,
                                meta, stamp)

    _print_summary(kept, issues, dedup, grand, result, elapsed, group_dims)
    print(_bold("输出文件："))
    for key in ("xlsx", "csv", "json"):
        print(f"  {paths[key]}")
    print(_bold(f"\n不含税合计 {grand.amount:,.2f} | 税额 {grand.tax:,.2f} | "
                f"价税合计 {grand.total:,.2f} | "
                f"勾稽{'通过' if grand.reconciled else '不通过'}"))

    if not kept:
        # 「跑完了但一张有效票都没有」独立成一个退出码，脚本里才好和「正常跑完」
        # 区分（类比 grep：没匹配到返回非 0）。报表仍然写出来了，异常清单可查原因。
        print(_red("\n没有一张有效发票：见上方「异常清单」与输出目录里的 xlsx。"),
              file=sys.stderr)
        return 3
    return 0


def _short(text: str, limit: int = 88) -> str:
    if len(text) <= limit:
        return text
    head = limit // 2 - 2
    return f"{text[:head]}...{text[- (limit - head - 3):]}"


#: ``ExtractResponse.message`` 里「自带完整含义」的英文 token -> 中文原因。
_FAILURE_TEXT = {
    "not_text": "PDF 无文本层（扫描件 / 纯图片），需先 OCR 或索取电子原件",
    "empty_file": "空文件（0 字节），没有可解析内容",
}

#: ``<类别>: <细节>`` 形态的 token -> 中文类别名。
_FAILURE_HEAD = {
    "parse_failed": "解析失败",
    "xml_parse_failed": "XML 解析失败",
    "ofd_parse_failed": "OFD 解析失败",
    "file_too_large": "文件超过大小上限",
    "unsupported_format": "不支持的文件格式",
}

#: pdfminer 往 stderr 打的告警行（截断 PDF 上常见），不该当「失败原因」展示。
_PDF_NOISE_RE = re.compile(r"^(?:WARNING|IGNORING|DEBUG|INFO)[:\s]", re.IGNORECASE)


def _strip_pdf_noise(detail: str, limit: int = 160) -> str:
    """压掉 pdfminer 的告警噪声，只留真正的失败原因（真失败不会被吞掉）。"""
    if not detail:
        return ""
    lines = [line.strip() for line in str(detail).splitlines()]
    kept = [line for line in lines if line and not _PDF_NOISE_RE.match(line)]
    text = " ".join(kept).strip()
    if len(text) > limit:
        return text[: limit - 1] + "…"
    return text


def _failure_reason(resp) -> str:
    """``ExtractResponse.message`` -> 中文失败原因（空串表示解析成功）。

    上游 message 的契约是 ``"success"`` / ``"not_text"`` / ``"error: <token>[: <detail>]"``。
    这里把 token 翻成中文，并把 pdfminer 的告警噪声压掉；**未知 token 原样附在末尾**，
    免得把一个没见过的失败类型翻译成含糊的「解析失败」而无从排查。
    """
    message = (getattr(resp, "message", "") or "").strip()
    if not message or message == "success":
        return ""
    text = message[len("error:"):].strip() if message.startswith("error:") else message
    token, _, detail = text.partition(":")
    token = token.strip()
    detail = _strip_pdf_noise(detail.strip())

    if token in _FAILURE_TEXT:
        return _FAILURE_TEXT[token]
    head = _FAILURE_HEAD.get(token)
    if head:
        return f"{head}：{detail}" if detail else head
    return f"解析未成功：{detail}" if detail else f"解析未成功（{text}）"


def _failure_reason(resp) -> str:
    """``ExtractResponse.message`` -> 人类可读的失败原因（空串表示解析成功）。

    上游的 message 契约是 ``"success"`` / ``"not_text"`` / ``"error: xxx"``；
    这里把三种情形翻译成中文原因，供「异常清单」直接展示。
    """
    message = getattr(resp, "message", "") or ""
    if not message or message == "success":
        return ""
    if message == "not_text":
        return "PDF 无文本层（扫描件 / 纯图片），需先 OCR 或索取电子原件"
    if message.startswith("error:"):
        detail = message[len("error:"):].strip()
        if detail.startswith("parse_failed:"):
            detail = detail[len("parse_failed:"):].strip()
        return f"解析失败：{detail or message}"
    return f"解析未成功：{message}"


def _print_summary(invoices: Sequence[BatchInvoice], issues: Sequence[dict],
                   dedup: DedupInfo, grand, result, elapsed: float,
                   group_dims: Optional[Sequence[str]] = None) -> None:
    line = "=" * 78
    print("\n" + line)
    print(_bold(f"{APP_NAME} 结果"))
    print(line)
    print(f"输入路径     : {len(result.inputs)} 个")
    print(f"待解析文件   : {len(result.files)} 个（忽略 {len(result.skipped)} 个不支持的文件）")
    print(f"发票张数     : {len(invoices)} 张")
    if group_dims:
        print(f"分组维度     : {'、'.join(group_dims)}")
    if dedup.duplicate_count:
        print(_yellow(f"重复         : {dedup.duplicate_count} 处已忽略"
                      f"（内容重复 {len(dedup.by_hash)}，同号码 {len(dedup.by_number)}）"))
        for dup, keep in list(dedup.by_number)[:5]:
            print(f"               - 号码重复：{_short(dup)}")
            print(f"                 已保留：{_short(keep)}")
    print(f"不含税金额   : {grand.amount:,.2f}")
    print(f"税额         : {grand.tax:,.2f}")
    print(f"价税合计     : {grand.total:,.2f}")
    if grand.reconciled:
        print("勾稽总校验   : " + _green("通过（总额 = 各票之和）"))
    else:
        print("勾稽总校验   : " + _red("不通过：" + "；".join(grand.notes)))
    bad = [inv for inv in invoices if inv.has_issue]
    detail_total = sum(len(inv.items) for inv in invoices)
    empty = [inv for inv in invoices if not inv.items]
    print(f"明细行       : {detail_total} 行"
          + ("" if not empty else _yellow(f"（{len(empty)} 张票没解析到明细）")))
    if bad:
        print("票级异常     : " + _red(f"{len(bad)} 张") + "（见「异常清单」sheet）")
    else:
        print("票级异常     : " + _green("0 张"))
    print(f"异常条目     : {len(issues)} 条")
    print(f"耗时         : {elapsed:.2f} 秒")
    print("-" * 78)
    print(f"{'序号':<4}{'开票日期':<12}{'发票号码':<22}{'销售方':<18}"
          f"{'金额':>10}{'税额':>9}{'价税合计':>10}{'行':>4}  校验")
    limit = 60
    for idx, inv in enumerate(invoices[:limit], start=1):
        seller = (inv.seller_name or "")[:16]
        flag = "通过" if not inv.has_issue else _red("异常")
        print(
            f"{idx:<4}{inv.invoice_date or UNREACHABLE:<12}"
            f"{(inv.invoice_no or UNREACHABLE)[:20]:<22}"
            f"{seller:<18}{money(inv.amount):>10.2f}{money(inv.tax):>9.2f}"
            f"{money(inv.total):>10.2f}{len(inv.items):>4}  {flag}"
        )
    if len(invoices) > limit:
        print(f"... 其余 {len(invoices) - limit} 张见 xlsx")
    print("-" * 78)


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_console()
    try:
        return run(argv)
    except KeyboardInterrupt:
        print("\n已取消。", file=sys.stderr)
        return 130
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        return 1


__all__ = ["build_parser", "main", "run"]
