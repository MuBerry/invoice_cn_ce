# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""批量输入层：把「单文件 / 目录 / ZIP（含嵌套）」展开成待解析文件清单。

设计要点：

* **过滤垃圾条目**：``__MACOSX``、``._*``（macOS 资源分叉）、``~$*``（Office 临时
  文件）、``.`` 开头的隐藏文件、``Thumbs.db`` / ``desktop.ini`` 一律跳过。ZIP 条目
  没有目录遍历帮忙剪枝，因此按**路径分段**逐段过滤（:func:`is_junk_path`）。
* **通配符**：输入里的 ``*`` / ``?`` / ``[`` 按 :mod:`glob` 展开（Windows cmd 不展开），
  支持 ``**`` 递归；一个都没匹配上时仍按「路径不存在」上报，不静默吞掉。
* **可溯源**：每个文件都带上「来源包 → 包内路径」，报表里能一路追回原始压缩包。
* **两级去重**：同一份文件在多个输入里出现时按内容 sha1 只解析一次；发票号码级
  去重留给 :mod:`invoice_cn_ce.batch.aggregate`。
* **安全**：解包前做 zip-slip 校验，成员解析后的绝对路径必须仍在临时目录内。
"""

from __future__ import annotations

import glob
import hashlib
import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Iterable, List, Tuple

#: 支持解析的扩展名 -> 交给 ``formats.parse_pdf`` 自行识别 XML / OFD / PDF。
SUPPORTED_EXTS = (".pdf", ".xml", ".ofd")
ZIP_EXT = ".zip"
#: 嵌套 ZIP 最大层数（防止 zip 炸弹式的递归结构）。
MAX_ZIP_DEPTH = 3
#: 解包时忽略的目录名（小写比较）。
SKIP_DIR_NAMES = frozenset({"__macosx", ".git", ".svn", ".idea", ".vscode"})
#: 解包时忽略的文件名（小写比较）。
SKIP_FILE_NAMES = frozenset({"thumbs.db", "desktop.ini", ".ds_store", "icon\r"})


def should_skip(name: str) -> bool:
    """判断遍历/解包时是否应忽略该文件或目录名。"""
    base = os.path.basename((name or "").replace("\\", "/")).strip()
    if not base:
        return True
    low = base.lower()
    if low in SKIP_FILE_NAMES:
        return True
    if base.startswith("."):  # 隐藏文件 & ._xxx
        return True
    if base.startswith("~$"):  # Office / WPS 临时文件
        return True
    if low.endswith(".crdownload") or low.endswith(".tmp"):
        return True
    return False


def is_junk_path(rel_path: str) -> bool:
    """相对路径（``发票/__MACOSX/._甲.pdf``）是否整体应被忽略。

    逐段检查：**任一分段**命中 :func:`should_skip`（隐藏文件、``._*``、``~$*``、
    ``Thumbs.db`` …）或 :data:`SKIP_DIR_NAMES`（``__MACOSX``、``.git`` …）就丢弃。
    ZIP 里的条目是「一整条路径字符串」，没有目录遍历能替我们剪枝，所以必须在这里
    逐段判；空路径（如纯目录条目）也视为垃圾。
    """
    parts = [p for p in (rel_path or "").replace("\\", "/").split("/") if p not in ("", ".")]
    if not parts:
        return True
    return any(should_skip(part) or part.lower() in SKIP_DIR_NAMES for part in parts)


#: 输入路径里出现这些字符就按通配符展开（Windows cmd 不展开 ``*.pdf``，只能自己来）。
_GLOB_CHARS = ("*", "?", "[")


def is_glob_pattern(text: str) -> bool:
    return any(ch in text for ch in _GLOB_CHARS)


def expand_input_path(raw: str) -> List[str]:
    """把用户输入的一个「路径参数」展开成绝对路径列表。

    * 普通路径：原样返回（后续由 ``collect`` 判断存在性并给出友好提示）；
    * 通配符（``*.pdf`` / ``**/票.pdf`` / ``发票?[0-9].pdf``）：按 :mod:`glob`
      展开（``recursive=True`` 支持 ``**``）；**一个都没匹配上**时仍返回原路径，
      让上层照常报「路径不存在」，而不是静默产出 0 份文件。
    """
    text = str(raw).strip().strip('"')
    if not text:
        return []
    pattern = os.path.expanduser(text)
    if not is_glob_pattern(pattern):
        return [os.path.abspath(pattern)]
    matches = sorted(glob.glob(pattern, recursive=True))
    if not matches:
        return [os.path.abspath(pattern)]
    return [os.path.abspath(item) for item in matches]


def sha1_of(path: str, chunk_size: int = 1 << 20) -> str:
    """文件内容 sha1。"""
    digest = hashlib.sha1()
    with open(path, "rb") as fp:
        while True:
            block = fp.read(chunk_size)
            if not block:
                break
            digest.update(block)
    return digest.hexdigest()


@dataclass
class SourceFile:
    """一个待解析文件的来源信息。"""

    path: str = ""       # 实际落盘路径
    display: str = ""    # 展示路径（ZIP 内文件形如 ``包名.zip!包内/路径.pdf``）
    archive: str = ""    # 来源 ZIP（空表示直接输入）
    inner: str = ""      # ZIP 内相对路径
    sha1: str = ""
    size: int = 0

    @property
    def ext(self) -> str:
        return os.path.splitext(self.path)[1].lower()

    def to_dict(self) -> dict:
        return {
            "file": self.display or self.path,
            "archive": self.archive,
            "inner_path": self.inner,
            "size": self.size,
            "sha1": self.sha1,
        }


@dataclass
class IngestResult:
    """输入收集结果。"""

    files: List[SourceFile] = field(default_factory=list)
    #: ``(展示路径, 原因)`` —— 被忽略的输入。
    skipped: List[Tuple[str, str]] = field(default_factory=list)
    #: 按内容 hash 去重：``(重复文件, 保留文件)``。
    hash_duplicates: List[Tuple[str, str]] = field(default_factory=list)
    temp_dirs: List[str] = field(default_factory=list)
    #: 用户传入的原始路径（写进报表「参数与说明」）。
    inputs: List[str] = field(default_factory=list)

    def cleanup(self) -> None:
        """删除解包用的临时目录。"""
        for path in self.temp_dirs:
            shutil.rmtree(path, ignore_errors=True)
        self.temp_dirs.clear()


class Ingestor:
    """把任意输入路径展开成文件清单。"""

    def __init__(self) -> None:
        self.result = IngestResult()
        self._seen_hash: dict = {}
        self._temp_root = ""

    # ------------------------------------------------------------ 对外

    def collect(self, paths: Iterable[str]) -> IngestResult:
        for raw in paths:
            if not raw:
                continue
            # 原始输入照录（报表「参数与说明」要能看到用户实际写了什么）
            self.result.inputs.append(os.path.abspath(
                os.path.expanduser(str(raw).strip().strip('"'))))
            for path in expand_input_path(str(raw)):
                if not os.path.exists(path):
                    self.result.skipped.append((path, "路径不存在"))
                    continue
                low = path.lower()
                if os.path.isdir(path):
                    self._walk_dir(path, archive="", inner_prefix="")
                elif low.endswith(ZIP_EXT):
                    self._expand_zip(path, depth=1, archive_name=os.path.basename(path),
                                     inner_prefix="", origin=path)
                elif low.endswith(SUPPORTED_EXTS):
                    self._add_file(path, archive="", inner="", display=path)
                else:
                    self.result.skipped.append((path, "不是 PDF / XML / OFD / ZIP，已忽略"))
        return self.result

    # ------------------------------------------------------------ 内部

    def _walk_dir(self, root: str, archive: str, inner_prefix: str) -> None:
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = sorted(d for d in dirnames if d.lower() not in SKIP_DIR_NAMES)
            for filename in sorted(filenames):
                if should_skip(filename):
                    continue
                full = os.path.join(dirpath, filename)
                rel = os.path.relpath(full, root).replace("\\", "/")
                inner = f"{inner_prefix}{rel}" if inner_prefix else rel
                display = f"{archive}!{inner}" if archive else full
                low = filename.lower()
                if low.endswith(SUPPORTED_EXTS):
                    self._add_file(full, archive=archive, inner=inner, display=display)
                elif low.endswith(ZIP_EXT):
                    self._expand_zip(full, depth=1, archive_name=filename,
                                     inner_prefix="", origin=display)
                else:
                    self.result.skipped.append((display, "不是 PDF / XML / OFD，已忽略"))

    def _expand_zip(self, zip_path: str, depth: int, archive_name: str,
                    inner_prefix: str, origin: str) -> None:
        """解包 ZIP，并递归处理其中的嵌套 ZIP。"""
        if depth > MAX_ZIP_DEPTH:
            self.result.skipped.append((origin, f"ZIP 嵌套超过 {MAX_ZIP_DEPTH} 层，已忽略"))
            return
        try:
            archive = zipfile.ZipFile(zip_path)
        except (zipfile.BadZipFile, OSError) as exc:
            self.result.skipped.append((origin, f"ZIP 无法打开：{exc}"))
            return
        with archive:
            if not self._temp_root:
                self._temp_root = tempfile.mkdtemp(prefix="invoice_cn_ce_batch_")
                self.result.temp_dirs.append(self._temp_root)
            tag = hashlib.sha1(f"{zip_path}|{depth}".encode("utf-8")).hexdigest()[:8]
            dest = os.path.join(self._temp_root, f"z{depth}_{tag}")
            os.makedirs(dest, exist_ok=True)
            for info in archive.infolist():
                if info.is_dir():
                    continue
                name = info.filename.replace("\\", "/")
                # ZIP 条目是一整条路径串，没有目录遍历帮忙剪枝 —— 逐段过滤
                # __MACOSX / .git / ._* / ~$* / 隐藏文件，缺一不可。
                if is_junk_path(name):
                    continue
                target = self._safe_join(dest, name)
                if target is None:
                    self.result.skipped.append(
                        (f"{archive_name}!{name}", "ZIP 内路径不安全，已忽略")
                    )
                    continue
                os.makedirs(os.path.dirname(target), exist_ok=True)
                try:
                    with archive.open(info) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                except (OSError, RuntimeError, zipfile.BadZipFile) as exc:
                    self.result.skipped.append((f"{archive_name}!{name}", f"解包失败：{exc}"))
                    continue
                low = name.lower()
                inner = f"{inner_prefix}{name}"
                display = f"{archive_name}!{inner}"
                if low.endswith(SUPPORTED_EXTS):
                    self._add_file(target, archive=archive_name, inner=inner, display=display)
                elif low.endswith(ZIP_EXT):
                    self._expand_zip(target, depth=depth + 1, archive_name=archive_name,
                                     inner_prefix=f"{inner}/", origin=display)
                else:
                    self.result.skipped.append((display, "不是 PDF / XML / OFD，已忽略"))

    @staticmethod
    def _safe_join(root: str, member: str):
        """防 zip-slip：解析后的目标必须仍在 ``root`` 之内。"""
        root_abs = os.path.abspath(root)
        target = os.path.abspath(os.path.join(root_abs, member))
        if target != root_abs and not target.startswith(root_abs + os.sep):
            return None
        return target

    def _add_file(self, path: str, archive: str, inner: str, display: str) -> None:
        try:
            size = os.path.getsize(path)
            digest = sha1_of(path)
        except OSError as exc:
            self.result.skipped.append((display, f"文件读取失败：{exc}"))
            return
        if digest in self._seen_hash:
            self.result.hash_duplicates.append((display, self._seen_hash[digest]))
            return
        self._seen_hash[digest] = display
        self.result.files.append(
            SourceFile(path=path, display=display, archive=archive, inner=inner,
                       sha1=digest, size=size)
        )


__all__ = [
    "IngestResult",
    "Ingestor",
    "SourceFile",
    "SUPPORTED_EXTS",
    "expand_input_path",
    "is_glob_pattern",
    "is_junk_path",
    "should_skip",
]
