# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""批量输入层：目录遍历、ZIP / 嵌套 ZIP 解包、垃圾过滤、zip-slip 防护、内容去重。"""

from __future__ import annotations

from pathlib import Path

import pytest

from invoice_cn_ce.batch.ingest import MAX_ZIP_DEPTH, Ingestor, should_skip
from tests.conftest import NOT_A_PDF, nested_zip_bytes, write_zip, zip_bytes


# ---------------------------------------------------------------- 过滤规则


@pytest.mark.parametrize(
    "name",
    [
        "._resource.pdf",
        "~$草稿.docx",
        ".hidden.pdf",
        "Thumbs.db",
        "desktop.ini",
        ".DS_Store",
        "下载中.pdf.crdownload",
        "临时.pdf.tmp",
        "",
    ],
)
def test_should_skip_rejects_junk(name: str) -> None:
    assert should_skip(name) is True


def test_should_skip_only_looks_at_the_basename() -> None:
    """``should_skip`` 只判单个名字；目录名过滤分别由遍历 / 解包负责。"""
    assert should_skip("发票/._资源.pdf") is True
    assert should_skip("__MACOSX") is False  # 目录名由 SKIP_DIR_NAMES 负责
    assert should_skip("__MACOSX/发票/甲.pdf") is False  # 见下面的 __MACOSX 用例


def test_macosx_directory_inside_zip_is_skipped(tmp_path: Path) -> None:
    """ZIP 内 __MACOSX 目录必须被逐段过滤（BUG-1 回归：修复后此用例应通过）。"""
    archive = write_zip(
        tmp_path / "包.zip",
        [
            ("__MACOSX/发票/甲.pdf", b"%PDF-1.7\nmacosx-resource-fork"),
            ("发票/甲.pdf", b"%PDF-1.7\nreal-file"),
        ],
    )
    result = Ingestor().collect([str(archive)])
    try:
        assert len(result.files) == 1
    finally:
        result.cleanup()


@pytest.mark.parametrize("name", ["发票.pdf", "2026年11月/发票 A.pdf", "invoice.PDF", "a.xml"])
def test_should_skip_keeps_normal_files(name: str) -> None:
    assert should_skip(name) is False


# ---------------------------------------------------------------- ZIP 解包


def test_zip_expands_supported_files_and_skips_junk(tmp_path: Path) -> None:
    archive = write_zip(
        tmp_path / "包.zip",
        [
            ("发票/甲.pdf", NOT_A_PDF),
            ("发票/说明.txt", b"hi"),
            ("发票/图片.jpg", b"\xff\xd8\xff"),
            ("__MACOSX/发票/._甲.pdf", b"junk"),
            ("~$临时.docx", b"junk"),
        ],
    )
    result = Ingestor().collect([str(archive)])
    try:
        assert [Path(f.display).name for f in result.files] == ["甲.pdf"]
        assert result.files[0].archive == "包.zip"
        assert result.files[0].inner == "发票/甲.pdf"
        reasons = {Path(p).name: reason for p, reason in result.skipped}
        # .txt / .jpg 明确记录为「不支持」，而不是静默消失
        assert "说明.txt" in reasons and "不是 PDF" in reasons["说明.txt"]
        assert "图片.jpg" in reasons
        # 垃圾条目（__MACOSX / ~$）不产生任何记录，也不产生文件
        assert not any("__MACOSX" in p or "~$" in p for p, _ in result.skipped)
    finally:
        result.cleanup()


def test_display_path_traces_back_to_archive(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("深层/目录/发票.pdf", NOT_A_PDF)])
    result = Ingestor().collect([str(archive)])
    try:
        assert result.files[0].display == "包.zip!深层/目录/发票.pdf"
    finally:
        result.cleanup()


def test_nested_zip_is_expanded(tmp_path: Path) -> None:
    inner = zip_bytes([("内层.pdf", NOT_A_PDF)])
    archive = write_zip(tmp_path / "外层.zip", [("内层.zip", inner)])
    result = Ingestor().collect([str(archive)])
    try:
        assert len(result.files) == 1
        # 嵌套包里的文件仍能一路追回最外层压缩包
        assert result.files[0].archive == "外层.zip"
        assert result.files[0].inner == "内层.zip/内层.pdf"
    finally:
        result.cleanup()


def test_zip_nesting_depth_is_bounded(tmp_path: Path) -> None:
    too_deep = nested_zip_bytes(MAX_ZIP_DEPTH + 1, "最深.pdf", NOT_A_PDF)
    archive = tmp_path / "太深.zip"
    archive.write_bytes(too_deep)
    result = Ingestor().collect([str(archive)])
    try:
        assert result.files == []
        assert any("嵌套超过" in reason for _, reason in result.skipped)
    finally:
        result.cleanup()


def test_zip_slip_members_never_escape(tmp_path: Path) -> None:
    archive = write_zip(
        tmp_path / "攻击.zip",
        [
            ("../../逃逸.pdf", NOT_A_PDF),
            ("发票/../../../逃逸2.pdf", NOT_A_PDF),
            ("/绝对路径.pdf", NOT_A_PDF),
            ("C:/Windows/逃逸3.pdf", NOT_A_PDF),
            ("正常/发票.pdf", NOT_A_PDF),
        ],
    )
    result = Ingestor().collect([str(archive)])
    try:
        # 只有包内正常路径被接收
        assert [f.inner for f in result.files] == ["正常/发票.pdf"]
        # 任何成员都不能落到临时目录之外
        temp_root = result.temp_dirs[0]
        for name in ("逃逸.pdf", "逃逸2.pdf", "逃逸3.pdf"):
            written = list(Path(temp_root).rglob(name))
            assert all(str(p).startswith(str(Path(temp_root))) for p in written)
        assert not (tmp_path / "逃逸.pdf").exists()
        assert not (tmp_path.parent / "逃逸2.pdf").exists()
    finally:
        result.cleanup()


def test_broken_zip_is_reported_not_raised(tmp_path: Path) -> None:
    broken = tmp_path / "坏包.zip"
    broken.write_bytes(b"PK\x03\x04 this is not a real zip")
    result = Ingestor().collect([str(broken)])
    assert result.files == []
    assert any("ZIP 无法打开" in reason for _, reason in result.skipped)


def test_zero_byte_member_is_kept_for_parser_to_judge(tmp_path: Path) -> None:
    """零字节 .pdf 不算「输入层」的错误：交给解析层判为 empty_file。"""
    archive = write_zip(tmp_path / "包.zip", [("空.pdf", b""), ("好.pdf", NOT_A_PDF)])
    result = Ingestor().collect([str(archive)])
    try:
        assert set(f.inner for f in result.files) == {"空.pdf", "好.pdf"}
    finally:
        result.cleanup()


# ---------------------------------------------------------------- 目录 / 单文件


def test_directory_walk_is_recursive_and_sorted(tmp_path: Path) -> None:
    (tmp_path / "b").mkdir()
    (tmp_path / "a").mkdir()
    (tmp_path / "a" / "1.pdf").write_bytes(NOT_A_PDF)
    (tmp_path / "b" / "2.xml").write_bytes(b"<x/>")
    (tmp_path / ".git").mkdir()
    (tmp_path / ".git" / "config.pdf").write_bytes(NOT_A_PDF)
    result = Ingestor().collect([str(tmp_path)])
    try:
        names = [Path(f.path).name for f in result.files]
        assert names == ["1.pdf", "2.xml"]
        assert not any(".git" in f.path for f in result.files)
    finally:
        result.cleanup()


def test_missing_path_is_recorded(tmp_path: Path) -> None:
    result = Ingestor().collect([str(tmp_path / "根本不存在.pdf")])
    assert result.files == []
    assert result.skipped and result.skipped[0][1] == "路径不存在"


def test_unsupported_single_file_is_recorded(tmp_path: Path) -> None:
    target = tmp_path / "说明.txt"
    target.write_text("hi", encoding="utf-8")
    result = Ingestor().collect([str(target)])
    assert result.files == []
    assert "不是 PDF" in result.skipped[0][1]


# ---------------------------------------------------------------- 内容去重


def test_identical_content_is_deduped_by_hash(tmp_path: Path) -> None:
    payload = b"%PDF-1.7\nidentical"
    (tmp_path / "a.pdf").write_bytes(payload)
    (tmp_path / "b.pdf").write_bytes(payload)
    (tmp_path / "c.pdf").write_bytes(b"%PDF-1.7\ndifferent")
    result = Ingestor().collect([str(tmp_path)])
    try:
        assert [Path(f.path).name for f in result.files] == ["a.pdf", "c.pdf"]
        assert len(result.hash_duplicates) == 1
        duplicate, kept = result.hash_duplicates[0]
        assert Path(duplicate).name == "b.pdf"
        assert Path(kept).name == "a.pdf"
    finally:
        result.cleanup()


def test_same_file_listed_twice_is_deduped(tmp_path: Path) -> None:
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(NOT_A_PDF)
    result = Ingestor().collect([str(pdf), str(pdf)])
    try:
        assert len(result.files) == 1
        assert len(result.hash_duplicates) == 1
    finally:
        result.cleanup()


def test_temp_dirs_are_removed_by_cleanup(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("发票.pdf", NOT_A_PDF)])
    result = Ingestor().collect([str(archive)])
    dirs = list(result.temp_dirs)
    assert dirs and Path(dirs[0]).exists()
    result.cleanup()
    assert not Path(dirs[0]).exists()
