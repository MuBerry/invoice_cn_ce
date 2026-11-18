# SPDX-License-Identifier: Apache-2.0
# Part of invoice_cn_ce — 见仓库根目录 LICENSE。
"""CLI 端到端：退出码、错误友好度、ZIP/目录混合输入、去重开关、分组维度。

素材全部动态生成（最小 ZIP + 非 PDF 文件），不依赖真实票据。
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.conftest import NOT_A_PDF, run_cli, write_zip


def _load(out_dir: Path) -> dict:
    return json.loads((out_dir / "数据.json").read_text(encoding="utf-8"))


def test_no_arguments_prints_help_and_exits_2() -> None:
    proc = run_cli([])
    assert proc.returncode == 2
    assert "invoice_cn_ce.batch" in proc.stdout
    assert "Traceback" not in proc.stderr


def test_missing_path_is_friendly_not_a_traceback(tmp_path: Path) -> None:
    proc = run_cli([str(tmp_path / "根本不存在" / "发票.pdf"), "-o", str(tmp_path / "out")])
    assert proc.returncode == 2
    assert "没有找到任何可解析的文件" in proc.stderr
    assert "路径不存在" in proc.stderr
    assert "Traceback" not in proc.stderr


def test_empty_directory_exits_2(tmp_path: Path) -> None:
    empty = tmp_path / "空 目录"
    empty.mkdir()
    proc = run_cli([str(empty), "-o", str(tmp_path / "out")])
    assert proc.returncode == 2
    assert "没有找到任何可解析的文件" in proc.stderr


def test_zip_of_non_pdf_files_yields_no_invoice(tmp_path: Path) -> None:
    archive = write_zip(
        tmp_path / "包.zip",
        [("说明.txt", b"hi"), ("图片.jpg", b"\xff\xd8\xff"), ("坏.pdf", NOT_A_PDF)],
    )
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    payload = _load(out)
    # 非 PDF 一律不进「发票」，只进异常清单
    assert payload["invoices"] == []
    assert payload["totals"]["count"] == 0
    reasons = " ".join(i["reason"] for i in payload["issues"])
    assert "不是 PDF" in reasons
    assert all(i["level"] == "文件级" for i in payload["issues"])


def test_zero_byte_pdf_goes_to_issues_not_to_totals(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("空.pdf", b"")])
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    payload = _load(out)
    assert payload["invoices"] == []
    assert any("empty_file" in i["reason"] for i in payload["issues"])


def test_invalid_group_option_exits_2_with_hint(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    proc = run_cli([str(archive), "-o", str(tmp_path / "out"), "--group", "颜色"])
    assert proc.returncode == 2
    assert "不支持的分组维度" in proc.stderr
    assert "销售方" in proc.stderr  # 提示可选项
    assert "Traceback" not in proc.stderr


def test_multiple_group_dimensions_are_written(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--group", "购买方,月份,税率", "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    payload = _load(out)
    assert payload["meta"]["group_dims"] == ["购买方", "月份", "税率"]
    assert set(payload["groups"]) == {"购买方", "月份", "税率"}


def test_outputs_are_named_and_present(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    names = sorted(p.name for p in out.iterdir())
    assert "数据.json" in names
    assert any(n.startswith("发票汇总-") and n.endswith(".xlsx") for n in names)
    assert any(n.startswith("发票汇总-") and n.endswith(".csv") for n in names)


def test_dedup_switch_is_recorded_in_meta(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    out = tmp_path / "out"
    run_cli([str(archive), "-o", str(out), "--quiet", "--no-dedup"])
    assert _load(out)["meta"]["dedup_enabled"] is False


def test_quiet_mode_prints_no_per_file_progress(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    proc = run_cli([str(archive), "-o", str(tmp_path / "out"), "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    assert "[1/1]" not in proc.stdout


def test_progress_mode_reports_each_file(tmp_path: Path) -> None:
    write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    proc = run_cli([str(tmp_path / "包.zip"), "-o", str(tmp_path / "out")])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    assert "[1/1]" in proc.stdout


def test_mixed_inputs_directory_and_zip(tmp_path: Path) -> None:
    folder = tmp_path / "素材 目录"
    folder.mkdir()
    (folder / "a.pdf").write_bytes(NOT_A_PDF)
    archive = write_zip(tmp_path / "包.zip", [("b.pdf", b"%PDF-1.7\nother")])
    out = tmp_path / "out"
    proc = run_cli([str(folder), str(archive), "-o", str(out), "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    assert _load(out)["meta"]["file_count"] == 2


def test_keep_temp_flag_retains_extraction_dir(tmp_path: Path) -> None:
    """``--keep-temp`` 只影响临时目录生命周期，不能影响结果数量。"""
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--quiet", "--keep-temp"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    assert _load(out)["meta"]["file_count"] == 1


def test_broken_zip_does_not_crash(tmp_path: Path) -> None:
    broken = tmp_path / "坏包.zip"
    broken.write_bytes(b"PK not a zip")
    proc = run_cli([str(broken), "-o", str(tmp_path / "out")])
    assert "Traceback" not in proc.stderr
    assert proc.returncode in (0, 2, 3)  # 坏 ZIP 解析不出任何票 → 3（或路径级 2）


def test_group_option_accepts_chinese_comma(tmp_path: Path) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    out = tmp_path / "out"
    proc = run_cli([str(archive), "-o", str(out), "--group", "销售方，月份", "--quiet"])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
    assert _load(out)["meta"]["group_dims"] == ["销售方", "月份"]


@pytest.mark.parametrize("flag", ["--no-color", "--quiet"])
def test_benign_flags_keep_documented_exit_code(tmp_path: Path, flag: str) -> None:
    archive = write_zip(tmp_path / "包.zip", [("坏.pdf", NOT_A_PDF)])
    proc = run_cli([str(archive), "-o", str(tmp_path / "out"), flag])
    assert proc.returncode == 3  # 素材全是非发票 PDF：0 张有效票 → 退出码 3
