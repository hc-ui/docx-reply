import json

import pytest

from docx_reply.cli import main
from conftest import anchored, comment_xml, p, r, write_docx


def reviewed_docx(tmp_path):
    body = (
        p(r("第3章 实验与分析"))
        + p(anchored("深度学习模型取得了显著进展", "0") + r("，尤其是卷积神经网络。"))
        + p(
            r("模型效果")
            + '<w:ins w:id="10" w:author="王老师" w:date="2026-08-11T03:00:00Z">'
            + r("显著")
            + "</w:ins>"
            + r("提升。")
        )
    )
    comments = comment_xml("0", "王老师", "这里需要补充引用来源。", para_id="P0")
    return write_docx(tmp_path / "reviewed.docx", body, comments)


def test_markdown_to_stdout(tmp_path, capsys):
    path = reviewed_docx(tmp_path)
    assert main([str(path)]) == 0
    captured = capsys.readouterr()
    assert "# 审阅意见对照表：reviewed.docx" in captured.out
    assert "王老师" in captured.out
    assert "共 1 条批注（含 0 条回复）、1 处修订" in captured.err


def test_write_markdown_file(tmp_path, capsys):
    path = reviewed_docx(tmp_path)
    out = tmp_path / "对照表.md"
    assert main([str(path), "-o", str(out)]) == 0
    assert "修改对照表" in out.read_text(encoding="utf-8")


def test_csv_file_has_bom_and_revisions_sidecar(tmp_path):
    path = reviewed_docx(tmp_path)
    out = tmp_path / "对照表.csv"
    assert main([str(path), "-f", "csv", "-o", str(out)]) == 0
    raw = out.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    sidecar = tmp_path / "对照表.revisions.csv"
    assert sidecar.exists()
    assert "插入" in sidecar.read_text(encoding="utf-8-sig")


def test_json_output_parses(tmp_path, capsys):
    path = reviewed_docx(tmp_path)
    assert main([str(path), "-f", "json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["stats"]["comments"] == 1
    assert payload["stats"]["revisions"] == 1


def test_no_revisions_flag(tmp_path, capsys):
    path = reviewed_docx(tmp_path)
    assert main([str(path), "--no-revisions"]) == 0
    captured = capsys.readouterr()
    assert "## 修订记录" not in captured.out
    assert "0 处修订" in captured.err


def test_invalid_input_exit_code(tmp_path, capsys):
    bogus = tmp_path / "fake.docx"
    bogus.write_text("not a docx", encoding="utf-8")
    assert main([str(bogus)]) == 2
    assert "错误" in capsys.readouterr().err


def test_version_flag(capsys):
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0
    assert "docx-reply" in capsys.readouterr().out
