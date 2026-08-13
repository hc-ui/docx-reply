"""Tests for the 0.2.0 features: replacements, headings, filters, docx output."""

import json
import zipfile
from xml.etree import ElementTree as ET

from conftest import anchored, comment_xml, p, r, write_docx

from docx_reply import extract_review, render_markdown
from docx_reply.cli import main


def ins(text, author="王老师", wid="10"):
    return (
        f'<w:ins w:id="{wid}" w:author="{author}" w:date="2026-08-11T03:00:00Z">'
        f"<w:r><w:t>{text}</w:t></w:r></w:ins>"
    )


def dele(text, author="王老师", wid="11"):
    return (
        f'<w:del w:id="{wid}" w:author="{author}" w:date="2026-08-11T03:01:00Z">'
        f"<w:r><w:delText>{text}</w:delText></w:r></w:del>"
    )


def heading(text, style="Heading1"):
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


# ---------------------------------------------------------------- replace


def test_adjacent_del_ins_merged_to_replace(docx_factory):
    body = p(r("模型效果") + dele("非常") + ins("显著") + r("提升。"))
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    rev = review.revisions[0]
    assert rev.kind == "replace"
    assert rev.deleted == "非常"
    assert rev.inserted == "显著"
    assert rev.text == "非常 → 显著"


def test_adjacent_ins_del_also_merged(docx_factory):
    body = p(r("模型效果") + ins("显著") + dele("非常") + r("提升。"))
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    assert review.revisions[0].text == "非常 → 显著"


def test_text_between_prevents_merge(docx_factory):
    body = p(dele("旧词") + r("中间隔了字") + ins("新词"))
    review = extract_review(docx_factory(body))
    assert [rev.kind for rev in review.revisions] == ["delete", "insert"]


def test_different_author_prevents_merge(docx_factory):
    body = p(dele("旧词", author="张老师") + ins("新词", author="王老师"))
    review = extract_review(docx_factory(body))
    assert [rev.kind for rev in review.revisions] == ["delete", "insert"]


def test_replace_in_markdown_and_json(docx_factory):
    body = p(r("效果") + dele("非常") + ins("显著") + r("好。"))
    path = docx_factory(body)
    md = render_markdown(extract_review(path))
    assert "| 替换 |" in md
    assert "非常 → 显著" in md

    payload = json.loads(__import__("docx_reply").render_json(extract_review(path)))
    rev = payload["revisions"][0]
    assert rev["kind"] == "replace"
    assert rev["deleted"] == "非常"
    assert rev["inserted"] == "显著"


# ---------------------------------------------------------------- headings


def test_comment_located_under_nearest_heading(docx_factory):
    body = (
        heading("2 相关工作")
        + p(r("无关段落。"))
        + p(anchored("需要引用的论断", "0"))
        + heading("3 方法", style="Heading1")
        + p(anchored("方法段", "1"))
    )
    comments = comment_xml("0", "王老师", "补引用。", para_id="P0") + comment_xml(
        "1", "王老师", "展开讲。", para_id="P1"
    )
    review = extract_review(docx_factory(body, comments))
    assert review.comments[0].heading == "2 相关工作"
    assert review.comments[1].heading == "3 方法"
    md = render_markdown(review)
    assert "2 相关工作 · 第3段" in md
    assert "3 方法 · 第5段" in md


def test_chinese_word_heading_style_id(docx_factory):
    body = heading("第一章 绪论", style="1") + p(anchored("正文", "0"))
    review = extract_review(docx_factory(body, comment_xml("0", "A", "评")))
    assert review.comments[0].heading == "第一章 绪论"


def test_custom_style_is_not_a_heading(docx_factory):
    body = heading("这不是标题", style="MyFancyStyle") + p(anchored("正文", "0"))
    review = extract_review(docx_factory(body, comment_xml("0", "A", "评")))
    assert review.comments[0].heading == ""


def test_revision_gets_heading(docx_factory):
    body = heading("4 实验") + p(r("精度") + ins("显著") + r("提升"))
    review = extract_review(docx_factory(body))
    assert review.revisions[0].heading == "4 实验"


# ---------------------------------------------------------------- filters


def _two_author_docx(tmp_path):
    body = p(anchored("片段一", "0")) + p(anchored("片段二", "1"))
    comments = comment_xml("0", "王老师", "意见一", para_id="P0") + comment_xml(
        "1", "张老师", "意见二", para_id="P1"
    )
    ex = '<w15:commentEx w15:paraId="P0" w15:done="1"/><w15:commentEx w15:paraId="P1" w15:done="0"/>'
    return write_docx(tmp_path / "two.docx", body, comments, ex)


def test_author_filter(tmp_path, capsys):
    path = _two_author_docx(tmp_path)
    assert main([str(path), "--author", "张老师"]) == 0
    out = capsys.readouterr().out
    assert "张老师" in out
    assert "王老师" not in out


def test_skip_resolved(tmp_path, capsys):
    path = _two_author_docx(tmp_path)
    assert main([str(path), "--skip-resolved"]) == 0
    out = capsys.readouterr().out
    assert "意见二" in out
    assert "意见一" not in out


# ---------------------------------------------------------------- docx out


def test_docx_output_roundtrip(tmp_path, capsys):
    src = _two_author_docx(tmp_path)
    out = tmp_path / "对照表.docx"
    assert main([str(src), "-f", "docx", "-o", str(out)]) == 0
    assert out.exists()
    with zipfile.ZipFile(out) as zf:
        doc = zf.read("word/document.xml").decode("utf-8")
    # well-formed XML containing the review data and the fillable column
    ET.fromstring(doc)
    assert "审阅意见对照表" in doc
    assert "意见一" in doc and "意见二" in doc
    assert "修改说明" in doc


def test_docx_format_requires_output(tmp_path, capsys):
    src = _two_author_docx(tmp_path)
    assert main([str(src), "-f", "docx"]) == 2
    assert "-o" in capsys.readouterr().err


def test_docx_output_is_parseable_by_word_readers(tmp_path):
    """The generated table document must itself be a valid minimal docx."""
    src = _two_author_docx(tmp_path)
    out = tmp_path / "table.docx"
    assert main([str(src), "-f", "docx", "-o", str(out)]) == 0
    review = extract_review(out)  # our own parser reads any valid docx
    assert review.paragraph_count > 0
    assert review.comments == []
