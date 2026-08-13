"""Tests for the 0.3.0 features: extra document stories, move & format revisions."""

import json

from docx_reply import extract_review, render_json, render_markdown
from conftest import anchored, comment_xml, footnotes_part, header_part, p, r


# ------------------------------------------------- comments beyond the body


def test_comment_in_footnote(docx_factory):
    footnotes = footnotes_part(
        '<w:footnote w:id="1">' + p(anchored("脚注里的引文", "0")) + "</w:footnote>"
    )
    path = docx_factory(
        p(r("正文段落。")),
        comments=comment_xml("0", "王老师", "脚注也要给出页码。"),
        extra_parts={"word/footnotes.xml": footnotes},
    )
    review = extract_review(path)
    assert len(review.comments) == 1
    c = review.comments[0]
    assert c.part == "脚注"
    assert c.quoted == "脚注里的引文"
    assert c.para_text == "脚注里的引文"
    md = render_markdown(review)
    assert "脚注 · 第1段" in md


def test_comment_in_header(docx_factory):
    header = header_part(p(anchored("页眉文字", "0")))
    path = docx_factory(
        p(r("正文。")),
        comments=comment_xml("0", "李老师", "页眉格式不对。"),
        extra_parts={"word/header1.xml": header},
    )
    review = extract_review(path)
    assert review.comments[0].part == "页眉"
    assert review.comments[0].quoted == "页眉文字"


def test_body_comments_sorted_before_footnote_comments(docx_factory):
    footnotes = footnotes_part(
        '<w:footnote w:id="1">' + p(anchored("脚注片段", "5")) + "</w:footnote>"
    )
    path = docx_factory(
        p(anchored("正文片段", "9")),
        comments=comment_xml("5", "A", "脚注意见", para_id="P5")
        + comment_xml("9", "B", "正文意见", para_id="P9"),
        extra_parts={"word/footnotes.xml": footnotes},
    )
    review = extract_review(path)
    assert [c.id for c in review.comments] == ["9", "5"]


def test_revision_in_footnote_carries_part(docx_factory):
    footnotes = footnotes_part(
        '<w:footnote w:id="1"><w:p>'
        '<w:ins w:id="10" w:author="王老师" w:date="2026-08-11T03:00:00Z">'
        + r("补充的脚注")
        + "</w:ins></w:p></w:footnote>"
    )
    path = docx_factory(p(r("正文。")), extra_parts={"word/footnotes.xml": footnotes})
    review = extract_review(path)
    assert len(review.revisions) == 1
    assert review.revisions[0].part == "脚注"
    assert review.revisions[0].text == "补充的脚注"
    # body paragraph count is not affected by footnote paragraphs
    assert review.paragraph_count == 1


def test_json_part_field(docx_factory):
    footnotes = footnotes_part(
        '<w:footnote w:id="1">' + p(anchored("片段", "0")) + "</w:footnote>"
    )
    path = docx_factory(
        p(r("正文。")),
        comments=comment_xml("0", "A", "评"),
        extra_parts={"word/footnotes.xml": footnotes},
    )
    payload = json.loads(render_json(extract_review(path)))
    assert payload["comments"][0]["part"] == "脚注"


# ------------------------------------------------- format revisions


def test_run_format_change_reported(docx_factory):
    body = p(
        r("前文")
        + '<w:r><w:rPr><w:b/><w:rPrChange w:id="30" w:author="王老师" w:date="2026-08-11T05:00:00Z">'
        '<w:rPr/></w:rPrChange></w:rPr><w:t>被加粗的关键结论</w:t></w:r>'
        + r("后文")
    )
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    rev = review.revisions[0]
    assert rev.kind == "format"
    assert rev.author == "王老师"
    assert rev.text == "被加粗的关键结论"
    md = render_markdown(review)
    assert "| 格式 |" in md


def test_consecutive_format_runs_merged(docx_factory):
    def fmt_run(text):
        return (
            '<w:r><w:rPr><w:rPrChange w:id="31" w:author="王老师" '
            'w:date="2026-08-11T05:00:00Z"><w:rPr/></w:rPrChange></w:rPr>'
            f"<w:t>{text}</w:t></w:r>"
        )

    body = p(fmt_run("一句被") + fmt_run("拆成多个") + fmt_run("run 的话"))
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    assert review.revisions[0].text == "一句被拆成多个run 的话"


def test_paragraph_format_change_reported(docx_factory):
    body = (
        "<w:p><w:pPr><w:pPrChange w:id=\"32\" w:author=\"张老师\" "
        "w:date=\"2026-08-11T06:00:00Z\"><w:pPr/></w:pPrChange></w:pPr>"
        + r("整段居中了")
        + "</w:p>"
    )
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    rev = review.revisions[0]
    assert rev.kind == "format"
    assert rev.author == "张老师"
    assert rev.text == "整段居中了"


def test_format_change_inside_insertion_not_double_counted(docx_factory):
    body = p(
        '<w:ins w:id="10" w:author="A" w:date="2026-08-11T00:00:00Z">'
        '<w:r><w:rPr><w:rPrChange w:id="33" w:author="A" w:date="2026-08-11T00:00:00Z">'
        "<w:rPr/></w:rPrChange></w:rPr><w:t>新增文字</w:t></w:r></w:ins>"
    )
    review = extract_review(docx_factory(body))
    assert [rev.kind for rev in review.revisions] == ["insert"]


# ------------------------------------------------- interactions


def test_replace_merge_still_works_with_format_rows_around(docx_factory):
    body = p(
        '<w:r><w:rPr><w:rPrChange w:id="34" w:author="B" w:date="2026-08-11T00:00:00Z">'
        "<w:rPr/></w:rPrChange></w:rPr><w:t>格式变了的</w:t></w:r>"
        + '<w:del w:id="1" w:author="A" w:date="2026-08-11T00:00:00Z">'
        "<w:r><w:delText>旧</w:delText></w:r></w:del>"
        + '<w:ins w:id="2" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + r("新")
        + "</w:ins>"
    )
    review = extract_review(docx_factory(body))
    kinds = sorted(rev.kind for rev in review.revisions)
    assert kinds == ["format", "replace"]


def test_move_kind_rendered_in_markdown(docx_factory):
    body = p(
        '<w:moveTo w:id="20" w:author="A" w:date="2026-08-11T00:00:00Z">' + r("挪过来的段落") + "</w:moveTo>"
    )
    md = render_markdown(extract_review(docx_factory(body)))
    assert "| 移动 |" in md
    assert "挪过来的段落" in md
