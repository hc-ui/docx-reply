"""Hostile-input tests: malformed, incomplete and randomly generated documents."""

import random

import pytest
from conftest import comment_xml, p, r

from docx_reply import DocxReviewError, extract_review, render_json, render_markdown
from docx_reply.cli import main


def test_deeply_nested_document_raises_friendly_error(docx_factory):
    body = "<w:x>" * 5000 + r("深处的文字") + "</w:x>" * 5000
    with pytest.raises(DocxReviewError, match="嵌套过深"):
        extract_review(docx_factory(body))


def test_range_end_without_start_is_ignored(docx_factory):
    body = p('<w:commentRangeEnd w:id="0"/>' + r("正文") + '<w:r><w:commentReference w:id="0"/></w:r>')
    review = extract_review(docx_factory(body, comment_xml("0", "A", "评")))
    assert review.comments[0].quoted == ""
    assert review.comments[0].para_text == "正文"


def test_duplicate_comment_ids_do_not_crash(docx_factory):
    body = p(
        '<w:commentRangeStart w:id="0"/>' + r("第一处") + '<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r>'
    ) + p(
        '<w:commentRangeStart w:id="0"/>' + r("第二处") + '<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    comments = comment_xml("0", "A", "评一", para_id="P1") + comment_xml("0", "B", "评二", para_id="P2")
    review = extract_review(docx_factory(body, comments))
    assert len(review.comments) == 2


def test_missing_attributes_do_not_crash(docx_factory):
    body = p(
        "<w:ins>" + r("无属性插入") + "</w:ins>"
        + "<w:del><w:r><w:delText>无属性删除</w:delText></w:r></w:del>"
        + "<w:r><w:commentReference/></w:r>"
    )
    comments = '<w:comment><w:p><w:r><w:t>无 id 的批注</w:t></w:r></w:p></w:comment>'
    review = extract_review(docx_factory(body, comments))
    # both revisions kept even though author/date are missing
    assert sorted(rev.kind for rev in review.revisions) == ["delete", "insert"]
    assert all(rev.author == "" for rev in review.revisions)
    assert len(review.comments) == 1


def test_empty_body_renders_fine(docx_factory):
    review = extract_review(docx_factory(""))
    assert review.paragraph_count == 0
    assert "未发现批注" in render_markdown(review)
    render_json(review)


def test_paragraph_with_only_properties(docx_factory):
    review = extract_review(docx_factory("<w:p><w:pPr/></w:p>"))
    assert review.paragraph_count == 1


def test_author_filter_strips_whitespace(docx_factory, tmp_path, capsys):
    body = p(
        '<w:commentRangeStart w:id="0"/>' + r("文字") + '<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    # Word occasionally stores the author with a trailing space
    path = docx_factory(body, comment_xml("0", "王老师 ", "意见"))
    out = tmp_path / "o.md"
    assert main([str(path), "--author", "王老师", "-o", str(out)]) == 0
    assert "王老师" in out.read_text(encoding="utf-8")
    assert "意见" in out.read_text(encoding="utf-8")


# ------------------------------------------------- randomized chaos


def _chaos_body(rng: random.Random, depth: int = 0) -> str:
    """Random document fragment from the grammar the parser understands."""
    if depth > 3:
        return r(f"叶{rng.randrange(100)}")
    n = rng.randrange(1, 4)
    pieces = []
    for _ in range(n):
        choice = rng.randrange(10)
        inner = _chaos_body(rng, depth + 1)
        if choice == 0:
            pieces.append(f"<w:p>{inner}</w:p>")
        elif choice == 1:
            pieces.append(r(f"文{rng.randrange(100)}"))
        elif choice == 2:
            pieces.append(
                f'<w:ins w:id="{rng.randrange(90)}" w:author="A" '
                f'w:date="2026-08-11T00:00:00Z">{inner}</w:ins>'
            )
        elif choice == 3:
            pieces.append(
                f'<w:del w:id="{rng.randrange(90)}" w:author="B" '
                f'w:date="2026-08-11T00:00:00Z"><w:r><w:delText>删{rng.randrange(100)}'
                "</w:delText></w:r></w:del>"
            )
        elif choice == 4:
            cid = rng.randrange(4)
            pieces.append(f'<w:commentRangeStart w:id="{cid}"/>')
        elif choice == 5:
            cid = rng.randrange(4)
            pieces.append(f'<w:commentRangeEnd w:id="{cid}"/>')
        elif choice == 6:
            pieces.append(f'<w:r><w:commentReference w:id="{rng.randrange(4)}"/></w:r>')
        elif choice == 7:
            pieces.append(f"<m:oMath><m:r><m:t>x{rng.randrange(10)}</m:t></m:r></m:oMath>")
        elif choice == 8:
            pieces.append(
                "<w:r><mc:AlternateContent><mc:Choice><w:txbxContent>"
                f"<w:p>{inner}</w:p></w:txbxContent></mc:Choice>"
                "<mc:Fallback><w:txbxContent><w:p/></w:txbxContent></mc:Fallback>"
                "</mc:AlternateContent></w:r>"
            )
        else:
            pieces.append(f"<w:unknown>{inner}</w:unknown>")
    return "".join(pieces)


def test_chaos_documents_never_crash(docx_factory):
    comments = "".join(comment_xml(str(i), "A", f"评{i}") for i in range(4))
    for seed in range(40):
        rng = random.Random(seed)
        body = "".join(f"<w:p>{_chaos_body(rng)}</w:p>" for _ in range(rng.randrange(1, 4)))
        review = extract_review(docx_factory(body, comments))
        render_markdown(review)
        render_json(review)
