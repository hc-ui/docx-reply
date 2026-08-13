"""Extraction-fidelity tests: text boxes, formulas, ruby text, docx table schema."""

from conftest import anchored, comment_xml, p, r

from docx_reply import extract_review, render_docx


def _textbox(inner: str) -> str:
    """A run holding a text box, duplicated as mc:Choice + mc:Fallback
    exactly like Word writes it (DrawingML and legacy VML carry the same
    w:txbxContent)."""
    content = f"<w:txbxContent>{inner}</w:txbxContent>"
    return (
        "<w:r><mc:AlternateContent>"
        f'<mc:Choice Requires="wps"><w:drawing>{content}</w:drawing></mc:Choice>'
        f"<mc:Fallback><w:pict>{content}</w:pict></mc:Fallback>"
        "</mc:AlternateContent></w:r>"
    )


# ------------------------------------------------- text boxes


def test_textbox_comment_extracted_once(docx_factory):
    body = p(r("正文。") + _textbox(p(anchored("文本框里的结论", "0"))))
    review = extract_review(docx_factory(body, comment_xml("0", "王老师", "结论太绝对。")))
    assert len(review.comments) == 1
    assert review.comments[0].quoted == "文本框里的结论"


def test_textbox_revision_not_duplicated(docx_factory):
    box = _textbox(
        "<w:p>"
        '<w:ins w:id="10" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + r("框内新增")
        + "</w:ins></w:p>"
    )
    review = extract_review(docx_factory(p(r("正文。") + box)))
    # mc:Choice and mc:Fallback hold the same content: count it once
    assert [rev.text for rev in review.revisions] == ["框内新增"]


def test_textbox_with_only_fallback(docx_factory):
    content = "<w:txbxContent>" + p(anchored("仅兼容形式", "0")) + "</w:txbxContent>"
    body = p(
        "<w:r><mc:AlternateContent>"
        f"<mc:Fallback><w:pict>{content}</w:pict></mc:Fallback>"
        "</mc:AlternateContent></w:r>"
    )
    review = extract_review(docx_factory(body, comment_xml("0", "A", "看这里")))
    assert review.comments[0].quoted == "仅兼容形式"


def test_plain_drawing_textbox_without_alternatecontent(docx_factory):
    body = p(
        "<w:r><w:drawing><w:txbxContent>"
        + p(anchored("直接画布内容", "0"))
        + "</w:txbxContent></w:drawing></w:r>"
    )
    review = extract_review(docx_factory(body, comment_xml("0", "A", "评")))
    assert review.comments[0].quoted == "直接画布内容"


# ------------------------------------------------- formulas


def math(text: str) -> str:
    return f"<m:oMath><m:r><m:t>{text}</m:t></m:r></m:oMath>"


def test_comment_on_formula_quotes_linearized_math(docx_factory):
    body = p(
        '<w:commentRangeStart w:id="0"/>'
        + math("E=mc²")
        + '<w:commentRangeEnd w:id="0"/>'
        + '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    review = extract_review(docx_factory(body, comment_xml("0", "王老师", "单位不对。")))
    assert review.comments[0].quoted == "E=mc²"
    assert review.comments[0].para_text == "E=mc²"


def test_deleted_formula_produces_revision(docx_factory):
    body = p(
        r("由")
        + '<w:del w:id="5" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + math("x=1")
        + "</w:del>"
        + r("可知")
    )
    review = extract_review(docx_factory(body))
    assert [(rev.kind, rev.text) for rev in review.revisions] == [("delete", "x=1")]
    assert review.revisions[0].para_text == "由可知"


def test_formula_between_del_and_ins_prevents_replace_merge(docx_factory):
    body = p(
        '<w:del w:id="1" w:author="A" w:date="2026-08-11T00:00:00Z">'
        "<w:r><w:delText>旧</w:delText></w:r></w:del>"
        + math("y=2")
        + '<w:ins w:id="2" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + r("新")
        + "</w:ins>"
    )
    review = extract_review(docx_factory(body))
    assert sorted(rev.kind for rev in review.revisions) == ["delete", "insert"]


# ------------------------------------------------- ruby (phonetic guides)


def test_ruby_base_kept_annotation_dropped(docx_factory):
    body = p(
        "<w:r><w:ruby>"
        "<w:rt><w:r><w:t>pīn yīn</w:t></w:r></w:rt>"
        "<w:rubyBase><w:r><w:t>拼音</w:t></w:r></w:rubyBase>"
        "</w:ruby></w:r>"
        + r("正文")
    )
    review = extract_review(docx_factory(body))
    assert review.paragraph_count == 1
    # verify via a comment range over the whole paragraph
    body2 = p(
        '<w:commentRangeStart w:id="0"/>'
        "<w:r><w:ruby>"
        "<w:rt><w:r><w:t>pīn yīn</w:t></w:r></w:rt>"
        "<w:rubyBase><w:r><w:t>拼音</w:t></w:r></w:rubyBase>"
        "</w:ruby></w:r>"
        + r("正文")
        + '<w:commentRangeEnd w:id="0"/>'
        + '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    review2 = extract_review(docx_factory(body2, comment_xml("0", "A", "评")))
    assert review2.comments[0].quoted == "拼音正文"


# ------------------------------------------------- format-merge precision


def _fmt_run(text: str, author: str = "王老师") -> str:
    return (
        f'<w:r><w:rPr><w:rPrChange w:id="40" w:author="{author}" '
        'w:date="2026-08-11T05:00:00Z"><w:rPr/></w:rPrChange></w:rPr>'
        f"<w:t>{text}</w:t></w:r>"
    )


def test_two_separate_format_ranges_stay_separate(docx_factory):
    # "AAA"(formatted) + plain text + "CCC"(formatted) are two distinct
    # formatting actions: they must not collapse into a misleading "AAACCC"
    body = p(_fmt_run("AAA") + r("中间普通文字") + _fmt_run("CCC"))
    review = extract_review(docx_factory(body))
    assert [rev.text for rev in review.revisions] == ["AAA", "CCC"]


def test_fragmented_format_range_still_merges(docx_factory):
    body = p(_fmt_run("一句被") + _fmt_run("拆碎的话"))
    review = extract_review(docx_factory(body))
    assert [rev.text for rev in review.revisions] == ["一句被拆碎的话"]


# ------------------------------------------------- misc robustness


def test_header_files_in_natural_order(docx_factory):
    from conftest import header_part

    headers = {
        "word/header1.xml": header_part(p(anchored("眉一", "1"))),
        "word/header2.xml": header_part(p(anchored("眉二", "2"))),
        "word/header10.xml": header_part(p(anchored("眉十", "3"))),
    }
    path = docx_factory(
        p(r("正文。")),
        comments=comment_xml("1", "A", "评一", para_id="A1")
        + comment_xml("2", "A", "评二", para_id="A2")
        + comment_xml("3", "A", "评三", para_id="A3"),
        extra_parts=headers,
    )
    review = extract_review(path)
    idx = {c.id: c.para_index for c in review.comments}
    # header2 comes before header10 (natural, not lexicographic, order)
    assert idx["1"] < idx["2"] < idx["3"]


def test_unterminated_comment_range_gets_best_effort_quote(docx_factory):
    body = p('<w:commentRangeStart w:id="0"/>' + r("没有终点的范围"))
    review = extract_review(docx_factory(body, comment_xml("0", "A", "评")))
    assert review.comments[0].quoted == "没有终点的范围"


def test_no_break_hyphen_kept(docx_factory):
    # verify via a comment anchored on the paragraph
    body2 = p(
        '<w:commentRangeStart w:id="0"/>'
        "<w:r><w:t>state</w:t><w:noBreakHyphen/><w:t>art</w:t></w:r>"
        '<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    review2 = extract_review(docx_factory(body2, comment_xml("0", "A", "评")))
    assert review2.comments[0].quoted == "state-art"


# ------------------------------------------------- docx output schema


def test_docx_table_has_tblgrid(docx_factory):
    review = extract_review(docx_factory(p(anchored("文字", "0")), comment_xml("0", "A", "评")))
    payload = render_docx(review)
    import io
    import zipfile

    with zipfile.ZipFile(io.BytesIO(payload)) as zf:
        doc = zf.read("word/document.xml").decode("utf-8")
    assert "<w:tblGrid>" in doc
    assert doc.count("<w:gridCol") >= 9
