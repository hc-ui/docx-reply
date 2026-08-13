import pytest

from docx_reply import DocxReviewError, extract_review
from conftest import anchored, comment_xml, p, r


def test_document_without_comments(docx_factory):
    path = docx_factory(p(r("这是一段普通文字。")) + p(r("第二段。")))
    review = extract_review(path)
    assert review.comments == []
    assert review.revisions == []
    assert review.paragraph_count == 2
    assert review.total_comments == 0


def test_single_comment_fields(docx_factory):
    body = p(r("引言部分。") + anchored("深度学习模型取得了显著进展", "0") + r("，后文继续。"))
    comments = comment_xml("0", "王老师", "这里需要补充引用来源。", para_id="AAAA0001")
    review = extract_review(docx_factory(body, comments))

    assert len(review.comments) == 1
    c = review.comments[0]
    assert c.author == "王老师"
    assert c.text == "这里需要补充引用来源。"
    assert c.quoted == "深度学习模型取得了显著进展"
    assert c.para_index == 0
    assert "后文继续" in c.para_text
    assert c.resolved is False
    assert c.date_short == "2026-08-10"


def test_point_comment_without_range(docx_factory):
    body = p(r("某段文字。") + '<w:r><w:commentReference w:id="0"/></w:r>')
    comments = comment_xml("0", "李老师", "整段建议重写。")
    review = extract_review(docx_factory(body, comments))
    c = review.comments[0]
    assert c.quoted == ""
    assert c.para_text == "某段文字。"
    assert c.para_index == 0


def test_reply_threading_and_resolved(docx_factory):
    body = p(anchored("图3-1分辨率太低", "0") + '<w:r><w:commentReference w:id="1"/></w:r>')
    comments = comment_xml("0", "王老师", "请替换为矢量图。", para_id="P0") + comment_xml(
        "1", "小明", "已替换为 PDF 矢量图。", para_id="P1"
    )
    ex = '<w15:commentEx w15:paraId="P0" w15:done="1"/><w15:commentEx w15:paraId="P1" w15:paraIdParent="P0" w15:done="0"/>'
    review = extract_review(docx_factory(body, comments, ex))

    assert len(review.comments) == 1
    parent = review.comments[0]
    assert parent.resolved is True
    assert len(parent.replies) == 1
    assert parent.replies[0].author == "小明"
    assert review.total_comments == 2
    assert review.reply_count == 1


def test_missing_comments_extended_means_flat_unresolved(docx_factory):
    body = p(anchored("片段甲", "0")) + p(anchored("片段乙", "1"))
    comments = comment_xml("0", "A", "第一条", para_id="P0") + comment_xml("1", "B", "第二条", para_id="P1")
    review = extract_review(docx_factory(body, comments))
    assert len(review.comments) == 2
    assert all(not c.resolved for c in review.comments)


def test_comment_range_across_paragraphs(docx_factory):
    body = (
        "<w:p>" + f'<w:commentRangeStart w:id="0"/>' + r("前一段结尾") + "</w:p>"
        "<w:p>" + r("后一段开头") + f'<w:commentRangeEnd w:id="0"/>'
        '<w:r><w:commentReference w:id="0"/></w:r>' + "</w:p>"
    )
    comments = comment_xml("0", "王老师", "这两段衔接生硬。")
    review = extract_review(docx_factory(body, comments))
    assert review.comments[0].quoted == "前一段结尾 后一段开头"
    assert review.comments[0].para_index == 0


def test_multiparagraph_comment_body_joined_with_newline(docx_factory):
    body = p(anchored("正文", "0"))
    comments = (
        '<w:comment w:id="0" w:author="王老师" w:date="2026-08-10T02:00:00Z" w:initials="W">'
        "<w:p><w:r><w:t>第一点：结构混乱。</w:t></w:r></w:p>"
        '<w:p w14:paraId="PX"><w:r><w:t>第二点：缺少数据支撑。</w:t></w:r></w:p>'
        "</w:comment>"
    )
    review = extract_review(docx_factory(body, comments))
    assert review.comments[0].text == "第一点：结构混乱。\n第二点：缺少数据支撑。"


def test_comment_inside_table_cell(docx_factory):
    body = (
        p(r("表格前的段落。"))
        + "<w:tbl><w:tr><w:tc>"
        + p(anchored("表内数据", "0"))
        + "</w:tc></w:tr></w:tbl>"
    )
    comments = comment_xml("0", "王老师", "数据来源？")
    review = extract_review(docx_factory(body, comments))
    c = review.comments[0]
    assert c.quoted == "表内数据"
    assert c.para_index == 1


def test_insertion_revision(docx_factory):
    body = p(
        r("模型效果")
        + '<w:ins w:id="10" w:author="王老师" w:date="2026-08-11T03:00:00Z">'
        + r("显著")
        + "</w:ins>"
        + r("提升。")
    )
    review = extract_review(docx_factory(body))
    assert len(review.revisions) == 1
    rev = review.revisions[0]
    assert rev.kind == "insert"
    assert rev.author == "王老师"
    assert rev.text == "显著"
    assert rev.para_index == 0
    # inserted text is part of the final document
    assert review.paragraph_count == 1
    assert rev.para_text == "模型效果显著提升。"


def test_deletion_revision(docx_factory):
    body = p(
        r("结果")
        + '<w:del w:id="11" w:author="王老师" w:date="2026-08-11T03:00:00Z">'
        + '<w:r><w:delText>非常</w:delText></w:r>'
        + "</w:del>"
        + r("理想。")
    )
    review = extract_review(docx_factory(body))
    rev = review.revisions[0]
    assert rev.kind == "delete"
    assert rev.text == "非常"
    # deleted text is NOT part of the final document
    assert rev.para_text == "结果理想。"


def test_insert_and_delete_in_same_paragraph(docx_factory):
    body = p(
        '<w:ins w:id="1" w:author="A" w:date="2026-08-11T00:00:00Z">' + r("新增") + "</w:ins>"
        + r("保留")
        + '<w:del w:id="2" w:author="B" w:date="2026-08-11T00:00:00Z">'
        + "<w:r><w:delText>删掉</w:delText></w:r></w:del>"
    )
    review = extract_review(docx_factory(body))
    kinds = [rev.kind for rev in review.revisions]
    assert kinds == ["insert", "delete"]
    assert review.revisions[0].para_text == "新增保留"


def test_quoted_keeps_text_a_reviewer_deleted(docx_factory):
    # the comment was anchored on the original text; if a tracked deletion
    # strikes part of it through, the quoted excerpt must not lose that part
    body = p(
        f'<w:commentRangeStart w:id="0"/>'
        + r("保留的")
        + '<w:del w:id="5" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + "<w:r><w:delText>被删的</w:delText></w:r></w:del>"
        + r("文字")
        + f'<w:commentRangeEnd w:id="0"/>'
        + '<w:r><w:commentReference w:id="0"/></w:r>'
    )
    comments = comment_xml("0", "王老师", "措辞再斟酌。")
    review = extract_review(docx_factory(body, comments))
    assert review.comments[0].quoted == "保留的被删的文字"
    # the final paragraph text still excludes the deletion
    assert review.comments[0].para_text == "保留的文字"


def test_comments_sorted_by_document_position(docx_factory):
    body = p(anchored("先出现", "7")) + p(anchored("后出现", "3"))
    # comments.xml lists them in the opposite order
    comments = comment_xml("3", "B", "第二条", para_id="P3") + comment_xml("7", "A", "第一条", para_id="P7")
    review = extract_review(docx_factory(body, comments))
    assert [c.id for c in review.comments] == ["7", "3"]


def test_tab_and_break_become_space(docx_factory):
    body = p("<w:r><w:t>甲</w:t><w:tab/><w:t>乙</w:t><w:br/><w:t>丙</w:t></w:r>")
    review = extract_review(docx_factory(body))
    assert review.paragraph_count == 1
    # spaces collapse to single separators
    path2 = docx_factory(body + p(anchored("锚点", "0")), comment_xml("0", "A", "查看段文本"))
    review2 = extract_review(path2)
    assert review2.comments[0].para_text == "锚点"


def test_move_reported_once_at_new_location(docx_factory):
    body = p(
        '<w:moveTo w:id="20" w:author="A" w:date="2026-08-11T00:00:00Z">' + r("移动来的") + "</w:moveTo>"
        + r("正文")
        + '<w:moveFrom w:id="21" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + "<w:r><w:delText>移走的</w:delText></w:r></w:moveFrom>"
    )
    review = extract_review(docx_factory(body))
    # moveTo listed once as a "move"; moveFrom (old location) not repeated
    assert [rev.kind for rev in review.revisions] == ["move"]
    assert review.revisions[0].text == "移动来的"
    # final text keeps moveTo content and drops moveFrom content
    assert review.comments == []
    path2 = docx_factory(
        "<w:p>"
        + f'<w:commentRangeStart w:id="0"/>'
        + '<w:moveTo w:id="20" w:author="A" w:date="2026-08-11T00:00:00Z">'
        + r("移动来的")
        + "</w:moveTo>"
        + r("正文")
        + f'<w:commentRangeEnd w:id="0"/>'
        + '<w:r><w:commentReference w:id="0"/></w:r>'
        + "</w:p>",
        comment_xml("0", "A", "检查可见性"),
    )
    review2 = extract_review(path2)
    assert review2.comments[0].quoted == "移动来的正文"


def test_not_a_zip_raises(tmp_path):
    bogus = tmp_path / "fake.docx"
    bogus.write_text("这不是一个 zip 文件", encoding="utf-8")
    with pytest.raises(DocxReviewError, match="不是有效的 .docx"):
        extract_review(bogus)


def test_zip_without_document_xml_raises(tmp_path):
    import zipfile

    path = tmp_path / "empty.docx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("hello.txt", "hi")
    with pytest.raises(DocxReviewError, match="缺少 word/document.xml"):
        extract_review(path)


def test_missing_file_raises(tmp_path):
    with pytest.raises(DocxReviewError, match="文件不存在"):
        extract_review(tmp_path / "nope.docx")
