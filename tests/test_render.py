import csv
import io
import json

from docx_reply import render_comments_csv, render_json, render_markdown, render_revisions_csv
from docx_reply.models import Comment, Review, Revision


def sample_review() -> Review:
    reply = Comment(id="1", author="小明", initials="M", date="2026-08-12T01:00:00Z", text="已补充引用[15]-[17]。")
    c1 = Comment(
        id="0",
        author="王老师",
        initials="W",
        date="2026-08-10T02:00:00Z",
        text="这一段缺少对相关工作的引用，请补充 2-3 篇近三年文献。",
        quoted="深度学习模型在图像识别领域取得了显著进展",
        para_index=1,
        para_text="深度学习模型在图像识别领域取得了显著进展，尤其是卷积神经网络。",
        replies=[reply],
    )
    c2 = Comment(
        id="2",
        author="王老师",
        initials="W",
        date="2026-08-10T02:05:00Z",
        text="图3-1 的分辨率太低，含 | 竖线与\n换行。",
        quoted="图3-1",
        para_index=3,
        para_text="实验结果如图3-1所示。",
        resolved=True,
    )
    rev = Revision(
        kind="insert",
        author="王老师",
        date="2026-08-11T03:00:00Z",
        text="显著",
        para_index=2,
        para_text="模型效果显著提升。",
    )
    return Review(source="thesis.docx", paragraph_count=5, comments=[c1, c2], revisions=[rev])


def test_markdown_structure():
    md = render_markdown(sample_review())
    assert md.startswith("# 审阅意见对照表：thesis.docx")
    assert "批注 3 条（含回复 1 条）" in md
    assert "修订 1 处" in md
    assert "| 序号 | 位置 | 原文摘录 | 批注人 | 日期 | 批注内容 | 回复 | 状态 | 修改说明 |" in md
    assert "| 1 | 第2段 |" in md
    assert "小明：已补充引用[15]-[17]。" in md
    assert "已解决" in md and "未解决" in md
    assert "## 修订记录" in md
    assert "| 1 | 第3段 | 插入 | 王老师 | 2026-08-11 | 显著 |" in md


def test_markdown_escapes_pipes_and_newlines():
    md = render_markdown(sample_review())
    assert "\\|" in md
    assert "含 \\| 竖线与<br>换行。" in md


def test_markdown_without_comments():
    review = Review(source="clean.docx", paragraph_count=3)
    md = render_markdown(review)
    assert "未发现批注。" in md
    assert "## 修订记录" not in md


def test_markdown_can_hide_revisions():
    md = render_markdown(sample_review(), include_revisions=False)
    assert "## 修订记录" not in md


def test_excerpt_truncation():
    long_quote = "很长的原文" * 30
    review = Review(
        source="x.docx",
        paragraph_count=1,
        comments=[Comment(id="0", author="A", initials="A", date="", text="评", quoted=long_quote, para_index=0)],
    )
    md = render_markdown(review)
    assert "…" in md
    assert long_quote not in md


def test_comments_csv_roundtrip():
    rows = list(csv.reader(io.StringIO(render_comments_csv(sample_review()))))
    assert rows[0] == ["序号", "位置", "原文摘录", "批注人", "日期", "批注内容", "回复", "状态", "修改说明"]
    assert len(rows) == 3
    assert rows[1][1] == "第2段"
    assert rows[1][6] == "小明：已补充引用[15]-[17]。"
    assert rows[2][7] == "已解决"
    # 修改说明 column is left blank for the student to fill in
    assert rows[1][8] == "" and rows[2][8] == ""


def test_revisions_csv():
    rows = list(csv.reader(io.StringIO(render_revisions_csv(sample_review()))))
    assert rows[0] == ["序号", "位置", "类型", "作者", "日期", "内容"]
    assert rows[1] == ["1", "第3段", "插入", "王老师", "2026-08-11", "显著"]


def test_json_structure():
    payload = json.loads(render_json(sample_review()))
    assert payload["source"] == "thesis.docx"
    assert payload["stats"] == {
        "comments": 3,
        "replies": 1,
        "resolved": 1,
        "revisions": 1,
        "authors": ["王老师", "小明"],
    }
    assert payload["comments"][0]["paragraph"] == 2
    assert payload["comments"][0]["replies"][0]["author"] == "小明"
    assert payload["revisions"][0]["kind"] == "insert"


def test_resolved_count_includes_nested_replies():
    reply = Comment(
        id="1",
        author="小明",
        initials="M",
        date="",
        text="已改",
        resolved=True,
    )
    parent = Comment(
        id="0",
        author="王老师",
        initials="W",
        date="",
        text="请改",
        replies=[reply],
    )
    review = Review(source="x.docx", comments=[parent])
    assert review.resolved_count == 1
    assert review.total_comments == 2
