"""Render a Review as a Markdown response table, CSV or JSON."""

from __future__ import annotations

import csv
import io
import json
from typing import Optional

from .models import Review

_KIND_ZH = {"insert": "插入", "delete": "删除"}

COMMENT_HEADERS = ["序号", "位置", "原文摘录", "批注人", "日期", "批注内容", "回复", "状态", "修改说明"]
REVISION_HEADERS = ["序号", "位置", "类型", "作者", "日期", "内容"]


def _excerpt(text: str, limit: int = 60) -> str:
    text = text.strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "…"


def _cell(text: str) -> str:
    """Escape a value for use inside a Markdown table cell."""
    return text.replace("|", "\\|").replace("\n", "<br>")


def _loc(para_index: Optional[int]) -> str:
    return "" if para_index is None else f"第{para_index + 1}段"


def _status(resolved: bool) -> str:
    return "已解决" if resolved else "未解决"


def _replies_text(comment, sep: str) -> str:
    return sep.join(f"{r.author}：{r.text}" for r in comment.replies)


def render_markdown(review: Review, include_revisions: bool = True) -> str:
    lines = [f"# 审阅意见对照表：{review.source}", ""]
    summary = [f"批注 {review.total_comments} 条"]
    if review.reply_count:
        summary[-1] += f"（含回复 {review.reply_count} 条）"
    summary.append(f"修订 {len(review.revisions)} 处")
    if review.authors:
        summary.append("审阅人：" + "、".join(review.authors))
    lines.append(" ｜ ".join(summary))
    lines.append("")

    if review.comments:
        lines.append("## 批注（修改对照表）")
        lines.append("")
        lines.append("| " + " | ".join(COMMENT_HEADERS) + " |")
        lines.append("| ---: | " + " | ".join("---" for _ in COMMENT_HEADERS[1:]) + " |")
        for i, c in enumerate(review.comments, 1):
            row = [
                str(i),
                _loc(c.para_index),
                _cell(_excerpt(c.quoted or c.para_text)),
                _cell(c.author),
                c.date_short,
                _cell(c.text),
                _cell(_replies_text(c, "\n")),
                _status(c.resolved),
                "",
            ]
            lines.append("| " + " | ".join(row) + " |")
        lines.append("")
    else:
        lines.append("未发现批注。")
        lines.append("")

    if include_revisions and review.revisions:
        lines.append("## 修订记录")
        lines.append("")
        lines.append("| " + " | ".join(REVISION_HEADERS) + " |")
        lines.append("| ---: | " + " | ".join("---" for _ in REVISION_HEADERS[1:]) + " |")
        for i, rev in enumerate(review.revisions, 1):
            row = [
                str(i),
                _loc(rev.para_index),
                _KIND_ZH.get(rev.kind, rev.kind),
                _cell(rev.author),
                rev.date_short,
                _cell(_excerpt(rev.text, 80)),
            ]
            lines.append("| " + " | ".join(row) + " |")
        lines.append("")

    return "\n".join(lines)


def render_comments_csv(review: Review) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerow(COMMENT_HEADERS)
    for i, c in enumerate(review.comments, 1):
        writer.writerow(
            [
                i,
                _loc(c.para_index),
                _excerpt(c.quoted or c.para_text),
                c.author,
                c.date_short,
                c.text,
                _replies_text(c, "\n"),
                _status(c.resolved),
                "",
            ]
        )
    return buf.getvalue()


def render_revisions_csv(review: Review) -> str:
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerow(REVISION_HEADERS)
    for i, rev in enumerate(review.revisions, 1):
        writer.writerow(
            [i, _loc(rev.para_index), _KIND_ZH.get(rev.kind, rev.kind), rev.author, rev.date_short, rev.text]
        )
    return buf.getvalue()


def render_json(review: Review) -> str:
    return json.dumps(review.to_dict(), ensure_ascii=False, indent=2)
