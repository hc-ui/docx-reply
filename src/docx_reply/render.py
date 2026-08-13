"""Render a Review as a Markdown / Word response table, CSV or JSON."""

from __future__ import annotations

import csv
import io
import json
import zipfile
from xml.sax.saxutils import escape as _xml_escape

from .models import Review

_KIND_ZH = {"insert": "插入", "delete": "删除", "replace": "替换", "move": "移动", "format": "格式"}

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


def _loc(item) -> str:
    """Location of a comment/revision: ``[脚注 ·] 2.1 实验设计 · 第14段``."""
    if item.para_index is None:
        return ""
    pieces = []
    if getattr(item, "part", ""):
        pieces.append(item.part)
    heading = getattr(item, "heading", "")
    if heading:
        pieces.append(_excerpt(heading, 24))
    pieces.append(f"第{item.para_index + 1}段")
    return " · ".join(pieces)


def _status(resolved: bool) -> str:
    return "已解决" if resolved else "未解决"


def _replies_text(comment, sep: str) -> str:
    return sep.join(f"{r.author}：{r.text}" for r in comment.replies)


def _summary(review: Review) -> str:
    parts = [f"批注 {review.total_comments} 条"]
    if review.reply_count:
        parts[-1] += f"（含回复 {review.reply_count} 条）"
    parts.append(f"修订 {len(review.revisions)} 处")
    if review.authors:
        parts.append("审阅人：" + "、".join(review.authors))
    return " ｜ ".join(parts)


def render_markdown(review: Review, include_revisions: bool = True) -> str:
    lines = [f"# 审阅意见对照表：{review.source}", ""]
    lines.append(_summary(review))
    lines.append("")

    if review.comments:
        lines.append("## 批注（修改对照表）")
        lines.append("")
        lines.append("| " + " | ".join(COMMENT_HEADERS) + " |")
        lines.append("| ---: | " + " | ".join("---" for _ in COMMENT_HEADERS[1:]) + " |")
        for i, c in enumerate(review.comments, 1):
            row = [
                str(i),
                _cell(_loc(c)),
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
                _cell(_loc(rev)),
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
                _loc(c),
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
            [i, _loc(rev), _KIND_ZH.get(rev.kind, rev.kind), rev.author, rev.date_short, rev.text]
        )
    return buf.getvalue()


def render_json(review: Review) -> str:
    return json.dumps(review.to_dict(), ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# .docx output: the response table as a Word document, ready to hand in
# ---------------------------------------------------------------------------

_DOCX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

_DOCX_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def _docx_p(text: str = "", bold: bool = False, size: "int | None" = None) -> str:
    if not text:
        return "<w:p/>"
    rpr = ""
    if bold or size:
        rpr = (
            "<w:rPr>"
            + ("<w:b/>" if bold else "")
            + (f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>' if size else "")
            + "</w:rPr>"
        )
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{_xml_escape(text)}</w:t></w:r></w:p>'


def _docx_cell(lines: "list[str]", bold: bool = False) -> str:
    paragraphs = "".join(_docx_p(line, bold=bold) for line in lines) or "<w:p/>"
    return f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>{paragraphs}</w:tc>'


def _docx_table(headers: "list[str]", rows: "list[list[list[str]]]") -> str:
    borders = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:color="auto"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    parts = [
        "<w:tbl><w:tblPr>"
        '<w:tblW w:w="5000" w:type="pct"/>'
        f"<w:tblBorders>{borders}</w:tblBorders>"
        "</w:tblPr>"
    ]
    parts.append("<w:tr>" + "".join(_docx_cell([h], bold=True) for h in headers) + "</w:tr>")
    for row in rows:
        parts.append("<w:tr>" + "".join(_docx_cell(cell) for cell in row) + "</w:tr>")
    parts.append("</w:tbl>")
    return "".join(parts)


def render_docx(review: Review, include_revisions: bool = True) -> bytes:
    """Build a .docx response table (修改对照表) ready to fill in and submit."""
    body: "list[str]" = [
        _docx_p(f"审阅意见对照表：{review.source}", bold=True, size=32),
        _docx_p(_summary(review)),
        _docx_p(),
    ]

    if review.comments:
        rows = []
        for i, c in enumerate(review.comments, 1):
            replies = [f"{r.author}：{r.text}" for r in c.replies]
            rows.append(
                [
                    [str(i)],
                    [_loc(c)],
                    [_excerpt(c.quoted or c.para_text)],
                    [c.author],
                    [c.date_short],
                    c.text.split("\n"),
                    replies,
                    [_status(c.resolved)],
                    [],
                ]
            )
        body.append(_docx_table(COMMENT_HEADERS, rows))
    else:
        body.append(_docx_p("未发现批注。"))

    if include_revisions and review.revisions:
        body.append(_docx_p())
        body.append(_docx_p("修订记录", bold=True, size=28))
        rows = [
            [
                [str(i)],
                [_loc(rev)],
                [_KIND_ZH.get(rev.kind, rev.kind)],
                [rev.author],
                [rev.date_short],
                [rev.text],
            ]
            for i, rev in enumerate(review.revisions, 1)
        ]
        body.append(_docx_table(REVISION_HEADERS, rows))

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(body)}</w:body></w:document>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _DOCX_CONTENT_TYPES)
        zf.writestr("_rels/.rels", _DOCX_RELS)
        zf.writestr("word/document.xml", document)
    return buf.getvalue()
