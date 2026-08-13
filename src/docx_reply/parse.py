"""Read review data (comments, replies, tracked changes) out of a .docx file.

A .docx file is a zip archive of XML parts. Everything the review workflow
needs lives in three of them:

- ``word/document.xml``          body text, comment anchors, tracked changes
- ``word/comments.xml``          comment text, author, date
- ``word/commentsExtended.xml``  reply threading and "resolved" state

Only the standard library is used (``zipfile`` + ``xml.etree``), so the
package has zero dependencies and works the same on Windows / macOS / Linux
with documents produced by Microsoft Word or WPS Office.
"""

from __future__ import annotations

import itertools
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from .models import Comment, Review, Revision

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


_BODY = _q(_W, "body")
_P = _q(_W, "p")
_R = _q(_W, "r")
_T = _q(_W, "t")
_DELTEXT = _q(_W, "delText")
_TAB = _q(_W, "tab")
_BR = _q(_W, "br")
_CR = _q(_W, "cr")
_INS = _q(_W, "ins")
_DEL = _q(_W, "del")
_MOVE_FROM = _q(_W, "moveFrom")
_MOVE_TO = _q(_W, "moveTo")
_RANGE_START = _q(_W, "commentRangeStart")
_RANGE_END = _q(_W, "commentRangeEnd")
_REFERENCE = _q(_W, "commentReference")
_COMMENT = _q(_W, "comment")
_ID = _q(_W, "id")
_AUTHOR = _q(_W, "author")
_DATE = _q(_W, "date")
_INITIALS = _q(_W, "initials")
_PARA_ID = _q(_W14, "paraId")
_EX = _q(_W15, "commentEx")
_EX_PARA = _q(_W15, "paraId")
_EX_PARENT = _q(_W15, "paraIdParent")
_EX_DONE = _q(_W15, "done")
_PPR = _q(_W, "pPr")
_PSTYLE = _q(_W, "pStyle")
_VAL = _q(_W, "val")

_WS_RE = re.compile(r"\s+")

# Style ids Word/WPS use for built-in heading styles: "Heading1", "1"
# (Chinese Word), "标题 1". Anything else is treated as body text.
_HEADING_STYLE_RE = re.compile(r"^(?:heading\s*[1-9]|标题\s*[1-9]|[1-9])$", re.IGNORECASE)


class DocxReviewError(Exception):
    """Raised when the input cannot be read as a .docx document."""


def _collapse(text: str) -> str:
    return _WS_RE.sub(" ", text).strip()


def _run_text(run: ET.Element) -> str:
    parts: List[str] = []
    for child in run:
        if child.tag in (_T, _DELTEXT):
            parts.append(child.text or "")
        elif child.tag in (_TAB, _BR, _CR):
            parts.append(" ")
    return "".join(parts)


class _DocumentData:
    """Raw facts collected in a single walk over document.xml."""

    def __init__(self) -> None:
        self.paragraphs: List[str] = []
        self.quoted: Dict[str, str] = {}
        self.anchor_para: Dict[str, Optional[int]] = {}
        self.anchor_order: Dict[str, int] = {}
        self.revisions: List[Revision] = []
        # (paragraph index, heading text) in document order
        self.headings: List[Tuple[int, str]] = []
        # revision -> (start, end) positions of the text-run counter,
        # used to detect truly adjacent del+ins pairs (replacements)
        self.rev_spans: List[Tuple[int, int]] = []

    def nearest_heading(self, idx: Optional[int]) -> str:
        if idx is None:
            return ""
        best = ""
        for h_idx, h_text in self.headings:
            if h_idx <= idx:
                best = h_text
            else:
                break
        return best


def _walk_document(root: ET.Element) -> _DocumentData:
    body = root.find(_BODY)
    if body is None:
        raise DocxReviewError("document.xml 中没有找到文档主体（w:body）")

    data = _DocumentData()
    # comment id -> text pieces accumulated while its range is open;
    # ranges may span paragraphs, so this lives outside the paragraph scope
    open_ranges: Dict[str, List[str]] = {}
    order = itertools.count()
    cur_idx: Optional[int] = None
    cur_texts: Optional[List[str]] = None
    # counts text-bearing runs; two revisions with touching spans have no
    # document text between them and can form a replacement pair
    run_clock = [0]

    def anchor(cid: str) -> None:
        data.anchor_para.setdefault(cid, cur_idx)
        data.anchor_order.setdefault(cid, next(order))

    def paragraph_style(el: ET.Element) -> str:
        ppr = el.find(_PPR)
        if ppr is not None:
            pstyle = ppr.find(_PSTYLE)
            if pstyle is not None:
                return pstyle.get(_VAL, "")
        return ""

    def visit(el: ET.Element, in_del: bool, rev_buf: Optional[List[str]]) -> None:
        nonlocal cur_idx, cur_texts
        tag = el.tag
        if tag == _P:
            prev_idx, prev_texts = cur_idx, cur_texts
            cur_idx = len(data.paragraphs)
            data.paragraphs.append("")
            cur_texts = []
            for child in el:
                visit(child, in_del, rev_buf)
            text = _collapse("".join(cur_texts))
            data.paragraphs[cur_idx] = text
            if text and _HEADING_STYLE_RE.match(paragraph_style(el)):
                data.headings.append((cur_idx, text))
            for buf in open_ranges.values():
                buf.append(" ")
            cur_idx, cur_texts = prev_idx, prev_texts
        elif tag == _RANGE_START:
            cid = el.get(_ID, "")
            open_ranges[cid] = []
            anchor(cid)
        elif tag == _RANGE_END:
            cid = el.get(_ID, "")
            buf = open_ranges.pop(cid, None)
            if buf is not None:
                data.quoted[cid] = _collapse("".join(buf))
        elif tag in (_INS, _MOVE_TO):
            buf: List[str] = []
            start = run_clock[0]
            for child in el:
                visit(child, in_del, buf)
            text = _collapse("".join(buf))
            if tag == _INS and text and not in_del:
                data.revisions.append(
                    Revision(
                        kind="insert",
                        author=el.get(_AUTHOR, ""),
                        date=el.get(_DATE, ""),
                        text=text,
                        para_index=cur_idx,
                    )
                )
                data.rev_spans.append((start, run_clock[0]))
        elif tag in (_DEL, _MOVE_FROM):
            del_buf: List[str] = []
            start = run_clock[0]
            for child in el:
                visit(child, True, del_buf)
            text = _collapse("".join(del_buf))
            if tag == _DEL and text:
                data.revisions.append(
                    Revision(
                        kind="delete",
                        author=el.get(_AUTHOR, ""),
                        date=el.get(_DATE, ""),
                        text=text,
                        para_index=cur_idx,
                    )
                )
                data.rev_spans.append((start, run_clock[0]))
        elif tag == _R:
            for child in el:
                if child.tag == _REFERENCE:
                    anchor(child.get(_ID, ""))
            text = _run_text(el)
            if not text:
                return
            run_clock[0] += 1
            if rev_buf is not None:
                rev_buf.append(text)
            if not in_del:
                # visible in the final document: plain text and insertions
                if cur_texts is not None:
                    cur_texts.append(text)
                for buf in open_ranges.values():
                    buf.append(text)
        else:
            for child in el:
                visit(child, in_del, rev_buf)

    visit(body, False, None)
    data.headings.sort(key=lambda h: h[0])
    return data


def _merge_replacements(revisions: List[Revision], spans: List[Tuple[int, int]]) -> List[Revision]:
    """Fold a del+ins pair with no text between them into one "replace".

    Selecting text in Word and typing over it records exactly this pair
    (in either order); presenting it as ``old → new`` is far more readable
    in a response table than two disconnected rows.
    """
    merged: List[Revision] = []
    i = 0
    while i < len(revisions):
        cur = revisions[i]
        if i + 1 < len(revisions):
            nxt = revisions[i + 1]
            adjacent = spans[i][1] == spans[i + 1][0]
            complementary = {cur.kind, nxt.kind} == {"insert", "delete"}
            if (
                adjacent
                and complementary
                and cur.author == nxt.author
                and cur.para_index == nxt.para_index
            ):
                deleted = cur.text if cur.kind == "delete" else nxt.text
                inserted = cur.text if cur.kind == "insert" else nxt.text
                merged.append(
                    Revision(
                        kind="replace",
                        author=cur.author,
                        date=cur.date,
                        text=f"{deleted} → {inserted}",
                        para_index=cur.para_index,
                        deleted=deleted,
                        inserted=inserted,
                    )
                )
                i += 2
                continue
        merged.append(cur)
        i += 1
    return merged


def _parse_comments(root: ET.Element) -> List[Comment]:
    comments: List[Comment] = []
    for c in root.findall(_COMMENT):
        pieces: List[str] = []
        last_para_id: Optional[str] = None
        for p in c.findall(f".//{_P}"):
            texts = [n.text or "" for n in p.iter() if n.tag in (_T, _DELTEXT)]
            piece = _collapse("".join(texts))
            if piece:
                pieces.append(piece)
            pid = p.get(_PARA_ID)
            if pid:
                last_para_id = pid
        comments.append(
            Comment(
                id=c.get(_ID, ""),
                author=c.get(_AUTHOR, ""),
                initials=c.get(_INITIALS, ""),
                date=c.get(_DATE, ""),
                text="\n".join(pieces),
                para_id=last_para_id,
            )
        )
    return comments


def _parse_extended(root: ET.Element) -> Dict[str, Tuple[Optional[str], bool]]:
    """paraId -> (paraIdParent, done)."""
    out: Dict[str, Tuple[Optional[str], bool]] = {}
    for ce in root.findall(_EX):
        pid = ce.get(_EX_PARA)
        if pid:
            out[pid] = (ce.get(_EX_PARENT), ce.get(_EX_DONE) == "1")
    return out


def _read_part(zf: zipfile.ZipFile, name: str) -> Optional[ET.Element]:
    try:
        payload = zf.read(name)
    except KeyError:
        return None
    try:
        return ET.fromstring(payload)
    except ET.ParseError as exc:
        raise DocxReviewError(f"无法解析 {name}：{exc}") from exc


def extract_review(path: "str | Path") -> Review:
    """Extract comments, replies, resolved state and tracked changes.

    Raises :class:`DocxReviewError` if the file is missing or is not a
    valid .docx document.
    """
    p = Path(path)
    if not p.is_file():
        raise DocxReviewError(f"文件不存在：{p}")
    try:
        zf = zipfile.ZipFile(p)
    except zipfile.BadZipFile as exc:
        raise DocxReviewError(
            f"{p.name} 不是有效的 .docx 文件（无法作为 zip 打开；.doc 老格式请先另存为 .docx）"
        ) from exc
    with zf:
        doc_root = _read_part(zf, "word/document.xml")
        if doc_root is None:
            raise DocxReviewError(f"{p.name} 不是有效的 .docx 文件（缺少 word/document.xml）")
        data = _walk_document(doc_root)
        comments_root = _read_part(zf, "word/comments.xml")
        ext_root = _read_part(zf, "word/commentsExtended.xml")

    all_comments = _parse_comments(comments_root) if comments_root is not None else []
    ext = _parse_extended(ext_root) if ext_root is not None else {}

    def para_text(idx: Optional[int]) -> str:
        if idx is not None and 0 <= idx < len(data.paragraphs):
            return data.paragraphs[idx]
        return ""

    by_para_id = {c.para_id: c for c in all_comments if c.para_id}
    top: List[Comment] = []
    for c in all_comments:
        c.quoted = data.quoted.get(c.id, "")
        c.para_index = data.anchor_para.get(c.id)
        c.para_text = para_text(c.para_index)
        c.heading = data.nearest_heading(c.para_index)
        parent_pid, done = ext.get(c.para_id, (None, False)) if c.para_id else (None, False)
        c.resolved = done
        parent = by_para_id.get(parent_pid) if parent_pid else None
        if parent is not None and parent is not c:
            parent.replies.append(c)
        else:
            top.append(c)

    big = 10**9
    top.sort(key=lambda c: data.anchor_order.get(c.id, big))

    revisions = _merge_replacements(data.revisions, data.rev_spans)
    for rev in revisions:
        rev.para_text = para_text(rev.para_index)
        rev.heading = data.nearest_heading(rev.para_index)

    return Review(
        source=p.name,
        paragraph_count=len(data.paragraphs),
        comments=top,
        revisions=revisions,
    )
