"""Read review data (comments, replies, tracked changes) out of a .docx file.

A .docx file is a zip archive of XML parts. Everything the review workflow
needs lives in a handful of them:

- ``word/document.xml``          body text, comment anchors, tracked changes
- ``word/footnotes.xml`` / ``word/endnotes.xml`` / ``word/header*.xml`` /
  ``word/footer*.xml``           the same, for the other document stories
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
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


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
_NBHYPHEN = _q(_W, "noBreakHyphen")
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
_RPR = _q(_W, "rPr")
_RPRCHANGE = _q(_W, "rPrChange")
_PPRCHANGE = _q(_W, "pPrChange")
_DRAWING = _q(_W, "drawing")
_PICT = _q(_W, "pict")
_OBJECT = _q(_W, "object")
_RUBY = _q(_W, "ruby")
_RT = _q(_W, "rt")
_ALTERNATE = _q(_MC, "AlternateContent")
_MC_CHOICE = _q(_MC, "Choice")
_MC_FALLBACK = _q(_MC, "Fallback")
_M_R = _q(_M, "r")
_M_T = _q(_M, "t")

# Run children that may hold nested document content (text boxes, ruby text)
_RUN_CONTAINERS = (_DRAWING, _PICT, _OBJECT, _ALTERNATE, _RUBY)

_WS_RE = re.compile(r"\s+")

# Style ids Word/WPS use for built-in heading styles: "Heading1", "1"
# (Chinese Word), "标题 1". Anything else is treated as body text.
_HEADING_STYLE_RE = re.compile(r"^(?:heading\s*[1-9]|标题\s*[1-9]|[1-9])$", re.IGNORECASE)

# Document stories beyond the main body, in scan order.
_EXTRA_PART_RE = re.compile(r"^word/(footnotes|endnotes|header\d*|footer\d*)\.xml$")
_PART_LABELS = (
    ("footnotes", "脚注"),
    ("endnotes", "尾注"),
    ("header", "页眉"),
    ("footer", "页脚"),
)


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
        elif child.tag == _NBHYPHEN:
            parts.append("-")
    return "".join(parts)


class _PartData:
    """Paragraphs and headings of one document story (body, footnotes, …)."""

    def __init__(self, label: str) -> None:
        self.label = label
        self.paragraphs: List[str] = []
        self.headings: List[Tuple[int, str]] = []

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

    def text_at(self, idx: Optional[int]) -> str:
        if idx is not None and 0 <= idx < len(self.paragraphs):
            return self.paragraphs[idx]
        return ""


class _WalkState:
    """Facts accumulated across all document stories."""

    def __init__(self) -> None:
        self.parts: Dict[str, _PartData] = {}
        self.quoted: Dict[str, str] = {}
        # comment id -> (part label, paragraph index within the part)
        self.anchor: Dict[str, Tuple[str, Optional[int]]] = {}
        self.anchor_order: Dict[str, int] = {}
        self.order = itertools.count()
        self.revisions: List[Revision] = []
        # revision -> (start, end) positions of the text-run counter, used
        # to detect truly adjacent del+ins pairs (replacements)
        self.rev_spans: List[Tuple[int, int]] = []
        self.run_clock = [0]

    def add_revision(self, rev: Revision, span: Tuple[int, int]) -> None:
        self.revisions.append(rev)
        self.rev_spans.append(span)


def _walk_part(container: ET.Element, label: str, state: _WalkState) -> None:
    pd = state.parts.setdefault(label, _PartData(label))
    # comment id -> text pieces accumulated while its range is open;
    # ranges may span paragraphs, so this lives outside the paragraph scope
    open_ranges: Dict[str, List[str]] = {}
    cur_idx: Optional[int] = None
    cur_texts: Optional[List[str]] = None
    clock = state.run_clock

    def anchor(cid: str) -> None:
        state.anchor.setdefault(cid, (label, cur_idx))
        state.anchor_order.setdefault(cid, next(state.order))

    def paragraph_child(el: ET.Element, parent_tag: str, child_tag: str) -> Optional[ET.Element]:
        parent = el.find(parent_tag)
        return parent.find(child_tag) if parent is not None else None

    def visit(el: ET.Element, in_del: bool, rev_buf: Optional[List[str]]) -> None:
        nonlocal cur_idx, cur_texts
        tag = el.tag
        if tag == _P:
            prev_idx, prev_texts = cur_idx, cur_texts
            cur_idx = len(pd.paragraphs)
            pd.paragraphs.append("")
            cur_texts = []
            for child in el:
                visit(child, in_del, rev_buf)
            text = _collapse("".join(cur_texts))
            pd.paragraphs[cur_idx] = text
            style_el = paragraph_child(el, _PPR, _PSTYLE)
            if text and style_el is not None and _HEADING_STYLE_RE.match(style_el.get(_VAL, "")):
                pd.headings.append((cur_idx, text))
            pch = paragraph_child(el, _PPR, _PPRCHANGE)
            if pch is not None and text:
                # sentinel span: paragraph-level format changes must never
                # be folded together with run-level ones
                state.add_revision(
                    Revision(
                        kind="format",
                        author=pch.get(_AUTHOR, ""),
                        date=pch.get(_DATE, ""),
                        text=text,
                        para_index=cur_idx,
                        part=label,
                    ),
                    (-1, -1),
                )
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
                state.quoted[cid] = _collapse("".join(buf))
        elif tag in (_INS, _MOVE_TO):
            ins_buf: List[str] = []
            start = clock[0]
            for child in el:
                visit(child, in_del, ins_buf)
            text = _collapse("".join(ins_buf))
            if text and not in_del:
                state.add_revision(
                    Revision(
                        kind="insert" if tag == _INS else "move",
                        author=el.get(_AUTHOR, ""),
                        date=el.get(_DATE, ""),
                        text=text,
                        para_index=cur_idx,
                        part=label,
                    ),
                    (start, clock[0]),
                )
        elif tag in (_DEL, _MOVE_FROM):
            del_buf: List[str] = []
            start = clock[0]
            for child in el:
                visit(child, True, del_buf)
            text = _collapse("".join(del_buf))
            # moveFrom is the old location of moved text: not visible in the
            # final document and already reported once as a "move" revision
            if tag == _DEL and text:
                state.add_revision(
                    Revision(
                        kind="delete",
                        author=el.get(_AUTHOR, ""),
                        date=el.get(_DATE, ""),
                        text=text,
                        para_index=cur_idx,
                        part=label,
                    ),
                    (start, clock[0]),
                )
        elif tag == _R:
            for child in el:
                if child.tag == _REFERENCE:
                    anchor(child.get(_ID, ""))
                elif child.tag in _RUN_CONTAINERS:
                    # text boxes / ruby carry real paragraphs inside the run
                    visit(child, in_del, rev_buf)
            text = _run_text(el)
            if not text:
                return
            clock[0] += 1
            if rev_buf is not None:
                rev_buf.append(text)
            # quoted ranges reflect what the comment was anchored on, so
            # they keep text that a tracked deletion has struck through
            for buf in open_ranges.values():
                buf.append(text)
            if not in_del:
                # visible in the final document: plain text and insertions
                if cur_texts is not None:
                    cur_texts.append(text)
                if rev_buf is None:
                    # plain run: surface formatting-only revisions too
                    rch = paragraph_child(el, _RPR, _RPRCHANGE)
                    if rch is not None:
                        state.add_revision(
                            Revision(
                                kind="format",
                                author=rch.get(_AUTHOR, ""),
                                date=rch.get(_DATE, ""),
                                text=text,
                                para_index=cur_idx,
                                part=label,
                            ),
                            (clock[0] - 1, clock[0]),
                        )
        elif tag == _M_R:
            # math run: linearize m:t fragments so formulas survive
            text = "".join(n.text or "" for n in el.iter(_M_T))
            if not text:
                return
            clock[0] += 1
            if rev_buf is not None:
                rev_buf.append(text)
            for buf in open_ranges.values():
                buf.append(text)
            if not in_del and cur_texts is not None:
                cur_texts.append(text)
        elif tag == _ALTERNATE:
            # mc:Choice and mc:Fallback duplicate the same content
            # (e.g. a text box as DrawingML and as legacy VML): walk one
            target = el.find(_MC_CHOICE)
            if target is None:
                target = el.find(_MC_FALLBACK)
            if target is not None:
                for child in target:
                    visit(child, in_del, rev_buf)
        elif tag == _RT:
            # phonetic guide annotation over ruby base text, not content
            return
        else:
            for child in el:
                visit(child, in_del, rev_buf)

    visit(container, False, None)
    # a range start without a matching end (malformed producer): best-effort
    # quoted text up to the end of the story instead of silently nothing
    for cid, buf in open_ranges.items():
        state.quoted.setdefault(cid, _collapse("".join(buf)))
    pd.headings.sort(key=lambda h: h[0])


def _merge_replacements(
    revisions: List[Revision], spans: List[Tuple[int, int]]
) -> Tuple[List[Revision], List[Tuple[int, int]]]:
    """Fold a del+ins pair with no text between them into one "replace".

    Selecting text in Word and typing over it records exactly this pair
    (in either order); presenting it as ``old → new`` is far more readable
    in a response table than two disconnected rows.
    """
    merged: List[Revision] = []
    mspans: List[Tuple[int, int]] = []
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
                and cur.part == nxt.part
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
                        part=cur.part,
                        deleted=deleted,
                        inserted=inserted,
                    )
                )
                mspans.append((spans[i][0], spans[i + 1][1]))
                i += 2
                continue
        merged.append(cur)
        mspans.append(spans[i])
        i += 1
    return merged, mspans


def _merge_format_runs(
    revisions: List[Revision], spans: List[Tuple[int, int]]
) -> List[Revision]:
    """Combine adjacent format-revision fragments into one row.

    One formatting action in Word (e.g. bolding a sentence) splits into as
    many runs as the sentence had; those fragments are adjacent on the run
    clock. Two separate formatted ranges in the same paragraph have plain
    text between them (non-adjacent spans) and stay separate rows.
    """
    merged: List[Revision] = []
    mspans: List[Tuple[int, int]] = []
    for rev, span in zip(revisions, spans):
        if merged:
            prev = merged[-1]
            pspan = mspans[-1]
            if (
                rev.kind == "format"
                and prev.kind == "format"
                and span != (-1, -1)
                and pspan[1] == span[0]
                and rev.author == prev.author
                and rev.part == prev.part
                and rev.para_index == prev.para_index
            ):
                prev.text = _collapse(prev.text + rev.text)
                mspans[-1] = (pspan[0], span[1])
                continue
        merged.append(rev)
        mspans.append(span)
    return merged


def _parse_comments(root: ET.Element) -> List[Comment]:
    comments: List[Comment] = []
    for c in root.findall(_COMMENT):
        pieces: List[str] = []
        last_para_id: Optional[str] = None
        for p in c.findall(f".//{_P}"):
            texts = [n.text or "" for n in p.iter() if n.tag in (_T, _DELTEXT, _M_T)]
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


def _extra_parts(names: List[str]) -> List[Tuple[str, str]]:
    """(zip name, display label) for footnote/endnote/header/footer parts."""
    found = [n for n in names if _EXTRA_PART_RE.match(n)]

    def sort_key(name: str) -> Tuple[int, int, str]:
        stem = name[len("word/") : -len(".xml")]
        for pos, (prefix, _) in enumerate(_PART_LABELS):
            if stem.startswith(prefix):
                suffix = stem[len(prefix) :]
                # natural order: header2.xml before header10.xml
                return (pos, int(suffix) if suffix.isdigit() else 0, name)
        return (len(_PART_LABELS), 0, name)

    out: List[Tuple[str, str]] = []
    for name in sorted(found, key=sort_key):
        stem = name[len("word/") : -len(".xml")]
        for prefix, label in _PART_LABELS:
            if stem.startswith(prefix):
                out.append((name, label))
                break
    return out


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
        body = doc_root.find(_BODY)
        if body is None:
            raise DocxReviewError("document.xml 中没有找到文档主体（w:body）")

        state = _WalkState()
        _walk_part(body, "", state)
        for name, label in _extra_parts(zf.namelist()):
            root = _read_part(zf, name)
            if root is not None:
                _walk_part(root, label, state)

        comments_root = _read_part(zf, "word/comments.xml")
        ext_root = _read_part(zf, "word/commentsExtended.xml")

    all_comments = _parse_comments(comments_root) if comments_root is not None else []
    ext = _parse_extended(ext_root) if ext_root is not None else {}

    body_part = state.parts[""]

    by_para_id = {c.para_id: c for c in all_comments if c.para_id}
    top: List[Comment] = []
    for c in all_comments:
        c.quoted = state.quoted.get(c.id, "")
        part_label, idx = state.anchor.get(c.id, ("", None))
        pd = state.parts.get(part_label, body_part)
        c.part = part_label
        c.para_index = idx
        c.para_text = pd.text_at(idx)
        c.heading = pd.nearest_heading(idx)
        parent_pid, done = ext.get(c.para_id, (None, False)) if c.para_id else (None, False)
        c.resolved = done
        parent = by_para_id.get(parent_pid) if parent_pid else None
        if parent is not None and parent is not c:
            parent.replies.append(c)
        else:
            top.append(c)

    big = 10**9
    top.sort(key=lambda c: state.anchor_order.get(c.id, big))

    revisions, spans = _merge_replacements(state.revisions, state.rev_spans)
    revisions = _merge_format_runs(revisions, spans)
    for rev in revisions:
        pd = state.parts.get(rev.part, body_part)
        rev.para_text = pd.text_at(rev.para_index)
        rev.heading = pd.nearest_heading(rev.para_index)

    return Review(
        source=p.name,
        paragraph_count=len(body_part.paragraphs),
        comments=top,
        revisions=revisions,
    )
