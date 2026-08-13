"""Data models shared by the parser and the renderers."""

from __future__ import annotations

from dataclasses import dataclass, field


def _short_date(date: str) -> str:
    """``2026-08-10T02:00:00Z`` -> ``2026-08-10``."""
    return date.split("T")[0] if date else ""


@dataclass
class Comment:
    """One comment, possibly carrying threaded replies."""

    id: str
    author: str
    initials: str
    date: str
    text: str
    # paraId of the comment's last paragraph; links to commentsExtended.xml
    para_id: str | None = None
    quoted: str = ""
    para_index: int | None = None
    para_text: str = ""
    heading: str = ""
    # document story holding the anchor: "" body, 脚注/尾注/页眉/页脚
    part: str = ""
    resolved: bool = False
    replies: list[Comment] = field(default_factory=list)

    @property
    def date_short(self) -> str:
        return _short_date(self.date)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "author": self.author,
            "date": self.date,
            "text": self.text,
            "quoted": self.quoted,
            "paragraph": None if self.para_index is None else self.para_index + 1,
            "paragraph_text": self.para_text,
            "heading": self.heading or None,
            "part": self.part or None,
            "resolved": self.resolved,
            "replies": [r.to_dict() for r in self.replies],
        }


@dataclass
class Revision:
    """One tracked change: insert, delete, replace, move or format.

    For ``kind == "replace"`` (a deletion immediately followed by an
    insertion by the same author, i.e. select-and-retype in Word),
    ``deleted``/``inserted`` hold both sides and ``text`` a readable
    ``old → new`` form. For ``kind == "move"`` the text is reported once,
    at its new location. For ``kind == "format"`` the text is unchanged
    content whose formatting was modified.
    """

    kind: str  # "insert" | "delete" | "replace" | "move" | "format"
    author: str
    date: str
    text: str
    para_index: int | None = None
    para_text: str = ""
    heading: str = ""
    part: str = ""
    deleted: str | None = None
    inserted: str | None = None

    @property
    def date_short(self) -> str:
        return _short_date(self.date)

    def to_dict(self) -> dict:
        d = {
            "kind": self.kind,
            "author": self.author,
            "date": self.date,
            "text": self.text,
            "paragraph": None if self.para_index is None else self.para_index + 1,
            "paragraph_text": self.para_text,
            "heading": self.heading or None,
            "part": self.part or None,
        }
        if self.kind == "replace":
            d["deleted"] = self.deleted
            d["inserted"] = self.inserted
        return d


def _count(comments: list[Comment]) -> int:
    return sum(1 + _count(c.replies) for c in comments)


@dataclass
class Review:
    """Everything extracted from one document."""

    source: str
    paragraph_count: int = 0
    comments: list[Comment] = field(default_factory=list)
    revisions: list[Revision] = field(default_factory=list)

    @property
    def total_comments(self) -> int:
        return _count(self.comments)

    @property
    def reply_count(self) -> int:
        return self.total_comments - len(self.comments)

    @property
    def resolved_count(self) -> int:
        return sum(1 for c in self.comments if c.resolved)

    @property
    def authors(self) -> list[str]:
        seen: list[str] = []

        def add(name: str) -> None:
            if name and name not in seen:
                seen.append(name)

        def walk(comments: list[Comment]) -> None:
            for c in comments:
                add(c.author)
                walk(c.replies)

        walk(self.comments)
        for rev in self.revisions:
            add(rev.author)
        return seen

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "paragraphs": self.paragraph_count,
            "stats": {
                "comments": self.total_comments,
                "replies": self.reply_count,
                "resolved": self.resolved_count,
                "revisions": len(self.revisions),
                "authors": self.authors,
            },
            "comments": [c.to_dict() for c in self.comments],
            "revisions": [r.to_dict() for r in self.revisions],
        }
