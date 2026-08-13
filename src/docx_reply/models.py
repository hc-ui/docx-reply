"""Data models shared by the parser and the renderers."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional


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
    para_id: Optional[str] = None
    quoted: str = ""
    para_index: Optional[int] = None
    para_text: str = ""
    resolved: bool = False
    replies: List["Comment"] = field(default_factory=list)

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
            "resolved": self.resolved,
            "replies": [r.to_dict() for r in self.replies],
        }


@dataclass
class Revision:
    """One tracked change (insertion or deletion)."""

    kind: str  # "insert" | "delete"
    author: str
    date: str
    text: str
    para_index: Optional[int] = None
    para_text: str = ""

    @property
    def date_short(self) -> str:
        return _short_date(self.date)

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "author": self.author,
            "date": self.date,
            "text": self.text,
            "paragraph": None if self.para_index is None else self.para_index + 1,
            "paragraph_text": self.para_text,
        }


def _count(comments: List[Comment]) -> int:
    return sum(1 + _count(c.replies) for c in comments)


@dataclass
class Review:
    """Everything extracted from one document."""

    source: str
    paragraph_count: int = 0
    comments: List[Comment] = field(default_factory=list)
    revisions: List[Revision] = field(default_factory=list)

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
    def authors(self) -> List[str]:
        seen: List[str] = []

        def add(name: str) -> None:
            if name and name not in seen:
                seen.append(name)

        def walk(comments: List[Comment]) -> None:
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
