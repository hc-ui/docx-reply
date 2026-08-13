"""Build minimal .docx files for tests using nothing but the stdlib."""

from __future__ import annotations

import zipfile

import pytest

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
)

DOCUMENT_TMPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f"<w:document {_NS}><w:body>{{body}}</w:body></w:document>"
)

COMMENTS_TMPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f"<w:comments {_NS}>{{comments}}</w:comments>"
)

COMMENTS_EX_TMPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f"<w15:commentsEx {_NS}>{{items}}</w15:commentsEx>"
)


def write_docx(path, body, comments=None, comments_ex=None):
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", DOCUMENT_TMPL.format(body=body))
        if comments is not None:
            zf.writestr("word/comments.xml", COMMENTS_TMPL.format(comments=comments))
        if comments_ex is not None:
            zf.writestr("word/commentsExtended.xml", COMMENTS_EX_TMPL.format(items=comments_ex))
    return path


@pytest.fixture
def docx_factory(tmp_path):
    counter = {"n": 0}

    def build(body, comments=None, comments_ex=None):
        counter["n"] += 1
        return write_docx(tmp_path / f"test{counter['n']}.docx", body, comments, comments_ex)

    return build


def p(*runs: str) -> str:
    """A paragraph made of raw inner XML fragments."""
    return "<w:p>" + "".join(runs) + "</w:p>"


def r(text: str) -> str:
    """A run holding plain text."""
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def comment_xml(cid, author, text, date="2026-08-10T02:00:00Z", para_id=None, initials="X"):
    pid = f' w14:paraId="{para_id}"' if para_id else ""
    return (
        f'<w:comment w:id="{cid}" w:author="{author}" w:date="{date}" w:initials="{initials}">'
        f"<w:p{pid}><w:r><w:t>{text}</w:t></w:r></w:p>"
        f"</w:comment>"
    )


def anchored(text: str, cid) -> str:
    """Text wrapped in a comment range with its reference run."""
    return (
        f'<w:commentRangeStart w:id="{cid}"/>'
        + r(text)
        + f'<w:commentRangeEnd w:id="{cid}"/>'
        + f'<w:r><w:commentReference w:id="{cid}"/></w:r>'
    )
