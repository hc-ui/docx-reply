"""Generate examples/sample.docx — a mini "thesis returned by the advisor".

The file is built with the stdlib only and contains what a reviewed
document typically carries: comments in the body and in a footnote (one
threaded with a student reply, one resolved), a select-and-retype
replacement revision, and a formatting revision. All parts, content
types and relationships are declared, so the file opens cleanly in
Microsoft Word and WPS Office.

Run:  python examples/make_sample.py
"""

from __future__ import annotations

import zipfile
from pathlib import Path

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>
"""

DOCUMENT = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NS}><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第3章 图像识别方法</w:t></w:r></w:p>
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r><w:t>深度学习模型在图像识别领域取得了显著进展</w:t></w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r><w:commentReference w:id="0"/></w:r>
  <w:r><w:commentReference w:id="1"/></w:r>
  <w:r><w:t xml:space="preserve">，尤其是卷积神经网络在大规模数据集上的表现</w:t></w:r>
  <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>
  <w:r><w:t>。</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t xml:space="preserve">实验表明，改进后的模型精度</w:t></w:r>
  <w:ins w:id="10" w:author="王老师" w:date="2026-08-11T03:00:00Z"><w:r><w:t>显著</w:t></w:r></w:ins>
  <w:del w:id="11" w:author="王老师" w:date="2026-08-11T03:01:00Z"><w:r><w:delText>非常</w:delText></w:r></w:del>
  <w:r><w:t xml:space="preserve">优于基线方法，</w:t></w:r>
  <w:r>
    <w:rPr><w:b/><w:rPrChange w:id="12" w:author="王老师" w:date="2026-08-11T03:02:00Z"><w:rPr/></w:rPrChange></w:rPr>
    <w:t>在两个公开数据集上均达到最优</w:t>
  </w:r>
  <w:r><w:t>。</w:t></w:r>
</w:p>
<w:p>
  <w:commentRangeStart w:id="2"/>
  <w:r><w:t>实验结果如图3-1所示</w:t></w:r>
  <w:commentRangeEnd w:id="2"/>
  <w:r><w:commentReference w:id="2"/></w:r>
  <w:r><w:t>。</w:t></w:r>
</w:p>
</w:body></w:document>
"""

FOOTNOTES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes {NS}>
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="2">
  <w:p>
    <w:commentRangeStart w:id="3"/>
    <w:r><w:t>LeCun Y, Bengio Y, Hinton G. Deep learning. Nature, 2015.</w:t></w:r>
    <w:commentRangeEnd w:id="3"/>
    <w:r><w:commentReference w:id="3"/></w:r>
  </w:p>
</w:footnote>
</w:footnotes>
"""

COMMENTS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments {NS}>
<w:comment w:id="0" w:author="王老师" w:date="2026-08-10T02:00:00Z" w:initials="W">
  <w:p w14:paraId="C0000001"><w:r><w:t>这一段缺少对相关工作的引用，请补充 2-3 篇近三年文献。</w:t></w:r></w:p>
</w:comment>
<w:comment w:id="1" w:author="李同学" w:date="2026-08-12T01:00:00Z" w:initials="L">
  <w:p w14:paraId="C0000002"><w:r><w:t>已补充引用[15]-[17]。</w:t></w:r></w:p>
</w:comment>
<w:comment w:id="2" w:author="王老师" w:date="2026-08-10T02:05:00Z" w:initials="W">
  <w:p w14:paraId="C0000003"><w:r><w:t>图3-1 的分辨率太低，请替换为矢量图。</w:t></w:r></w:p>
</w:comment>
<w:comment w:id="3" w:author="王老师" w:date="2026-08-10T02:08:00Z" w:initials="W">
  <w:p w14:paraId="C0000004"><w:r><w:t>参考文献格式请按 GB/T 7714 调整，补全卷期页码。</w:t></w:r></w:p>
</w:comment>
</w:comments>
"""

COMMENTS_EX = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx {NS}>
<w15:commentEx w15:paraId="C0000001" w15:done="0"/>
<w15:commentEx w15:paraId="C0000002" w15:paraIdParent="C0000001" w15:done="0"/>
<w15:commentEx w15:paraId="C0000003" w15:done="1"/>
<w15:commentEx w15:paraId="C0000004" w15:done="0"/>
</w15:commentsEx>
"""


def main() -> None:
    out = Path(__file__).parent / "sample.docx"
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/_rels/document.xml.rels", DOC_RELS)
        zf.writestr("word/document.xml", DOCUMENT)
        zf.writestr("word/footnotes.xml", FOOTNOTES)
        zf.writestr("word/comments.xml", COMMENTS)
        zf.writestr("word/commentsExtended.xml", COMMENTS_EX)
    print(f"written: {out}")


if __name__ == "__main__":
    main()
