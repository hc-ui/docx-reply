"""Generate examples/sample.docx — a mini "thesis returned by the advisor".

The file is built with the stdlib only, and contains what a reviewed
document typically carries: two comments (one threaded with a student
reply, one resolved) plus one insertion and one deletion.

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
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
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
  <w:r><w:t xml:space="preserve">，尤其是卷积神经网络在大规模数据集上的表现。</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:t xml:space="preserve">实验表明，改进后的模型精度</w:t></w:r>
  <w:ins w:id="10" w:author="王老师" w:date="2026-08-11T03:00:00Z"><w:r><w:t>显著</w:t></w:r></w:ins>
  <w:del w:id="11" w:author="王老师" w:date="2026-08-11T03:01:00Z"><w:r><w:delText>非常</w:delText></w:r></w:del>
  <w:r><w:t xml:space="preserve">优于基线方法。</w:t></w:r>
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
</w:comments>
"""

COMMENTS_EX = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx {NS}>
<w15:commentEx w15:paraId="C0000001" w15:done="0"/>
<w15:commentEx w15:paraId="C0000002" w15:paraIdParent="C0000001" w15:done="0"/>
<w15:commentEx w15:paraId="C0000003" w15:done="1"/>
</w15:commentsEx>
"""


def main() -> None:
    out = Path(__file__).parent / "sample.docx"
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", DOCUMENT)
        zf.writestr("word/comments.xml", COMMENTS)
        zf.writestr("word/commentsExtended.xml", COMMENTS_EX)
    print(f"written: {out}")


if __name__ == "__main__":
    main()
