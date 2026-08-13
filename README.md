# docx-reply

[![CI](https://github.com/hc-ui/docx-reply/actions/workflows/ci.yml/badge.svg)](https://github.com/hc-ui/docx-reply/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**把 Word 批注和修订一键导出成"修改对照表"。**

Turn the review comments, threaded replies and tracked changes inside a Word document (`.docx`) into a fillable **revision response table** — as Markdown, CSV (Excel/WPS ready) or JSON. Zero dependencies, works offline, one command.

## 为什么需要它

论文盲审返修、导师批注返稿、期刊 revise——拿到一份满是批注和修订的 Word 文档后，几乎所有人都要做同一件事：**把每条意见逐条复制出来，做成"修改对照表 / 审稿意见回复表"再逐条答复**。手工复制费时、容易漏条，批注一多就是灾难。

`docx-reply` 直接读取 `.docx` 内部的批注数据，一条命令生成带"修改说明"空列的对照表：

- **批注全要素**：批注人、日期、批注内容、被批注的原文（精确到批注框选的文字）、所在位置（最近标题 + 段号，如"第3章 图像识别方法 · 第2段"）
- **全文档范围**：正文之外，脚注、尾注、页眉、页脚、文本框里的批注和修订同样提取，位置自动标注（如"脚注2 · 第1段"，脚注按可见编号计）；公式以线性文本呈现（如 `E=mc²`）
- **回复线程**：学生/合作者对批注的回复原样带出，主批注与回复不混淆
- **解决状态**：Word 里"标记为已解决"的批注自动标注，配合 `--skip-resolved` 过滤已处理项
- **修订全类型**：插入/删除/移动/格式变更都能列出；"选中旧词改新词"产生的紧邻删除+插入自动合并为一条"替换：旧词 → 新词"
- **四种输出**：Markdown 表格（默认）、CSV（带 BOM，Excel/WPS 双击直接打开不乱码）、JSON（供脚本处理）、**Word 版对照表（.docx，填完直接提交）**

## 安装

```bash
pip install git+https://github.com/hc-ui/docx-reply.git
```

无任何第三方依赖，Python 3.9+，支持 Microsoft Word 与 WPS Office 产生的文档。（PyPI 包名 `docx-reply` 上架中，上架后可直接 `pip install docx-reply`。）

## 使用

```bash
# 生成 Markdown 修改对照表
docx-reply 论文_导师批注.docx -o 修改对照表.md

# 生成 Word 版对照表，填完"修改说明"直接提交
docx-reply 论文_导师批注.docx -f docx -o 修改对照表.docx

# 生成 CSV（Excel/WPS 直接打开；修订另存为 修改对照表.revisions.csv）
docx-reply 论文_导师批注.docx -f csv -o 修改对照表.csv

# JSON 输出，供脚本/AI 工作流使用
docx-reply 论文_导师批注.docx -f json

# 只看某位审阅人的意见；跳过已解决的批注；不要修订记录
docx-reply 论文_导师批注.docx --author 王老师 --skip-resolved --no-revisions
```

对仓库自带的 [examples/sample.docx](examples/sample.docx) 运行 `docx-reply examples/sample.docx`，输出：

```markdown
# 审阅意见对照表：sample.docx

批注 4 条（含回复 1 条） ｜ 修订 2 处 ｜ 审阅人：王老师、李同学

## 批注（修改对照表）

| 序号 | 位置 | 原文摘录 | 批注人 | 日期 | 批注内容 | 回复 | 状态 | 修改说明 |
| ---: | --- | --- | --- | --- | --- | --- | --- | --- |
| 1 | 第3章 图像识别方法 · 第2段 | 深度学习模型在图像识别领域取得了显著进展 | 王老师 | 2026-08-10 | 这一段缺少对相关工作的引用，请补充 2-3 篇近三年文献。 | 李同学：已补充引用[15]-[17]。 | 未解决 |  |
| 2 | 第3章 图像识别方法 · 第4段 | 实验结果如图3-1所示 | 王老师 | 2026-08-10 | 图3-1 的分辨率太低，请替换为矢量图。 |  | 已解决 |  |
| 3 | 脚注1 · 第1段 | LeCun Y, Bengio Y, Hinton G. Deep learning. Nature, 2015. | 王老师 | 2026-08-10 | 参考文献格式请按 GB/T 7714 调整，补全卷期页码。 |  | 未解决 |  |

## 修订记录

| 序号 | 位置 | 类型 | 作者 | 日期 | 内容 |
| ---: | --- | --- | --- | --- | --- |
| 1 | 第3章 图像识别方法 · 第3段 | 替换 | 王老师 | 2026-08-11 | 非常 → 显著 |
| 2 | 第3章 图像识别方法 · 第3段 | 格式 | 王老师 | 2026-08-11 | 在两个公开数据集上均达到最优 |
```

注意几处细节："选中'非常'改成'显著'"在 Word 内部是一条删除加一条插入，`docx-reply` 自动合并为一条"替换"；老师把一句话加粗属于格式修订，以"格式"行列出；脚注里的批注按可见脚注编号定位。

最后一列"修改说明"留空，填完即可直接提交给导师或审稿人。

## 作为 Python 库使用

```python
from docx_reply import extract_review, render_markdown

review = extract_review("论文_导师批注.docx")
for comment in review.comments:
    print(comment.author, comment.text, comment.quoted, comment.resolved)
    for reply in comment.replies:
        print("  回复:", reply.author, reply.text)
for rev in review.revisions:
    print(rev.kind, rev.author, rev.text)

print(render_markdown(review))
```

JSON 结构（`-f json`）：

```json
{
  "source": "sample.docx",
  "stats": {"comments": 4, "replies": 1, "resolved": 1, "revisions": 2, "authors": ["王老师", "李同学"]},
  "comments": [
    {
      "author": "王老师",
      "text": "这一段缺少对相关工作的引用，请补充 2-3 篇近三年文献。",
      "quoted": "深度学习模型在图像识别领域取得了显著进展",
      "paragraph": 2,
      "heading": "第3章 图像识别方法",
      "resolved": false,
      "replies": [{"author": "李同学", "text": "已补充引用[15]-[17]。"}]
    }
  ],
  "revisions": [
    {"kind": "replace", "author": "王老师", "text": "非常 → 显著", "deleted": "非常", "inserted": "显著", "paragraph": 3}
  ]
}
```

## 与相关工具的区别

| 工具 | 形态 | 差异 |
|------|------|------|
| python-docx | Python 库 | 1.2 起有批注 API，但无 CLI、不读修订，需要自己写代码 |
| docx2python | Python 库 | 面向开发者的正文抽取，无回复线程/解决状态/对照表输出 |
| docx-review (Rust) | 开发者 CLI | 输出原始 JSON，面向自动化管线，需要 cargo 安装 |
| docxreview / docxtractr (R) | R 包 | 需要 R 环境 |
| **docx-reply** | **pip CLI + 库** | **一条命令直接得到可填写的修改对照表（含 Word 版 .docx 输出），回复线程、解决状态、替换合并、标题定位；零依赖** |

## Features (English)

- **Made for the revision-response workflow.** The default output is a fillable response table (修改对照表) — the exact artifact students and authors must produce after receiving a reviewed manuscript — not a raw data dump. `-f docx` even writes the table as a Word document ready to hand in.
- **Complete comment model.** Anchored text (exactly what the reviewer selected), heading-based location, threaded replies from `commentsExtended.xml`, and the resolved flag.
- **Tracked changes, humanized.** Insertions, deletions, moves and formatting changes (`w:ins` / `w:del` / `w:moveTo` / `w:rPrChange` / `w:pPrChange`) with author and date; an adjacent delete+insert pair by the same author (select-and-retype in Word) is folded into a single readable `old → new` replacement, and moved text is reported once, at its new location.
- **Whole-document coverage.** Comments and revisions in footnotes, endnotes, headers, footers and text boxes are extracted too, with the story name in the location column; formulas are linearized (`E=mc²`) instead of dropped.
- **Filters.** `--author 王老师` and `--skip-resolved` narrow the table to what still needs action.
- **Excel-friendly CSV.** Written with a UTF-8 BOM so Chinese text opens correctly in Excel and WPS by double-clicking.
- **Zero dependencies.** Pure standard library (`zipfile` + `xml.etree`), Python 3.9+, offline, cross-platform, console-encoding safe on Windows.
- **Scriptable.** `--format json` plus a small typed API (`extract_review`, `Review`, `Comment`, `Revision`).

## 局限与说明

- 位置以"最近标题 · 第 N 段"表示（段号按所在部分的全部段落顺序计）。`.docx` 文件里没有页码——页码是排版时才产生的，任何不调用 Word 排版引擎的工具都无法给出准确页码。
- 标题识别基于 Word/WPS 内置标题样式（Heading 1-9 / 标题 1-9）；纯手工加粗放大的"标题"无法识别，此时退回纯段号。
- 移动的文字在新位置以"移动"列出一次，暂不标注来源位置；格式修订只报告"哪些文字的格式变了"，不展开具体格式差异。
- 公式按阅读顺序线性提取为纯文本，不保留分式/上下标结构。
- 老式 `.doc` 二进制格式不支持，请先在 Word/WPS 中另存为 `.docx`。

## Roadmap

- [x] 按最近标题定位（"第3章 · 第2段"）— v0.2.0
- [x] `--author` / `--skip-resolved` 过滤 — v0.2.0
- [x] 直接输出 `.docx` 格式的修改对照表 — v0.2.0
- [x] 紧邻删除+插入合并为"替换" — v0.2.0
- [x] 脚注/尾注/页眉/页脚中的批注与修订 — v0.3.0
- [x] 移动修订、格式修订 — v0.3.0
- [ ] 修改说明列的 AI 草拟（对接本地 LLM）

## 贡献

欢迎 issue 与 PR。跑测试：

```bash
pip install -e ".[dev]"
pytest
```

遇到解析不正确的文档，欢迎[提 issue](https://github.com/hc-ui/docx-reply/issues) 并附上出问题的最小样例（可用 Word 新建一个小文档复现，不必上传原文）。

## License

[MIT](LICENSE)
