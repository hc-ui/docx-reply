# Changelog

## 0.1.0 - 2026-08-13

首个版本。

- 解析 `.docx` 批注：批注人、日期、批注内容、被框选的原文、所在段落
- 回复线程与"已解决"状态（`commentsExtended.xml`），主批注按文档位置排序
- 修订记录：插入（`w:ins`）与删除（`w:del`），移动的文字（moveFrom/moveTo）在正文中正确呈现
- 输出：Markdown 修改对照表（默认，带空白"修改说明"列）、CSV（UTF-8 BOM，Excel/WPS 友好，修订另存 `.revisions.csv`）、JSON
- 纯标准库实现，零第三方依赖；Windows 控制台编码安全
- 支持批注跨段落、表格内批注、无范围的点批注
