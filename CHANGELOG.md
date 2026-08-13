# Changelog

## 0.2.0 - 2026-08-13

- **Word 版对照表输出**：`-f docx -o 修改对照表.docx` 直接生成可填写提交的 Word 表格
- **替换合并**：同一作者紧邻的删除+插入（Word 里"选中改写"的产物）合并为一条"替换：旧词 → 新词"（JSON 中 `kind: "replace"`，含 `deleted` / `inserted` 字段）
- **标题定位**：位置列升级为"最近标题 · 第 N 段"（识别 Word/WPS 内置标题样式），JSON 新增 `heading` 字段
- **过滤器**：`--author 姓名`（可重复）只保留指定审阅人；`--skip-resolved` 跳过已解决批注
- `-f docx` 未指定 `-o` 时给出明确错误提示

## 0.1.0 - 2026-08-13

首个版本。

- 解析 `.docx` 批注：批注人、日期、批注内容、被框选的原文、所在段落
- 回复线程与"已解决"状态（`commentsExtended.xml`），主批注按文档位置排序
- 修订记录：插入（`w:ins`）与删除（`w:del`），移动的文字（moveFrom/moveTo）在正文中正确呈现
- 输出：Markdown 修改对照表（默认，带空白"修改说明"列）、CSV（UTF-8 BOM，Excel/WPS 友好，修订另存 `.revisions.csv`）、JSON
- 纯标准库实现，零第三方依赖；Windows 控制台编码安全
- 支持批注跨段落、表格内批注、无范围的点批注
