"""Command-line interface: ``docx-reply reviewed.docx [-f md|csv|json] [-o out]``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .parse import DocxReviewError, extract_review
from .render import render_comments_csv, render_json, render_markdown, render_revisions_csv


def main(argv: "list[str] | None" = None) -> int:
    # Never crash on a console that cannot display some characters
    # (e.g. Chinese text on a non-Chinese Windows console).
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except (ValueError, OSError):
                pass

    parser = argparse.ArgumentParser(
        prog="docx-reply",
        description="提取 Word 文档（.docx）中的批注、回复与修订，生成可直接填写的修改对照表。",
        epilog="示例：docx-reply 论文_导师批注.docx -o 修改对照表.md",
    )
    parser.add_argument("input", help="带批注/修订的 .docx 文件路径")
    parser.add_argument(
        "-f",
        "--format",
        choices=("md", "csv", "json"),
        default="md",
        help="输出格式：md 修改对照表（默认）/ csv 表格（Excel、WPS 可直接打开）/ json 机器可读",
    )
    parser.add_argument("-o", "--output", help="结果写入该文件（默认打印到标准输出）")
    parser.add_argument("--no-revisions", action="store_true", help="不输出修订记录（插入/删除）")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    try:
        review = extract_review(args.input)
    except DocxReviewError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2

    if args.no_revisions:
        review.revisions = []

    if args.format == "md":
        content = render_markdown(review)
    elif args.format == "json":
        content = render_json(review)
    else:
        content = render_comments_csv(review)

    if args.output:
        out = Path(args.output)
        # BOM so that double-clicking the CSV opens correctly in Excel/WPS
        encoding = "utf-8-sig" if args.format == "csv" else "utf-8"
        out.write_text(content, encoding=encoding)
        written = [str(out)]
        if args.format == "csv" and review.revisions:
            rev_path = out.with_name(out.stem + ".revisions.csv")
            rev_path.write_text(render_revisions_csv(review), encoding="utf-8-sig")
            written.append(str(rev_path))
        print("已写入 " + "、".join(written), file=sys.stderr)
    else:
        print(content)
        if args.format == "csv" and review.revisions:
            print(
                "提示：csv 模式只打印批注表；加 -o 输出文件可同时得到 <文件名>.revisions.csv 修订表",
                file=sys.stderr,
            )

    print(
        f"共 {review.total_comments} 条批注（含 {review.reply_count} 条回复）、{len(review.revisions)} 处修订",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
