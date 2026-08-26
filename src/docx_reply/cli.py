"""Command-line interface: ``docx-reply reviewed.docx [-f md|csv|json] [-o out]``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .models import Comment, Review
from .parse import DocxReviewError, extract_review
from .render import (
    render_comments_csv,
    render_docx,
    render_json,
    render_markdown,
    render_revisions_csv,
)


def main(argv: list[str] | None = None) -> int:
    # Emit UTF-8 regardless of the console code page, so that redirecting
    # stdout to a file (docx-reply x.docx > out.md) yields a UTF-8 file on
    # Windows too instead of a locale-encoded (e.g. GBK) one; and never
    # crash on a console that cannot display some characters.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
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
        choices=("md", "csv", "json", "docx"),
        default="md",
        help="输出格式：md 修改对照表（默认）/ csv 表格（Excel、WPS 可直接打开）/ "
        "json 机器可读 / docx Word 版对照表（需配合 -o）",
    )
    parser.add_argument("-o", "--output", help="结果写入该文件（默认打印到标准输出）")
    parser.add_argument("--no-revisions", action="store_true", help="不输出修订记录（插入/删除）")
    parser.add_argument(
        "--author",
        action="append",
        metavar="姓名",
        help="只保留指定批注人/修订作者的记录，可重复使用",
    )
    parser.add_argument("--skip-resolved", action="store_true", help="跳过已标记为解决的批注")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = parser.parse_args(argv)

    if args.format == "docx" and not args.output:
        print("错误：-f docx 需要用 -o 指定输出文件，例如 -o 修改对照表.docx", file=sys.stderr)
        return 2

    try:
        review = extract_review(args.input)
    except DocxReviewError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2

    if args.no_revisions:
        review.revisions = []
    if args.author:
        # Word sometimes stores author names with stray surrounding spaces
        wanted = {a.strip() for a in args.author}
        review.comments = _filter_comments_by_author(review.comments, wanted)
        review.revisions = [r for r in review.revisions if r.author.strip() in wanted]
    if args.skip_resolved:
        review.comments = [c for c in review.comments if not c.resolved]

    try:
        return _emit(review, args)
    except OSError as exc:
        print(f"错误：无法写入输出文件：{exc}", file=sys.stderr)
        return 2


def _filter_comments_by_author(comments: list[Comment], wanted: set[str]) -> list[Comment]:
    """Keep a thread if the author or any nested reply matches.

    Matching a parent keeps the whole conversation. Matching only a reply
    keeps the parent for context and the matching reply subtree.
    """
    kept: list[Comment] = []
    for comment in comments:
        self_hit = comment.author.strip() in wanted
        child_kept = _filter_comments_by_author(comment.replies, wanted)
        if self_hit:
            kept.append(comment)
        elif child_kept:
            comment.replies = child_kept
            kept.append(comment)
    return kept


def _write_output(path: Path, *, text: str | None = None, data: bytes | None = None, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if data is not None:
        path.write_bytes(data)
        return
    path.write_text(text or "", encoding=encoding)


def _emit(review: Review, args: argparse.Namespace) -> int:
    if args.format == "docx":
        out = Path(args.output)
        _write_output(out, data=render_docx(review))
        print(f"已写入 {out}", file=sys.stderr)
        print(
            f"共 {review.total_comments} 条批注（含 {review.reply_count} 条回复）、{len(review.revisions)} 处修订",
            file=sys.stderr,
        )
        return 0

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
        _write_output(out, text=content, encoding=encoding)
        written = [str(out)]
        if args.format == "csv" and review.revisions:
            rev_path = out.with_name(out.stem + ".revisions.csv")
            _write_output(rev_path, text=render_revisions_csv(review), encoding="utf-8-sig")
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
