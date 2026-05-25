"""
docx2md — convert an academic Word .docx to clean Markdown.

Usage:
  python convert.py input.docx
  python convert.py input.docx -o output.md
  python convert.py input.docx --comments
  python convert.py input.docx -o output.md --comments comments.md

Default output paths (when -o / --comments are not given):
  Markdown  →  <docx_stem>.md         (same directory as input)
  Comments  →  <docx_stem>_comment.md (same directory as input)
  Images    →  <md_parent>/images/
"""
import argparse
from pathlib import Path

from _ast_filter import convert as _convert_md
from _extract_comments import extract as _extract_comments


def main() -> None:
    ap = argparse.ArgumentParser(
        prog="convert.py",
        description="Convert academic .docx to clean Markdown (+ optional comment extraction).",
    )
    ap.add_argument("docx", type=Path, help="Input .docx file")
    ap.add_argument(
        "-o", "--output", type=Path, default=None,
        metavar="MD_PATH",
        help="Output Markdown path (default: <stem>.md next to input)",
    )
    ap.add_argument(
        "--comments", nargs="?", const=True, default=False,
        metavar="COMMENTS_PATH",
        help="Also extract unresolved comments. Optional path (default: <stem>_comment.md)",
    )
    args = ap.parse_args()

    docx = args.docx.resolve()
    if not docx.exists():
        ap.error(f"File not found: {docx}")

    md_out = args.output or docx.with_suffix(".md")
    md_out = Path(md_out).resolve()

    # Step 1: convert docx → Markdown
    _convert_md(docx, md_out)
    print(f"Markdown : {md_out}")
    print(f"Images   : {md_out.parent / 'images'}")

    # Step 2: extract comments (optional)
    if args.comments is not False:
        if args.comments is True:
            comments_out = docx.with_name(docx.stem + "_comment.md")
        else:
            comments_out = Path(args.comments).resolve()
        if _extract_comments(docx, comments_out):
            print(f"Comments : {comments_out}")


if __name__ == "__main__":
    main()
