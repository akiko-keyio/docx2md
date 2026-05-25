#!/usr/bin/env python3
"""
Pandoc JSON AST filter for docx → clean Markdown.

Pipeline:
  pandoc input.docx --to json -o tmp.json
  python ast_filter.py tmp.json filtered.json
  pandoc --from json --wrap=none --to markdown filtered.json -o output.md

Or use the convert() helper for a single call.

Deterministic transformations (no heuristics, no regex on text):
  1. Two-column equation table → RawBlock("$$ ... \\tag{N} $$")
  2. Cross-reference Link(#_Fig*/#_Tab*/#_Eq*) → plain inline text
  3. Anchor Span(id=_Fig*/_Tab*/_Eq*, class=anchor) → strip wrapper
  4. Image with width/height attrs → clean Image with md2docx hint in alt
  5. Strong/Emph wrapping only a cross-ref → unwrap (ref was bolded in Word for display)
"""

import json
import re
import sys
from pathlib import Path


# ── Pandoc AST node shape helpers ─────────────────────────────────────────────

def _raw_block(text: str, fmt: str = "markdown") -> dict:
    return {"t": "RawBlock", "c": [fmt, text]}


def _collect_str(inlines: list) -> str:
    """Recursively collect plain string from inline nodes."""
    parts = []
    for n in inlines:
        if not isinstance(n, dict):
            continue
        t = n.get("t")
        if t == "Str":
            parts.append(n["c"])
        elif t == "Space":
            parts.append(" ")
        elif t in ("Strong", "Emph", "Span", "Quoted", "Link", "Superscript", "Subscript"):
            inner = n["c"]
            # content inlines are always the last list element for these types
            if isinstance(inner, list):
                parts.append(_collect_str(inner[-1] if isinstance(inner[-1], list) else inner))
    return "".join(parts)


def _collect_str_blocks(blocks: list) -> str:
    """Collect plain text from block nodes (Plain/Para wrapping inlines)."""
    parts = []
    for block in blocks:
        if isinstance(block, dict) and block.get("t") in ("Plain", "Para"):
            parts.append(_collect_str(block.get("c", [])))
    return " ".join(parts)


def _is_crossref_url(url: str) -> bool:
    return bool(re.match(r"#_(?:Fig|Tab|Eq)\d+", url))


def _is_anchor_span(attrs: list) -> bool:
    """attrs = [id, classes, kvs]"""
    if not attrs or len(attrs) < 2:
        return False
    span_id = attrs[0]
    classes = attrs[1]
    return bool(re.match(r"_(?:Fig|Tab|Eq)\d+", span_id)) or "anchor" in classes


# ── Table → equation detection ─────────────────────────────────────────────────

def _find_display_math(blocks: list):
    """Return LaTeX string if blocks contain a DisplayMath node, else None."""
    for block in blocks:
        if not isinstance(block, dict):
            continue
        t = block.get("t")
        if t in ("Plain", "Para"):
            for inline in block.get("c", []):
                if isinstance(inline, dict) and inline.get("t") == "Math":
                    if isinstance(inline["c"], list) and len(inline["c"]) >= 2:
                        if inline["c"][0].get("t") == "DisplayMath":
                            return inline["c"][1]
    return None


def _iter_table_cells(table_node: dict):
    """Yield (cell_idx_in_row, cell_blocks) for every cell in a table."""
    c = table_node["c"]
    # c[3] = head [head_attr, [rows]]
    # c[4] = bodies [[body_attr, head_rows, row_head_cols, [rows]], ...]
    # c[5] = foot [foot_attr, [rows]]
    def iter_rows(rows):
        for row in rows:
            # row = [row_attr, [cells]]
            cells = row[1] if isinstance(row, list) and len(row) > 1 else []
            for idx, cell in enumerate(cells):
                # cell = [cell_attr, align, rowspan, colspan, [blocks]]
                blocks = cell[4] if isinstance(cell, list) and len(cell) > 4 else []
                yield idx, blocks

    # head
    if isinstance(c[3], list) and len(c[3]) > 1:
        yield from iter_rows(c[3][1])
    # bodies
    for body in (c[4] if isinstance(c[4], list) else []):
        # body = [body_attr, row_head_cols, head_rows, body_rows]
        if isinstance(body, list) and len(body) > 3:
            yield from iter_rows(body[3])
    # foot
    if isinstance(c[5], list) and len(c[5]) > 1:
        yield from iter_rows(c[5][1])


def _try_equation_table(table_node: dict):
    """
    If table is a 2-column display-equation table, return (latex, tag_num).
    Otherwise return (None, None).
    Strategy: look for col spec count == 2, first column cell has DisplayMath.
    """
    c = table_node["c"]
    col_specs = c[2]
    if not isinstance(col_specs, list) or len(col_specs) != 2:
        return None, None

    math_tex = None
    tag_str = ""

    for cell_idx, cell_blocks in _iter_table_cells(table_node):
        if cell_idx == 0:
            math_tex = _find_display_math(cell_blocks)
        elif cell_idx == 1 and math_tex is not None:
            tag_str = _collect_str_blocks(cell_blocks).strip()

    if math_tex is None:
        return None, None

    num_match = re.search(r"\d+", tag_str)
    tag_num = num_match.group() if num_match else ""
    return math_tex, tag_num


# ── Image width hint ───────────────────────────────────────────────────────────

def _image_width_hint(kvs: list) -> str | None:
    """Return 'half' or '' (full) based on width attribute, or None if no width."""
    for k, v in kvs:
        if k == "width":
            m = re.search(r"([\d.]+)(in|cm|px|pt)", v)
            if m:
                val, unit = float(m.group(1)), m.group(2)
                val_in = {"in": val, "cm": val / 2.54, "px": val / 96, "pt": val / 72}.get(unit, val)
                return "" if val_in >= 5.0 else "half"
    return None


# ── Cross-ref parent unwrapping (Strong wrapping only a ref) ───────────────────

def _only_crossref_content(inlines: list) -> bool:
    """True if inlines contain AT LEAST ONE cross-ref Link and nothing else significant.
    Str and Space are allowed alongside links. Pure-text Strong nodes return False."""
    has_link = False
    for il in inlines:
        if not isinstance(il, dict):
            continue
        t = il.get("t")
        if t in ("Str", "Space"):
            continue
        if t == "Link" and _is_crossref_url(il["c"][2][0]):
            has_link = True
            continue
        return False  # something unexpected → don't unwrap
    return has_link


def _merge_adjacent_strong(inlines: list) -> list:
    """Merge adjacent Strong nodes (possibly separated by Space) into one Strong.
    Fixes captions rendered as **Fig.** **1** → **Fig. 1**."""
    result = []
    i = 0
    while i < len(inlines):
        cur = inlines[i]
        if isinstance(cur, dict) and cur.get("t") == "Strong":
            merged = list(cur["c"])
            j = i + 1
            while j < len(inlines):
                nxt = inlines[j]
                if not isinstance(nxt, dict):
                    break
                if nxt.get("t") == "Space" and j + 1 < len(inlines):
                    after = inlines[j + 1]
                    if isinstance(after, dict) and after.get("t") == "Strong":
                        merged.append(nxt)          # include Space
                        merged.extend(after["c"])   # absorb next Strong's content
                        j += 2
                        continue
                break
            result.append({"t": "Strong", "c": merged})
            i = j
        else:
            result.append(cur)
            i += 1
    return result


# ── Walker ────────────────────────────────────────────────────────────────────

def _walk_blocks(blocks: list) -> list:
    result = []
    for block in blocks:
        if not isinstance(block, dict):
            result.append(block)
            continue

        t = block.get("t")
        c = block.get("c")

        if t == "Table":
            math_tex, tag_num = _try_equation_table(block)
            if math_tex is not None:
                tag_part = f" \\tag{{{tag_num}}}" if tag_num else ""
                result.append(_raw_block(f"$$\n{math_tex}{tag_part}\n$$"))
                continue
            # Regular table: walk cell contents
            result.append(_walk_table(block))
            continue

        # Blocks that contain child blocks
        if t == "BlockQuote" and isinstance(c, list):
            result.append({"t": t, "c": _walk_blocks(c)})
        elif t in ("BulletList",) and isinstance(c, list):
            result.append({"t": t, "c": [_walk_blocks(item) for item in c]})
        elif t == "OrderedList" and isinstance(c, list):
            attrs, items = c
            result.append({"t": t, "c": [attrs, [_walk_blocks(item) for item in items]]})
        elif t == "Div" and isinstance(c, list):
            attrs, inner = c
            result.append({"t": t, "c": [attrs, _walk_blocks(inner)]})
        # Blocks that contain inlines
        elif t in ("Para", "Plain") and isinstance(c, list):
            result.append({"t": t, "c": _walk_inlines(c)})
        elif t == "Header" and isinstance(c, list):
            level, attrs, inlines = c
            result.append({"t": t, "c": [level, attrs, _walk_inlines(inlines)]})
        elif t == "LineBlock" and isinstance(c, list):
            result.append({"t": t, "c": [_walk_inlines(row) for row in c]})
        else:
            result.append(block)

    return result


def _walk_table(table_node: dict) -> dict:
    """Walk inlines inside a regular (non-equation) table."""
    c = list(table_node["c"])

    def walk_rows(rows):
        new_rows = []
        for row in rows:
            row_attr, cells = row[0], row[1]
            new_cells = []
            for cell in cells:
                cell_attr, align, rs, cs, blocks = cell
                new_cells.append([cell_attr, align, rs, cs, _walk_blocks(blocks)])
            new_rows.append([row_attr, new_cells])
        return new_rows

    # head: [head_attr, [rows]]
    if isinstance(c[3], list) and len(c[3]) > 1:
        c[3] = [c[3][0], walk_rows(c[3][1])]

    # bodies
    new_bodies = []
    for body in (c[4] if isinstance(c[4], list) else []):
        if isinstance(body, list) and len(body) > 3:
            body = list(body)
            body[3] = walk_rows(body[3])
        new_bodies.append(body)
    c[4] = new_bodies

    # foot: [foot_attr, [rows]]
    if isinstance(c[5], list) and len(c[5]) > 1:
        c[5] = [c[5][0], walk_rows(c[5][1])]

    return {"t": "Table", "c": c}


def _walk_inlines(inlines: list) -> list:
    result = []
    for inline in inlines:
        if not isinstance(inline, dict):
            result.append(inline)
            continue

        t = inline.get("t")
        c = inline.get("c")

        # Cross-reference link → flatten to content inlines
        if t == "Link" and _is_crossref_url(c[2][0]):
            result.extend(_walk_inlines(c[1]))
            continue

        # Anchor span → strip wrapper (keep content if any)
        if t == "Span" and _is_anchor_span(c[0]):
            result.extend(_walk_inlines(c[1]))
            continue

        # Image: strip pandoc dimension attrs (they render as noise in Markdown)
        if t == "Image":
            img_attrs, alt, target = c
            if any(k in ("width", "height") for k, _v in img_attrs[2]):
                new_attrs = [img_attrs[0], img_attrs[1], []]  # strip dimension attrs
                result.append({"t": "Image", "c": [new_attrs, alt, target]})
                continue

        # Strong/Emph containing only cross-ref links → unwrap (remove bold)
        if t in ("Strong", "Emph") and isinstance(c, list):
            if _only_crossref_content(c):
                result.extend(_walk_inlines(c))
                continue
            result.append({"t": t, "c": _walk_inlines(c)})
            continue

        # Other inlines with child inlines
        if t in ("Emph", "Underline", "Strikeout", "Superscript", "Subscript", "SmallCaps"):
            result.append({"t": t, "c": _walk_inlines(c)})
        elif t == "Quoted" and isinstance(c, list):
            qt, inner = c
            result.append({"t": t, "c": [qt, _walk_inlines(inner)]})
        elif t == "Span" and isinstance(c, list):
            attrs, inner = c
            result.append({"t": t, "c": [attrs, _walk_inlines(inner)]})
        elif t == "Link" and isinstance(c, list):
            attrs, inner, target = c
            result.append({"t": t, "c": [attrs, _walk_inlines(inner), target]})
        elif t == "Note" and isinstance(c, list):
            result.append({"t": t, "c": _walk_blocks(c)})
        else:
            result.append(inline)

    return _merge_adjacent_strong(result)


# ── Public API ────────────────────────────────────────────────────────────────

def filter_ast(ast: dict) -> dict:
    """Apply all deterministic transformations to a pandoc JSON AST."""
    ast = dict(ast)
    ast["blocks"] = _walk_blocks(ast["blocks"])
    return ast


# Greek char → LaTeX command name (for \mathbf{\cmd} matching in _postclean)
_GREEK_LATEX = {
    'Α': 'Alpha', 'Β': 'Beta', 'Γ': 'Gamma', 'Δ': 'Delta',
    'Ε': 'Epsilon', 'Ζ': 'Zeta', 'Η': 'Eta', 'Θ': 'Theta',
    'Ι': 'Iota', 'Κ': 'Kappa', 'Λ': 'Lambda', 'Μ': 'Mu',
    'Ν': 'Nu', 'Ξ': 'Xi', 'Ο': 'Omicron', 'Π': 'Pi',
    'Ρ': 'Rho', 'Σ': 'Sigma', 'Τ': 'Tau', 'Υ': 'Upsilon',
    'Φ': 'Phi', 'Χ': 'Chi', 'Ψ': 'Psi', 'Ω': 'Omega',
    'α': 'alpha', 'β': 'beta', 'γ': 'gamma', 'δ': 'delta',
    'ε': 'varepsilon', 'ζ': 'zeta', 'η': 'eta', 'θ': 'theta',
    'ι': 'iota', 'κ': 'kappa', 'λ': 'lambda', 'μ': 'mu',
    'ν': 'nu', 'ξ': 'xi', 'π': 'pi',
    'ρ': 'rho', 'σ': 'sigma', 'τ': 'tau', 'υ': 'upsilon',
    'φ': 'phi', 'χ': 'chi', 'ψ': 'psi', 'ω': 'omega',
    'ϕ': 'phi', 'ϑ': 'vartheta', 'ϵ': 'epsilon',
    'ϰ': 'varkappa', 'ϱ': 'varrho', 'ϖ': 'varpi',
    '∇': 'nabla', '∂': 'partial',
}


# Standard LaTeX operators that pandoc already wraps in \operatorname{}.
_PANDOC_OPERATORS = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth',
    'log', 'ln', 'exp', 'lim', 'sup', 'inf',
    'min', 'max', 'arg', 'det', 'dim', 'deg',
    'gcd', 'hom', 'ker', 'mod', 'Pr',
}


def _scan_docx_math(docx_path: Path) -> tuple[set, set]:
    """Single-pass scan of DOCX math runs.

    Returns (bi_chars, operator_names):
      - bi_chars: chars that appear exclusively with m:sty val="bi" (never "b")
      - operator_names: multi-letter text with m:sty val="p" but NO m:nor
        (m:nor = text label → pandoc wraps in \\text{}, not our concern)
    """
    import zipfile
    from lxml import etree

    M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    bi_chars, b_chars, op_names = set(), set(), set()

    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            doc_xml = zin.read('word/document.xml')
    except (KeyError, Exception):
        return set(), set()

    tree = etree.fromstring(doc_xml)
    for mr in tree.iter(f'{{{M}}}r'):
        mrPr = mr.find(f'{{{M}}}rPr')
        if mrPr is None:
            continue
        mt = mr.find(f'{{{M}}}t')
        if mt is None or not mt.text:
            continue
        sty = mrPr.find(f'{{{M}}}sty')
        nor = mrPr.find(f'{{{M}}}nor')
        val = sty.get(f'{{{M}}}val', '') if sty is not None else ''

        # Collect bold / bold-italic chars
        if val == 'bi':
            bi_chars.update(mt.text)
        elif val == 'b':
            b_chars.update(mt.text)

        # Operator names: m:sty val="p" WITHOUT m:nor
        if val == 'p' and nor is None:
            text = mt.text.strip()
            if (len(text) >= 2 and text[0].isalpha()
                    and text.isalnum() and text not in _PANDOC_OPERATORS):
                op_names.add(text)

    return bi_chars - b_chars, op_names


def _postclean(md_path: Path, bi_chars: set = None,
               operator_names: set = None) -> None:
    """Clean up pandoc Markdown artifacts in prose, preserving math blocks."""
    text = md_path.read_text(encoding="utf-8")

    # Make image paths relative and use images/ folder name
    media_abs = str(md_path.parent).replace("\\", "/") + "/media/"
    media_abs_win = str(md_path.parent) + "/media/"
    text = text.replace(media_abs, "images/")
    if media_abs_win != media_abs:
        text = text.replace(media_abs_win, "images/")
    text = text.replace("](media/", "](images/")

    # Split into math / non-math segments to protect math from changes
    segments = re.split(r"(\$\$.*?\$\$|\$[^\$\n]+?\$)", text, flags=re.DOTALL)

    cleaned = []
    for i, seg in enumerate(segments):
        if i % 2 == 1:
            # Inside math — fix pandoc OMML round-trip artifacts
            # \lbrack / \rbrack → [ / ] (OMML uses Unicode bracket names)
            seg = seg.replace(r"\left\lbrack", r"\left[")
            seg = seg.replace(r"\right\rbrack", r"\right]")
            # \ (backslash-space) → space (OMML spacing element read back as thin space)
            # Guard (?<!\\) to preserve \\ array row separators; then collapse double spaces
            seg = re.sub(r"(?<!\\)\\ ", " ", seg)
            seg = re.sub(r" {2,}", " ", seg)   # multiple spaces → one (safe in math)
            # \overset{\hat{}}{X} → \widehat{X} (pandoc encodes hat accent as overset)
            seg = re.sub(
                r"\\overset\{\\hat\{\}\}\{((?:[^{}]|\{[^{}]*\})*)\}",
                r"\\widehat{\1}", seg
            )
            # \mathbf{X} → \boldsymbol{X} for chars that had bi style in DOCX
            if bi_chars:
                for ch in bi_chars:
                    seg = seg.replace(f'\\mathbf{{{ch}}}', f'\\boldsymbol{{{ch}}}')
                    seg = seg.replace(f'{{\\mathbf{{{ch}}}}}', f'\\boldsymbol{{{ch}}}')
                    # Greek: also match LaTeX command form \mathbf{\phi} etc.
                    latex_name = _GREEK_LATEX.get(ch)
                    if latex_name:
                        seg = seg.replace(
                            f'\\mathbf{{\\{latex_name}}}',
                            f'\\boldsymbol{{\\{latex_name}}}'
                        )
            # Bare operator names that pandoc doesn't wrap → \operatorname{}
            if operator_names:
                pat = '|'.join(re.escape(n) for n in
                               sorted(operator_names, key=len, reverse=True))
                # (?<![a-zA-Z{]) prevents matching inside \text{}, \mathrm{}
                seg = re.sub(
                    rf'(?<![a-zA-Z{{])({pat})(?![a-zA-Z{{}}])',
                    r'\\operatorname{\1}', seg
                )
            cleaned.append(seg)
        else:
            seg = seg.replace("\\'", "'")   # coefficient\'s → coefficient's
            seg = seg.replace("\\~", "~")   # \~60 → ~60
            cleaned.append(seg)

    md_path.write_text("".join(cleaned), encoding="utf-8")


def convert(docx_path: str | Path, md_path: str | Path) -> None:
    """
    Full pipeline: docx → clean Markdown.
    Requires pandoc on PATH.
    """
    import subprocess, tempfile, os

    docx_path = Path(docx_path)
    md_path = Path(md_path)

    # Single-pass scan of DOCX math: bold-italic chars + operator names
    bi_chars, operator_names = _scan_docx_math(docx_path)

    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
        tmp_json = tmp.name

    try:
        # Step 1: docx → JSON AST (also extract embedded images)
        media_dir = md_path.parent
        subprocess.run(
            ["pandoc", "--track-changes=accept",
             f"--extract-media={media_dir}",
             str(docx_path), "--to", "json", "-o", tmp_json],
            check=True,
        )
        # Rename pandoc's default media/ to images/
        pandoc_media = media_dir / "media"
        images_dir = media_dir / "images"
        if pandoc_media.exists():
            if images_dir.exists():
                import shutil
                shutil.rmtree(images_dir)
            pandoc_media.rename(images_dir)

        # Step 2: filter AST
        with open(tmp_json, encoding="utf-8") as f:
            ast = json.load(f)
        ast = filter_ast(ast)
        with open(tmp_json, "w", encoding="utf-8") as f:
            json.dump(ast, f, ensure_ascii=False)

        # Step 3: JSON AST → Markdown (pipe_tables for clean rendering)
        subprocess.run(
            ["pandoc", "--from", "json", "--wrap=none",
             "--to", "markdown-simple_tables-multiline_tables-grid_tables+pipe_tables",
             tmp_json, "-o", str(md_path)],
            check=True,
        )

        # Step 4: clean up pandoc escape artifacts in prose
        _postclean(md_path, bi_chars=bi_chars,
                   operator_names=operator_names)

    finally:
        os.unlink(tmp_json)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(
        description="Filter pandoc JSON AST: equations, cross-refs, anchors, images."
    )
    parser.add_argument("input", nargs="?", help="Input JSON file (default: stdin)")
    parser.add_argument("-o", "--output", help="Output JSON file (default: stdout)")
    parser.add_argument("--docx", help="Shortcut: run full pipeline from .docx to .md")
    parser.add_argument("--md", help="Output .md path (used with --docx)")
    args = parser.parse_args()

    if args.docx:
        md_out = args.md or str(Path(args.docx).with_suffix(".md"))
        convert(args.docx, md_out)
        print(f"Written: {md_out}")
        return

    src = open(args.input, encoding="utf-8") if args.input else sys.stdin
    with src:
        ast = json.load(src)

    ast = filter_ast(ast)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(ast, f, ensure_ascii=False)
    else:
        json.dump(ast, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    main()
