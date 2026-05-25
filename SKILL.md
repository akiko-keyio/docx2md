---
name: docx2md
description: "Convert academic Word .docx (with equations, tracked changes, comments) to clean Markdown. Extracts unresolved comments separately. Use when user asks to convert Word docs to Markdown."
---

# docx2md — Academic Word Manuscript to Markdown

Convert a complex academic `.docx` manuscript (with equations, tracked changes, and comments) to clean Markdown plus a separate unresolved-comments file.

## Dependencies

- **pandoc** (must be on PATH)
- **Python 3.8+** (standard library only — no pip packages needed)
- **`docx` skill** for `scripts/office/unpack.py` (docx unpacking)

## Usage

```bash
# Step 1: pandoc convert (accept all tracked changes, no hard wrap)
pandoc --track-changes=accept --wrap=none "<DOCX_PATH>" -o "<MD_PATH>"

# Step 2: Fix equations (table-wrapped → display math, in-place)
python "<SKILL_DIR>/scripts/fix_docx_equations.py" "<MD_PATH>"

# Step 3: Extract unresolved comments
python scripts/office/unpack.py "<DOCX_PATH>" "<UNPACKED_DIR>"
python "<SKILL_DIR>/scripts/extract_docx_comments.py" "<UNPACKED_DIR>" -o "<COMMENTS_MD_PATH>"

# Step 4 (optional): Clean up unpacked directory
Remove-Item -Recurse "<UNPACKED_DIR>"
```

> `<SKILL_DIR>` = this skill's install directory, typically `~/.claude/skills/docx2md`.

## Output Files

| File | Content |
|------|---------|
| `<stem>.md` | Clean Markdown with proper `$$ ... \tag{N} $$` equations |
| `<stem>_comment.md` | Unresolved comments with section context and annotated text |

## What Each Script Does

### `fix_docx_equations.py`

Word stores display equations inside single-cell borderless tables. Pandoc converts them to ASCII dash-bordered blocks:

```
  ---------------------------------------------------------------------------
  $$equation$$   \(N\)
  ------------------------------------------------------------------- -------

  ---------------------------------------------------------------------------
```

The script converts all such blocks to standard display math:

```markdown
$$
equation \tag{N}
$$
```

Handles:
- Single-line and multi-line equations (`\begin{matrix}` etc.)
- Tag formats: `\(N\)`, `(N)`, and `[]{#_EqN .anchor}(N)`
- Arbitrary dash-line widths

### `extract_docx_comments.py`

Parses three XML files from the unpacked docx:
- `word/comments.xml` — comment id, author, date, and text body
- `word/commentsExtended.xml` — `done="0"` / `done="1"` resolved status (linked by `paraId`)
- `word/document.xml` — `commentRangeStart`/`commentRangeEnd` for annotated text + heading context

Output format:

```markdown
## Comment {id} — {author} ({date})
**Section:** {nearest heading}
**Annotated text:** {text between commentRangeStart and commentRangeEnd}
**Comment:** {actual comment body}
  - **Reply** ({author}, {date}): {reply text}
```

## Why Not Just Pandoc?

| Problem | Pandoc behavior | This skill's fix |
|---------|----------------|------------------|
| Equations in tables | Outputs ASCII dash-bordered blocks | `fix_docx_equations.py` → `$$...\tag{N}$$` |
| Comments | Mixes comment text with annotated text, loses resolved status | `extract_docx_comments.py` parses XML directly |
| Line wrapping | Hard wraps at ~72 chars | `--wrap=none` flag |
| Tracked changes | Shows messy ins/del markup | `--track-changes=accept` for clean final text |

## Key Conventions

- **`--wrap=none` is essential** — without it, pandoc inserts hard line breaks making paragraphs unreadable
- **Never use pandoc for comment extraction** — it has no concept of resolved/unresolved
- **`commentsExtended.xml`** is the only place storing the resolved flag (`done="1"`)
- Equation fixer edits the file **in place**
