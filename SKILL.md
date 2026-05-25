---
name: docx2md
description: "Convert academic Word .docx (with equations, tracked changes, comments) to clean Markdown. Extracts unresolved comments separately. Use when user asks to convert Word docs to Markdown."
---

# docx2md

## Dependencies

pandoc (on PATH), Python 3.8+ (stdlib only)

## Workflow

### Step 1. Run converter

```bash
python "<SKILL_DIR>/scripts/convert.py" "<DOCX_PATH>" --comments
```

Output (all in same directory as input):
- `<stem>.md` — clean Markdown
- `images/` — extracted images
- `<stem>_comment.md` — unresolved comments (if `--comments`)

### Step 2. Review and fix

Read the output `.md`, fix any obvious conversion artifacts (garbled formatting, broken structure, etc.).

