# docx2md

将复杂学术 Word 文档（含公式、修订、批注）转换为干净的 Markdown，并单独导出未解决的批注。

## How to Use

> "把这个 docx 转成 Markdown"

> "提取 manuscript.docx 中未解决的批注"

## Note

- 依赖 **pandoc**（需在 PATH 中）和 Python 3.8+（标准库，无需 pip 安装）。
- Word 中的表格包裹公式会被 pandoc 转成 ASCII 虚线块，`fix_docx_equations.py` 自动修复为标准 `$$ ... \tag{N} $$` 格式。
- 批注提取直接解析 docx XML，能区分已解决/未解决批注（pandoc 无法做到）。
- 需要 `docx` skill 的 `scripts/office/unpack.py` 来解包 docx。
