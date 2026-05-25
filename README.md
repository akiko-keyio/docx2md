# docx2md

将复杂学术 Word 文档（含公式、修订、批注）转换为干净的 Markdown，并单独导出未解决的批注。

## How to Use

> "把这个 docx 转成 Markdown"

> "提取 manuscript.docx 中未解决的批注"

## Note

- 依赖 **pandoc**（需在 PATH 中），Python 3.8+ 标准库即可，无需 pip 安装。
- 自动修复 Word 公式转换问题，正确输出 `$$ ... \tag{N} $$` 格式。
- 批注导出能区分已解决/未解决状态（pandoc 本身无法做到）。
