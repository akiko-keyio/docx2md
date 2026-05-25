"""Fix pandoc-converted equations: replace table-wrapped equations with display math.

Usage: python scripts/fix_docx_equations.py <markdown_file>
Edits the file in place.
"""
import re, sys
from pathlib import Path


def is_dash_line(line):
    """Line is all dashes and spaces (at least 10 dashes total)."""
    stripped = line.strip()
    if not stripped:
        return False
    clean = stripped.replace(' ', '')
    return len(clean) >= 10 and all(c == '-' for c in clean)


def fix(path):
    text = path.read_text(encoding='utf-8')
    lines = text.split('\n')
    result = []
    i = 0
    fixed = 0

    while i < len(lines):
        if is_dash_line(lines[i]):
            block_start = i
            i += 1
            eq_lines = []

            while i < len(lines) and not is_dash_line(lines[i]):
                eq_lines.append(lines[i])
                i += 1

            eq_text_joined = '\n'.join(eq_lines)
            if i < len(lines) and is_dash_line(lines[i]) and '$$' in eq_text_joined:
                i += 1
                if i < len(lines) and lines[i].strip() == '':
                    next_i = i + 1
                    if next_i < len(lines) and is_dash_line(lines[next_i]):
                        i = next_i + 1
                    else:
                        i += 1

                eq_text = '\n'.join(line.rstrip() for line in eq_lines).strip()

                # Remove anchor tags like []{#_Eq2 .anchor}
                eq_text = re.sub(r'\[\]\{#_Eq\d+\s+\.anchor\}', '', eq_text)

                # Extract tag number: \(N\) or (N)
                tag = None
                tag_match = re.search(r'\\?\((\d+)\\?\)\s*$', eq_text, re.MULTILINE)
                if tag_match:
                    tag = tag_match.group(1)
                    eq_text = eq_text[:tag_match.start()] + eq_text[tag_match.end():]
                    eq_text = eq_text.strip()

                # Remove surrounding $$ markers
                if eq_text.startswith('$$'):
                    eq_text = eq_text[2:]
                if eq_text.endswith('$$'):
                    eq_text = eq_text[:-2]
                eq_text = eq_text.strip()

                if tag:
                    result.append(f'$$\n{eq_text} \\tag{{{tag}}}\n$$')
                else:
                    result.append(f'$$\n{eq_text}\n$$')
                fixed += 1
            else:
                result.append(lines[block_start])
                for el in eq_lines:
                    result.append(el)
                if i < len(lines) and is_dash_line(lines[i]):
                    result.append(lines[i])
                    i += 1
        else:
            result.append(lines[i])
            i += 1

    path.write_text('\n'.join(result), encoding='utf-8')
    print(f'Fixed {fixed} equations in {path.name}')


if __name__ == '__main__':
    fix(Path(sys.argv[1]))
