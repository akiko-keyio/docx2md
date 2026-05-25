"""Extract unresolved comments from unpacked docx XML.

Usage: python scripts/extract_docx_comments.py <unpacked_dir> [-o output.md]
"""
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}


def get_all_text(elem):
    texts = []
    for t in elem.iter(f'{{{NS["w"]}}}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def extract(base: Path, output: Path):
    # --- 1. Parse comments.xml ---
    comments = {}
    for c in ET.parse(base / 'word/comments.xml').findall('.//w:comment', NS):
        cid = c.get(f'{{{NS["w"]}}}id')
        texts = []
        for t in c.iter(f'{{{NS["w"]}}}t'):
            if t.text:
                texts.append(t.text)
        para_ids = []
        for p in c.findall(f'{{{NS["w"]}}}p', NS):
            pid = p.get(f'{{{NS["w14"]}}}paraId')
            if pid:
                para_ids.append(pid)
        comments[cid] = {
            'author': c.get(f'{{{NS["w"]}}}author'),
            'date': c.get(f'{{{NS["w"]}}}date', '')[:10],
            'text': ''.join(texts),
            'paraIds': para_ids,
        }

    # --- 2. Parse commentsExtended.xml for resolved status ---
    done_map, parent_map = {}, {}
    for ex in ET.parse(base / 'word/commentsExtended.xml').findall('.//w15:commentEx', NS):
        pid = ex.get(f'{{{NS["w15"]}}}paraId')
        done_map[pid] = ex.get(f'{{{NS["w15"]}}}done') == '1'
        ppid = ex.get(f'{{{NS["w15"]}}}paraIdParent')
        if ppid:
            parent_map[pid] = ppid

    resolved = {}
    comment_parent = {}
    for cid, info in comments.items():
        for pid in info['paraIds']:
            if pid in done_map:
                resolved[cid] = done_map[pid]
            if pid in parent_map:
                ppid = parent_map[pid]
                for oid, oinfo in comments.items():
                    if ppid in oinfo['paraIds']:
                        comment_parent[cid] = oid
                        break

    # --- 3. Parse document.xml for annotated text and section context ---
    doc_root = ET.parse(base / 'word/document.xml').getroot()
    body = doc_root.find(f'{{{NS["w"]}}}body')
    all_paras = list(body.iter(f'{{{NS["w"]}}}p'))

    # Headings
    heading_at = {}  # para index -> heading text
    for idx, p in enumerate(all_paras):
        pPr = p.find(f'{{{NS["w"]}}}pPr')
        if pPr is not None:
            pStyle = pPr.find(f'{{{NS["w"]}}}pStyle')
            if pStyle is not None:
                val = pStyle.get(f'{{{NS["w"]}}}val', '')
                if 'Heading' in val or val in ('1', '2', '3'):
                    heading_at[idx] = get_all_text(p)

    def section_for(para_idx):
        for i in range(para_idx, -1, -1):
            if i in heading_at:
                return heading_at[i]
        return ''

    # Comment ranges
    range_start, range_end = {}, {}
    for idx, p in enumerate(all_paras):
        for elem in p.iter():
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if tag == 'commentRangeStart':
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    range_start[cid] = idx
            elif tag == 'commentRangeEnd':
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    range_end[cid] = idx

    annotated = {}
    for cid in comments:
        if cid in range_start and cid in range_end:
            parts = []
            for i in range(range_start[cid], range_end[cid] + 1):
                t = get_all_text(all_paras[i]).strip()
                if t:
                    parts.append(t)
            annotated[cid] = ' '.join(parts)

    section_map = {cid: section_for(range_start[cid]) for cid in comments if cid in range_start}

    # --- 4. Output unresolved comments ---
    unresolved = [cid for cid in comments if not resolved.get(cid, False)]
    unresolved.sort(key=lambda x: int(x))
    top_level = [cid for cid in unresolved if cid not in comment_parent]
    replies = {cid: [] for cid in top_level}
    for cid in unresolved:
        if cid in comment_parent and comment_parent[cid] in replies:
            replies[comment_parent[cid]].append(cid)

    lines = [f'# Unresolved comments — {base.name}\n']
    for cid in top_level:
        info = comments[cid]
        section = section_map.get(cid, '')
        ann = annotated.get(cid, '(range not found)')
        if len(ann) > 400:
            ann = ann[:400] + '...'

        lines.append('---\n')
        lines.append(f'## Comment {cid} — {info["author"]} ({info["date"]})\n')
        if section:
            lines.append(f'**Section:** {section}\n')
        lines.append(f'**Annotated text:** {ann}\n')
        lines.append(f'**Comment:** {info["text"]}\n')
        for rid in replies.get(cid, []):
            ri = comments[rid]
            lines.append(f'  - **Reply** ({ri["author"]}, {ri["date"]}): {ri["text"]}\n')
        lines.append('')

    output.write_text('\n'.join(lines), encoding='utf-8')
    n_resolved = sum(1 for v in resolved.values() if v)
    print(f'Total: {len(comments)}, resolved: {n_resolved}, unresolved output: {len(top_level)} top-level + {sum(len(v) for v in replies.values())} replies')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('unpacked_dir', type=Path)
    ap.add_argument('-o', '--output', type=Path, default=None)
    args = ap.parse_args()
    out = args.output or (args.unpacked_dir.parent / 'comments.md')
    extract(args.unpacked_dir, out)
