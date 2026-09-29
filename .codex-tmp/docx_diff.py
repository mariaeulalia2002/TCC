import json
import sys
from difflib import SequenceMatcher
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def paragraphs(path):
    doc = Document(path)
    out = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        has_drawing = bool(p._p.xpath('.//w:drawing | .//w:pict'))
        out.append({
            'index': i,
            'text': text,
            'style': p.style.name if p.style else None,
            'drawing': has_drawing,
        })
    return doc, out


def main(old_path, new_path):
    old_doc, old = paragraphs(old_path)
    new_doc, new = paragraphs(new_path)
    old_text = [p['text'] for p in old]
    new_text = [p['text'] for p in new]
    matcher = SequenceMatcher(None, old_text, new_text, autojunk=False)
    changes = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
        changes.append({
            'tag': tag,
            'old_range': [i1, i2],
            'new_range': [j1, j2],
            'old': old[i1:i2],
            'new': new[j1:j2],
        })
    result = {
        'old_paragraphs': len(old),
        'new_paragraphs': len(new),
        'old_images': len(old_doc.inline_shapes),
        'new_images': len(new_doc.inline_shapes),
        'changes': changes,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
