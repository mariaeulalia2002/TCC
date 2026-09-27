from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree
import re, json, shutil

SRC = Path('.codex-work/latest-google-docs.docx')
CLEAN = Path('.codex-work/latest-google-docs-clean.docx')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W, 'm': M}

with ZipFile(SRC) as zin, ZipFile(CLEAN, 'w', ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.endswith('.xml'):
            root = etree.fromstring(data)
            changed = False
            for el in root.iter():
                for key, value in list(el.attrib.items()):
                    if key.startswith('{'+W+'}') and re.fullmatch(r'-?\d+\.\d+', value):
                        el.attrib[key] = str(round(float(value)))
                        changed = True
            if changed:
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        zout.writestr(item, data)

with ZipFile(CLEAN) as z:
    root = etree.fromstring(z.read('word/document.xml'))

paras = []
for i, p in enumerate(root.xpath('.//w:body/w:p', namespaces=NS)):
    pieces = []
    for node in p.iter():
        if node.tag in {f'{{{W}}}t', f'{{{M}}}t'} and node.text:
            pieces.append(node.text)
        elif node.tag == f'{{{W}}}tab':
            pieces.append('\t')
        elif node.tag == f'{{{W}}}br':
            pieces.append('\n')
    text = ''.join(pieces).strip()
    omath = bool(p.xpath('.//m:oMath|.//m:oMathPara', namespaces=NS))
    style = p.xpath('./w:pPr/w:pStyle/@w:val', namespaces=NS)
    if text or omath:
        paras.append({'index': i, 'style': style[0] if style else '', 'equation': omath, 'text': text})

Path('.codex-work/docx-paragraphs.json').write_text(json.dumps(paras, ensure_ascii=False, indent=2), encoding='utf-8')
eqs = [p for p in paras if p['equation'] or '\\\\[' in p['text']]
Path('.codex-work/equations.txt').write_text('\n'.join(f"{p['index']}\t{p['style']}\t{p['text']}" for p in eqs), encoding='utf-8')
print(json.dumps({'paragraphs': len(paras), 'equations': len(eqs), 'clean': str(CLEAN)}, ensure_ascii=False))
