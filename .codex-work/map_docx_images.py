from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

docx = Path(r"C:\Users\maria_ym81dto\Documents\TCC REAL\docs\template-tcc-geologia-ufsc.docx")
ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
with ZipFile(docx) as z:
    root = ET.fromstring(z.read("word/document.xml"))
    relroot = ET.fromstring(z.read("word/_rels/document.xml.rels"))
rels = {e.get("Id"): e.get("Target") for e in relroot}
paras = root.findall(".//w:body//w:p", ns)
for i, p in enumerate(paras):
    embeds = [b.get("{" + ns["r"] + "}embed") for b in p.findall(".//a:blip", ns)]
    if not embeds:
        continue
    before = ""
    after = ""
    for q in reversed(paras[max(0, i-4):i]):
        s = "".join((t.text or "") for t in q.findall(".//w:t", ns)).strip()
        if s:
            before = s
            break
    for q in paras[i+1:i+5]:
        s = "".join((t.text or "") for t in q.findall(".//w:t", ns)).strip()
        if s:
            after = s
            break
    for rid in embeds:
        print(f"{rels.get(rid)}\tBEFORE={before[:100]}\tAFTER={after[:140]}")
