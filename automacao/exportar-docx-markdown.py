"""Exporta o DOCX do Google Docs para o Markdown intermediario do TCC.

O conversor principal do projeto consome Markdown porque ele precisa preservar
a ordem textual de legendas, tabelas e imagens. Este exportador le diretamente
o OOXML, mantendo estilos de titulo, italico/negrito, links, tabelas e a ordem
dos objetos graficos sem reduzir a resolucao das imagens.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "a": A, "r": R, "m": M}


def q(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def enabled(properties: ET.Element | None, name: str) -> bool:
    if properties is None:
        return False
    node = properties.find(f"w:{name}", NS)
    if node is None:
        return False
    value = node.attrib.get(q(W, "val"), "1").lower()
    return value not in {"0", "false", "off", "none"}


class Exporter:
    def __init__(self, source: Path):
        self.source = source
        self.document = Document(source)
        self.image_number = 0

    def relationship_url(self, element: ET.Element) -> str | None:
        rid = element.attrib.get(q(R, "id"))
        if not rid:
            return None
        rel = self.document.part.rels.get(rid)
        return rel.target_ref if rel and rel.is_external else None

    def run_text(self, run: ET.Element) -> str:
        pieces: list[str] = []
        for node in run.iter():
            if node.tag in {q(W, "t"), q(M, "t")}:
                pieces.append(node.text or "")
            elif node.tag == q(W, "tab"):
                pieces.append(" ")
            elif node.tag in {q(W, "br"), q(W, "cr")}:
                pieces.append("  \n")
            elif node.tag == q(W, "drawing"):
                self.image_number += 1
                pieces.append(f"![Imagem][image{self.image_number}]")
        text = "".join(pieces)
        if not text:
            return ""
        props = run.find("w:rPr", NS)
        if enabled(props, "b"):
            text = f"**{text}**"
        if enabled(props, "i"):
            text = f"*{text}*"
        return text

    def paragraph_text(self, element: ET.Element) -> str:
        parts: list[str] = []
        highlighted = False
        for child in element:
            if child.tag == q(W, "r"):
                parts.append(self.run_text(child))
                props = child.find("w:rPr", NS)
                mark = props.find("w:highlight", NS) if props is not None else None
                if mark is not None and mark.attrib.get(q(W, "val"), "none") != "none":
                    highlighted = True
            elif child.tag == q(W, "hyperlink"):
                label = "".join(self.run_text(run) for run in child.findall("w:r", NS))
                url = self.relationship_url(child)
                parts.append(f"[{label}]({url})" if url and label else label)
            elif child.tag in {q(M, "oMath"), q(M, "oMathPara")}:
                parts.append("".join(node.text or "" for node in child.findall(".//m:t", NS)))
        text = "".join(parts).replace("\u00a0", " ").strip()
        text = re.sub(r"[ \t]+", " ", text)
        if highlighted and text:
            text = f"*\\*{text}*"
        return text

    def paragraph_markdown(self, paragraph: Paragraph) -> str:
        text = self.paragraph_text(paragraph._p)
        style = paragraph.style.name if paragraph.style else ""
        match = re.fullmatch(r"Heading (\d+)", style)
        if match and text:
            return f"{'#' * int(match.group(1))} {text}"
        return text

    def cell_text(self, cell: _Cell) -> str:
        parts: list[str] = []
        for paragraph in cell.paragraphs:
            value = self.paragraph_text(paragraph._p)
            if value:
                parts.append(value)
        return "<br>".join(parts).replace("|", r"\|")

    def table_markdown(self, table: Table) -> list[str]:
        rows = [[self.cell_text(cell) for cell in row.cells] for row in table.rows]
        if not rows:
            return []
        columns = max(len(row) for row in rows)
        rows = [row + [""] * (columns - len(row)) for row in rows]
        output = ["| " + " | ".join(rows[0]) + " |"]
        output.append("| " + " | ".join("---" for _ in range(columns)) + " |")
        output.extend("| " + " | ".join(row) + " |" for row in rows[1:])
        return output

    def blocks(self):
        body = self.document.element.body
        for child in body.iterchildren():
            if child.tag == q(W, "p"):
                yield Paragraph(child, self.document)
            elif child.tag == q(W, "tbl"):
                yield Table(child, self.document)

    def export(self) -> str:
        output: list[str] = []
        for block in self.blocks():
            if isinstance(block, Paragraph):
                value = self.paragraph_markdown(block)
                if value:
                    output.append(value)
                    output.append("")
            else:
                output.extend(self.table_markdown(block))
                output.append("")
        return "\n".join(output).rstrip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    markdown = Exporter(args.source).export()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(markdown, encoding="utf-8")


if __name__ == "__main__":
    main()
