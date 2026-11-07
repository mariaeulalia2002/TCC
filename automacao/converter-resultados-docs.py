from __future__ import annotations

import base64
import io
import math
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageChops, ImageOps, ImageStat


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "docs" / "template-tcc-geologia-ufsc.docx"
MARKDOWN = ROOT / "docs" / "google-docs-latest.md"
RESULTADOS = ROOT / "latex" / "capitulos" / "05-resultados.tex"
DISCUSSAO = ROOT / "latex" / "capitulos" / "06-discussao.tex"
CONCLUSOES = ROOT / "latex" / "capitulos" / "07-conclusoes.tex"
APENDICE_D = ROOT / "latex" / "pos-textuais" / "apendice-d.tex"
FIGDIR = ROOT / "latex" / "figuras" / "resultados"
REFERENCIAS = ROOT / "latex" / "referencias.bib"
LANDSCAPE_FIGURES: set[int] = set()

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def between(text: str, start: str, end: str) -> str:
    a = text.index(start)
    b = text.index(end, a)
    return text[a:b]


def between_headings(text: str, start_pattern: str, end_pattern: str) -> str:
    """Return a Markdown section, accepting Google Docs bold heading markers."""
    start = re.search(start_pattern, text, flags=re.MULTILINE)
    if not start:
        raise ValueError(f"Cabeçalho inicial não encontrado: {start_pattern}")
    end = re.search(end_pattern, text[start.end():], flags=re.MULTILINE)
    if not end:
        raise ValueError(f"Cabeçalho final não encontrado: {end_pattern}")
    return text[start.start():start.end() + end.start()]


def clean_heading(text: str) -> str:
    text = re.sub(r"^#+\s*", "", text).strip()
    if text.startswith("**") and text.endswith("**"):
        text = text[2:-2].strip()
    text = re.sub(r"^(?:\d+(?:\.\d+)*|D\.\d+)\s+", "", text)
    return text.strip()


def slug(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    text = re.sub(r"[^a-zA-Z0-9]+", "-", text).strip("-").lower()
    return text[:48]


def protect_math(text: str):
    parts: list[str] = []

    def repl(match):
        parts.append(match.group(0))
        return f"@@MATH{len(parts)-1}@@"

    return re.sub(r"\$[^$]+\$", repl, text), parts


def normalize_author(value: str) -> str:
    value = value.replace("{", "").replace("}", "")
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    value = re.sub(r"\s+", " ", value).strip().upper()
    return value


def citation_keys() -> dict[tuple[str, str], str]:
    """Index BibTeX keys by first author and year for Docs citation conversion."""
    bib = REFERENCIAS.read_text(encoding="utf-8")
    result: dict[tuple[str, str], str] = {}
    entries = re.finditer(r"(?ms)^@\w+\{([^,]+),(.*?)(?=^@\w+\{|\Z)", bib)
    for entry in entries:
        key, fields = entry.group(1).strip(), entry.group(2)
        year = re.search(r"(?i)\byear\s*=\s*[\{\"](\d{4})", fields)
        author = re.search(r"(?is)\bauthor\s*=\s*\{(.+?)\}\s*,", fields)
        if not (year and author):
            continue
        first = author.group(1).split(" and ", 1)[0].strip()
        if "," in first:
            surname = first.split(",", 1)[0]
        else:
            surname = first.split()[-1]
        result[(normalize_author(surname), year.group(1))] = key

    # Corporate authors and conventional citation forms used in the TCC.
    aliases = {
        ("SESAME", "2004"): "sesame2004",
        ("ICS", "2026"): "ics2026",
        ("INTERNATIONAL COMMISSION ON STRATIGRAPHY", "2026"): "ics2026",
        ("JCGM", "2008"): "jcgm2008",
        ("BROWN JR.", "1980"): "brownfisher1980",
        ("MITCHUM JR.", "1977"): "mitchum1977",
        ("MECKEL", "2011"): "meckel2011",
    }
    result.update(aliases)
    return result


CITATION_KEYS = citation_keys()


def lookup_citation(author: str, year: str) -> str | None:
    author = re.sub(r"\*?et al\.\*?", "", author, flags=re.IGNORECASE).strip()
    author = author.split("&", 1)[0].strip()
    return CITATION_KEYS.get((normalize_author(author), year))


def convert_citations(text: str) -> str:
    """Convert ABNT author-date citations exported by Google Docs to BibTeX links."""

    def parenthetical(match: re.Match[str]) -> str:
        content = match.group(1)
        if not re.search(r"(?:19|20)\d{2}", content):
            return match.group(0)
        keys: list[str] = []
        current_author: str | None = None
        for part in [p.strip() for p in content.split(";")]:
            item = re.fullmatch(
                r"(.+?),\s*((?:19|20)\d{2})(?:,\s*((?:19|20)\d{2}))?",
                part,
            )
            if not item:
                return match.group(0)
            current_author = item.group(1).strip()
            years = [item.group(2)] + ([item.group(3)] if item.group(3) else [])
            for year in years:
                key = lookup_citation(current_author, year)
                if not key:
                    return match.group(0)
                keys.append(key)
        return "\\cite{" + ",".join(keys) + "}"

    # Parenthetical citations, including multiple authors separated by semicolons
    # and repeated years for the same first author.
    text = re.sub(r"\(([^()]+)\)", parenthetical, text)

    narrative = re.compile(
        r"\b([A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ][\wÁÀÂÃÉÊÍÓÔÕÚÜÇáàâãéêíóôõúüç’'.-]*"
        r"(?:\s+Jr\.)?"
        r"(?:\s*&\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ][\wÁÀÂÃÉÊÍÓÔÕÚÜÇáàâãéêíóôõúüç’'.-]*(?:\s+Jr\.)?"
        r"|\s+\*?et al\.\*?)?)\s*\(((?:19|20)\d{2})\)"
    )

    def narrative_repl(match: re.Match[str]) -> str:
        key = lookup_citation(match.group(1), match.group(2))
        return f"\\citeonline{{{key}}}" if key else match.group(0)

    return narrative.sub(narrative_repl, text)


def inline(text: str, crossrefs: bool = True) -> str:
    text = text.strip().replace("\\.", ".").replace("\\_", "_").replace("\\=", "=")
    text = re.sub(r"\\([()\[\],!<>])", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\(https?://[^)]+\)", r"\1", text)
    text = text.replace(
        "Comissão Internacional de Estratigrafia (2026)",
        r"Comissão Internacional de Estratigrafia \cite{ics2026}",
    )
    text = convert_citations(text)
    text, maths = protect_math(text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\\textbf{\1}", text)
    text = re.sub(r"\*([^*]+)\*", r"\\textit{\1}", text)
    if crossrefs:
        text = re.sub(r"\bFiguras\s+(\d+)\s+e\s+(\d+)", r"Figuras~\\ref{fig:resultado-\1} e \\ref{fig:resultado-\2}", text)
        text = re.sub(r"\bFigura\s+(\d+)", r"Figura~\\ref{fig:resultado-\1}", text)
        text = re.sub(r"\bTabelas\s+(\d+)\s+e\s+(\d+)", r"Tabelas~\\ref{tab:resultado-\1} e \\ref{tab:resultado-\2}", text)
        text = re.sub(r"\bTabela\s+(\d+)", r"Tabela~\\ref{tab:resultado-\1}", text)
        text = re.sub(r"\bQuadro\s+(\d+)", r"Quadro~\\ref{qua:resultado-\1}", text)
        text = text.replace("Apêndice D", r"Apêndice~\ref{ap:dados-morfometria}")
    placeholders = {}
    command_pattern = r"\\(?:cite|ref|textbf|textit)(?:online)?\{[^}]+\}|\\includegraphics(?:\[[^\]]+\])?\{[^}]+\}"
    for i, cmd in enumerate(re.findall(command_pattern, text)):
        key = f"@@CMD{i}@@"
        placeholders[key] = cmd
        text = text.replace(cmd, key, 1)
    text = text.replace("\\", "@@BS@@")
    text = text.replace("&", r"\&").replace("%", r"\%").replace("#", r"\#")
    text = text.replace("_", r"\_").replace("–", "--").replace("−", "$-$")
    text = text.replace("@@BS@@", "\\")
    for key, cmd in placeholders.items():
        text = text.replace(key, cmd)
    for i, math in enumerate(maths):
        text = text.replace(f"@@MATH{i}@@", math)
    return text


def image_targets(docx: Path) -> dict[int, tuple[str, bytes]]:
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
        relroot = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in relroot}
        found: list[tuple[str, dict[str, str]]] = []
        body = root.find(".//w:body", NS)
        if body is None:
            return {}
        for p in body.findall(".//w:p", NS):
            for drawing in p.findall(".//w:drawing", NS):
                blip = drawing.find(".//a:blip", NS)
                if blip is None:
                    continue
                rid = blip.attrib.get(f"{{{NS['r']}}}embed")
                if rid and rid in rels:
                    crop = drawing.find(".//a:srcRect", NS)
                    found.append((rels[rid], dict(crop.attrib) if crop is not None else {}))
        out = {}
        for i, (target, crop) in enumerate(found, 1):
            member = "word/" + target.replace("\\", "/")
            extension = Path(target).suffix.lower()
            data = zf.read(member)
            if crop and any(int(crop.get(side, "0")) for side in ("l", "t", "r", "b")):
                with Image.open(io.BytesIO(data)) as source:
                    width, height = source.size
                    left = round(width * int(crop.get("l", "0")) / 100000)
                    top = round(height * int(crop.get("t", "0")) / 100000)
                    right = round(width * (1 - int(crop.get("r", "0")) / 100000))
                    bottom = round(height * (1 - int(crop.get("b", "0")) / 100000))
                    cropped = source.crop((left, top, right, bottom))
                    buffer = io.BytesIO()
                    if extension in {".jpg", ".jpeg"}:
                        cropped.convert("RGB").save(buffer, format="JPEG", quality=95, subsampling=0)
                    else:
                        extension = ".png"
                        cropped.save(buffer, format="PNG")
                    data = buffer.getvalue()
            out[i] = (extension, data)
        return out


def image_targets_markdown(markdown: Path) -> dict[int, tuple[str, bytes]]:
    """Read Google Docs image definitions without relying on DOCX relationship order."""
    out: dict[int, tuple[str, bytes]] = {}
    pattern = re.compile(
        r"^\[image(\d+)\]:\s*<data:image/(png|jpe?g);base64,([^>]+)>.*$"
    )
    with markdown.open("r", encoding="utf-8") as source:
        for line in source:
            match = pattern.match(line.rstrip("\r\n"))
            if not match:
                continue
            number = int(match.group(1))
            subtype = match.group(2).lower()
            extension = ".jpg" if subtype in {"jpg", "jpeg"} else ".png"
            out[number] = (extension, base64.b64decode(match.group(3)))
    return out


def high_resolution_image_targets(docx: Path, markdown: Path) -> dict[int, tuple[str, bytes]]:
    """Match Markdown image IDs to their full-resolution DOCX drawings.

    Google Docs keeps image IDs in the Markdown, but the DOCX relationship
    order can contain an extra floating drawing. A small order-preserving
    image comparison finds that extra object and retains the correct mapping.
    """
    docx_images = image_targets(docx)
    markdown_images = image_targets_markdown(markdown)

    def feature(data: bytes) -> tuple[float, Image.Image]:
        with Image.open(io.BytesIO(data)) as source:
            ratio = source.width / source.height
            thumb = ImageOps.autocontrast(ImageOps.fit(source.convert("L"), (32, 32)))
        return ratio, thumb

    docx_features = {key: feature(value[1]) for key, value in docx_images.items()}
    markdown_features = {key: feature(value[1]) for key, value in markdown_images.items()}
    docx_keys = sorted(docx_images)
    markdown_keys = sorted(markdown_images)
    rows, columns = len(docx_keys), len(markdown_keys)
    infinity = float("inf")
    costs = [[infinity] * (columns + 1) for _ in range(rows + 1)]
    previous: list[list[tuple[str, int, int] | None]] = [[None] * (columns + 1) for _ in range(rows + 1)]
    costs[0][0] = 0.0

    def match_cost(docx_key: int, markdown_key: int) -> float:
        docx_ratio, docx_thumb = docx_features[docx_key]
        markdown_ratio, markdown_thumb = markdown_features[markdown_key]
        difference = ImageChops.difference(docx_thumb, markdown_thumb)
        mean_difference = ImageStat.Stat(difference).mean[0] / 255
        return 0.65 * mean_difference + 0.35 * abs(math.log(docx_ratio / markdown_ratio))

    for row in range(rows + 1):
        for column in range(columns + 1):
            current = costs[row][column]
            if row < rows and current + 0.35 < costs[row + 1][column]:
                costs[row + 1][column] = current + 0.35
                previous[row + 1][column] = ("skip-docx", row, column)
            if column < columns and current + 1.5 < costs[row][column + 1]:
                costs[row][column + 1] = current + 1.5
                previous[row][column + 1] = ("skip-markdown", row, column)
            if row < rows and column < columns:
                candidate = current + match_cost(docx_keys[row], markdown_keys[column])
                if candidate < costs[row + 1][column + 1]:
                    costs[row + 1][column + 1] = candidate
                    previous[row + 1][column + 1] = ("match", row, column)

    mapping: dict[int, tuple[str, bytes]] = {}
    row, column = rows, columns
    while row or column:
        step = previous[row][column]
        if step is None:
            break
        operation, prior_row, prior_column = step
        if operation == "match":
            mapping[markdown_keys[prior_column]] = docx_images[docx_keys[prior_row]]
        row, column = prior_row, prior_column
    # If Docs applied a visual crop that was not serialized as a:srcRect,
    # retain the Markdown display aspect ratio while keeping DOCX pixels.
    for key, (extension, data) in list(mapping.items()):
        with Image.open(io.BytesIO(data)) as source, Image.open(io.BytesIO(markdown_images[key][1])) as display:
            source_ratio = source.width / source.height
            display_ratio = display.width / display.height
            if abs(source_ratio - display_ratio) / max(source_ratio, display_ratio) <= 0.03:
                continue
            if source_ratio > display_ratio:
                target_width = round(source.height * display_ratio)
                left = (source.width - target_width) // 2
                cropped = source.crop((left, 0, left + target_width, source.height))
            else:
                target_height = round(source.width / display_ratio)
                top = (source.height - target_height) // 2
                cropped = source.crop((0, top, source.width, top + target_height))
            buffer = io.BytesIO()
            if extension in {".jpg", ".jpeg"}:
                cropped.convert("RGB").save(buffer, format="JPEG", quality=95, subsampling=0)
            else:
                extension = ".png"
                cropped.save(buffer, format="PNG")
            mapping[key] = (extension, buffer.getvalue())
    return mapping


def landscape_figure_numbers(docx: Path) -> set[int]:
    """Return figures that belong to landscape sections in the Docs DOCX.

    In WordprocessingML, section properties stored in a paragraph describe the
    section that ends at that paragraph. Grouping body paragraphs this way
    preserves the portrait/landscape choices made in Google Docs instead of
    relying on a hard-coded figure-number interval.
    """
    with zipfile.ZipFile(docx) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(".//w:body", NS)
    if body is None:
        return set()
    current: list[str] = []
    result: set[int] = set()

    def finish(section_properties: ET.Element | None) -> None:
        if section_properties is None:
            current.clear()
            return
        page_size = section_properties.find("w:pgSz", NS)
        orientation = page_size.attrib.get(f"{{{NS['w']}}}orient") if page_size is not None else None
        if orientation == "landscape":
            for paragraph in current:
                match = re.match(r"Figura\s+(\d+)\s*[—–-]", paragraph)
                if match:
                    result.add(int(match.group(1)))
        current.clear()

    for child in list(body):
        if child.tag == f"{{{NS['w']}}}p":
            current.append("".join(node.text or "" for node in child.findall(".//w:t", NS)))
            finish_at = child.find("w:pPr/w:sectPr", NS)
            if finish_at is not None:
                finish(finish_at)
        elif child.tag == f"{{{NS['w']}}}sectPr":
            finish(child)
    return result


def table_tex(rows: list[list[str]], caption: str | None, label: str | None, environment: str = "table") -> list[str]:
    if not rows:
        return []
    if environment == "quadro" and len(rows) > 4:
        header, body = rows[0], rows[1:]
        if len(header) == 5:
            column_spec = (
                r">{\raggedright\arraybackslash}p{0.10\linewidth} "
                r">{\raggedright\arraybackslash}p{0.14\linewidth} "
                r">{\raggedright\arraybackslash}p{0.18\linewidth} "
                r">{\raggedright\arraybackslash}X "
                r">{\centering\arraybackslash}p{0.18\linewidth}"
            )
        else:
            column_spec = (
                r">{\raggedright\arraybackslash}p{0.15\linewidth} "
                r">{\raggedright\arraybackslash}X "
                r">{\raggedright\arraybackslash}p{0.20\linewidth} "
                r">{\centering\arraybackslash}p{0.25\linewidth}"
            )
        midpoint = (len(body) + 1) // 2
        parts = [body[:midpoint], body[midpoint:]]
        out: list[str] = []
        for part_index, part in enumerate(parts):
            out += [r"\begin{landscape}", r"\begin{quadro}", r"  \centering"]
            if part_index == 0:
                if caption:
                    out.append(f"  \\caption{{{inline(caption, crossrefs=False)}}}")
                if label:
                    out.append(f"  \\label{{{label}}}")
            else:
                continuation = inline(caption or "", crossrefs=False)
                out.append(
                    f"  \\noindent\\textbf{{Quadro~\\thequadro{{}} — {continuation} (continuação).}}\\par\\smallskip"
                )
            out += [
                r"  \scriptsize",
                r"  \setlength{\tabcolsep}{3pt}",
                r"  \renewcommand{\arraystretch}{1.08}",
                f"  \\begin{{tabularx}}{{\\linewidth}}{{{column_spec}}}",
                r"    \toprule",
                "    " + " & ".join(f"\\textbf{{{inline(c, crossrefs=False)}}}" for c in header) + " \\\\",
                r"    \midrule",
            ]
            for row in part:
                cells = [inline(c, crossrefs=False) for c in row]
                cells = [
                    c.replace(r"\includegraphics[width=\linewidth]", r"\includegraphics[width=\linewidth,height=2.2cm,keepaspectratio]")
                    for c in cells
                ]
                out.append("    " + " & ".join(cells) + " \\\\")
            out += [
                r"    \bottomrule",
                r"  \end{tabularx}",
                r"\end{quadro}",
                r"\end{landscape}",
            ]
        return out
    if environment == "table" and len(rows) >= 12:
        cols = max(len(r) for r in rows)
        first_width = 0.20 if cols <= 6 else 0.16
        other_width = (0.90 - first_width) / max(cols - 1, 1)
        spec_parts = [rf">{{\raggedright\arraybackslash}}p{{{first_width:.3f}\linewidth}}"]
        spec_parts += [rf">{{\centering\arraybackslash}}p{{{other_width:.3f}\linewidth}}"] * (cols - 1)
        spec = "@{} " + " ".join(spec_parts) + " @{}"
        header = [inline(c, crossrefs=False) for c in rows[0]] + [""] * (cols - len(rows[0]))
        header_line = "    " + " & ".join(f"\\textbf{{{c}}}" for c in header) + r" \\"
        continuation = inline(caption or "", crossrefs=False)
        out = [
            r"{\scriptsize",
            r"\setlength{\tabcolsep}{2pt}",
            r"\renewcommand{\arraystretch}{1.08}",
            f"\\begin{{longtable}}{{{spec}}}",
        ]
        if caption:
            out.append(f"  \\caption{{{inline(caption, crossrefs=False)}}}" + (f"\\label{{{label}}}" if label else "") + r" \\")
        elif label:
            out.append(f"  \\label{{{label}}}\\")
        out += [r"    \toprule", header_line, r"    \midrule", r"    \endfirsthead"]
        if caption:
            out.append(f"  \\caption[]{{{continuation} (continuação)}} \\")
        out += [r"    \toprule", header_line, r"    \midrule", r"    \endhead"]
        out += [
            rf"    \midrule \multicolumn{{{cols}}}{{r}}{{\textit{{continua}}}} \\",
            r"    \endfoot",
            r"    \bottomrule",
            r"    \endlastfoot",
        ]
        for row in rows[1:]:
            cells = [inline(c, crossrefs=False) for c in row] + [""] * (cols - len(row))
            out.append("    " + " & ".join(cells) + r" \\")
        out += [r"\end{longtable}", r"}"]
        return out
    cols = max(len(r) for r in rows)
    spec = ">{\\raggedright\\arraybackslash}X " + " ".join([">{\\centering\\arraybackslash}X"] * (cols - 1))
    placement = "[htbp]" if environment == "table" else ""
    out = [f"\\begin{{{environment}}}{placement}", r"  \centering"]
    if caption:
        out.append(f"  \\caption{{{inline(caption, crossrefs=False)}}}")
    if label:
        out.append(f"  \\label{{{label}}}")
    out += [r"  \scriptsize", r"  \setlength{\tabcolsep}{2.5pt}", f"  \\begin{{tabularx}}{{\\textwidth}}{{{spec}}}", r"    \toprule"]
    for ri, row in enumerate(rows):
        cells = [inline(c, crossrefs=False) for c in row] + [""] * (cols - len(row))
        if ri == 0:
            cells = [f"\\textbf{{{c}}}" for c in cells]
            cells = [c.replace("Depósito/evento", "Depósito/\\newline evento") for c in cells]
        out.append("    " + " & ".join(cells) + r" \\")
        if ri == 0:
            out.append(r"    \midrule")
    out += [r"    \bottomrule", r"  \end{tabularx}", f"\\end{{{environment}}}"]
    return out


def parse_markdown(
    section: str,
    image_map: dict[int, tuple[str, bytes]],
    appendix: bool = False,
    chapter_number: str = "5",
    chapter_title: str = "Resultados",
) -> str:
    lines = section.splitlines()
    out: list[str] = []
    pending_caption: tuple[str, str, str] | None = None
    pending_image: int | None = None
    table_count = 0
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip()
        s = raw.strip()
        if not s:
            i += 1
            continue
        mimg = re.search(r"\[image(\d+)\]", s)
        if mimg:
            pending_image = int(mimg.group(1))
            i += 1
            continue
        mbold_heading = re.fullmatch(
            rf"\*\*({re.escape(chapter_number)}(?:\.\d+)+)\s+(.+?)\*\*",
            s,
        )
        if mbold_heading:
            depth = mbold_heading.group(1).count(".")
            title = inline(mbold_heading.group(2), crossrefs=False)
            command = "section" if depth == 1 else "subsection" if depth == 2 else "subsubsection"
            out += [f"\\{command}{{{title}}}", ""]
            i += 1
            continue
        if s.startswith("#"):
            level = len(s) - len(s.lstrip("#"))
            raw_title = re.sub(r"^#+\s*", "", s).strip()
            if raw_title.startswith("**") and raw_title.endswith("**"):
                raw_title = raw_title[2:-2].strip()
            title = clean_heading(s)
            chapter_heading = f"{chapter_number} {chapter_title.upper()}"
            is_appendix_heading = appendix and raw_title.upper().startswith("APÊNDICE D")
            is_acknowledgements = raw_title == "8 AGRADECIMENTOS"
            if not (is_appendix_heading or is_acknowledgements or raw_title == chapter_heading or re.match(rf"^{re.escape(chapter_number)}(?:\.\d+)*\s+\S", raw_title)):
                out += [inline(title), ""]
                i += 1
                continue
            if level == 1:
                if appendix:
                    title = re.sub(r"^APÊNDICE D\s*[—-]\s*", "", title, flags=re.IGNORECASE)
                    out += [f"\\chapter{{{inline(title, crossrefs=False)}}}\\label{{ap:dados-morfometria}}", ""]
                elif is_acknowledgements:
                    out += [r"\chapter{Agradecimentos}\label{cap:agradecimentos}", ""]
                else:
                    label = {
                        "5": "resultados",
                        "6": "discussao",
                        "7": "conclusoes",
                    }.get(chapter_number, slug(chapter_title))
                    out += [f"\\chapter{{{chapter_title}}}\\label{{cap:{label}}}", ""]
                    if chapter_number == "5":
                        out += [r"\setcounter{figure}{28}", ""]
            elif level == 2:
                out += [f"\\section{{{inline(title, crossrefs=False)}}}", ""]
            elif level == 3:
                out.append(f"\\subsection{{{inline(title, crossrefs=False)}}}")
                if title.casefold() == "velocidades utilizadas na análise hvsr".casefold():
                    out.append(r"\label{subsec:resultados-velocidades-hvsr}")
                out.append("")
            else:
                out += [f"\\subsubsection{{{inline(title, crossrefs=False)}}}", ""]
            i += 1
            continue
        mfig = re.match(r"Figura\s+(\d+)\s*[—–-]\s*(.*)", s)
        if mfig:
            num, caption = int(mfig.group(1)), mfig.group(2).replace("\\.", ".")
            figure_data = image_map.get(pending_image) if pending_image is not None else None
            # No DOCX, a Figura 80 é um objeto menor/flutuante ancorado no
            # parágrafo da legenda anterior e, por isso, não aparece no
            # Markdown exportado. Recupera-se diretamente a mídia correta.
            if figure_data is None and num == 80 and DOCX.exists():
                with zipfile.ZipFile(DOCX) as archive:
                    figure_data = (".png", archive.read("word/media/image1.png"))
            if figure_data is None:
                out += [
                    f"% PENDÊNCIA: imagem da Figura {num} não encontrada na exportação.",
                    r"\begin{figure}[htbp]",
                    r"  \centering",
                    r"  \fbox{\parbox[c][5cm][c]{0.9\textwidth}{\centering IMAGEM PENDENTE NO GOOGLE DOCS}}",
                    f"  \\caption{{{inline(caption, crossrefs=False)}}}",
                    f"  \\label{{fig:resultado-{num}}}",
                    r"\end{figure}",
                    "",
                ]
            else:
                ext, data = figure_data
                if ext not in {".png", ".jpg", ".jpeg", ".pdf"}:
                    ext = ".png"
                filename = f"figura-{num:02d}-{slug(caption)}{ext}"
                FIGDIR.mkdir(parents=True, exist_ok=True)
                (FIGDIR / filename).write_bytes(data)
                landscape_figure = num in LANDSCAPE_FIGURES
                if landscape_figure:
                    out.append(r"\begin{landscape}")
                out += [
                    r"\begin{figure}[p]" if landscape_figure else r"\begin{figure}[htbp]",
                    r"  \centering",
                    (f"  \\includegraphics[width=\\linewidth,height=0.80\\textheight,keepaspectratio]{{figuras/resultados/{filename}}}"
                     if landscape_figure else
                     f"  \\includegraphics[width=\\textwidth,height=0.82\\textheight,keepaspectratio]{{figuras/resultados/{filename}}}"),
                    f"  \\caption{{{inline(caption, crossrefs=False)}}}",
                    f"  \\label{{fig:resultado-{num}}}",
                    *([r"  \label{fig:velocidade-profundidade-poco}"] if num == 52 else []),
                    r"\end{figure}",
                ]
                if landscape_figure:
                    out.append(r"\end{landscape}")
                out.append("")
            pending_image = None
            i += 1
            continue
        mtab = re.match(r"(Tabela|Quadro)\s+((?:D\.)?\d+)\s+[—–-]\s*(.*)", s)
        if mtab:
            kind, n, cap = mtab.group(1), mtab.group(2), mtab.group(3).replace("\\.", ".")
            prefix = "tab" if kind == "Tabela" else "qua"
            environment = "table" if kind == "Tabela" else "quadro"
            pending_caption = (cap, f"{prefix}:resultado-{n.replace('.', '-')}", environment)
            i += 1
            continue
        if s.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    converted = []
                    for cell in cells:
                        def table_image(match):
                            image_number = int(match.group(1))
                            if image_number not in image_map:
                                return "[imagem não encontrada]"
                            ext, data = image_map[image_number]
                            if ext not in {".png", ".jpg", ".jpeg", ".pdf"}:
                                ext = ".png"
                            filename = f"quadro-imagem-{image_number:03d}{ext}"
                            FIGDIR.mkdir(parents=True, exist_ok=True)
                            (FIGDIR / filename).write_bytes(data)
                            return rf"\includegraphics[width=\linewidth]{{figuras/resultados/{filename}}}"
                        converted.append(re.sub(r"!\[[^\]]*\]\[image(\d+)\]", table_image, cell))
                    cells = converted
                    rows.append(cells)
                i += 1
            cap, lab, environment = pending_caption if pending_caption else (None, None, "table")
            out += table_tex(rows, cap, lab, environment) + [""]
            pending_caption = None
            table_count += 1
            continue
        if s.startswith("Nota:") or s.startswith("*Nota:") or s.startswith("*Nota específica:"):
            note = s.strip("*")
            note = re.sub(r"^Nota(?: específica)?:\s*", "", note)
            out += [f"\\nota{{{inline(note)}}}", ""]
            i += 1
            continue
        if s.startswith("PENDÊNCIA:"):
            out += [r"\noindent\colorbox{yellow}{\parbox{0.94\linewidth}{" + inline(s) + "}}", ""]
            i += 1
            continue
        if s.startswith("*\\*"):
            highlighted = s[3:-1] if s.endswith("*") else s[3:]
            out += [r"\noindent\colorbox{yellow}{\parbox{0.94\linewidth}{" + inline(highlighted) + "}}", ""]
            i += 1
            continue
        if s.lower().startswith("apresentar "):
            out += [f"\\TextoModelo{{{inline(s)}}}", ""]
            i += 1
            continue
        out += [inline(s), ""]
        i += 1
    return "\n".join(out).rstrip() + "\n"


def main() -> None:
    global LANDSCAPE_FIGURES
    if len(sys.argv) == 2:
        source_md = Path(sys.argv[1])
        MARKDOWN.write_text(source_md.read_text(encoding="utf-8"), encoding="utf-8")
    text = MARKDOWN.read_text(encoding="utf-8")
    # The Markdown export scales many Docs images down to roughly 605 px.
    # The DOCX keeps the embedded originals (commonly up to 2048 px), so it is
    # the authoritative image source for the LaTeX build.
    images = high_resolution_image_targets(DOCX, MARKDOWN)
    LANDSCAPE_FIGURES = landscape_figure_numbers(DOCX)
    resultados = between_headings(
        text,
        r"^#\s+\*\*5\s+RESULTADOS\*\*\s*$",
        r"^#\s+\*\*6\s+DISCUSS(?:ÃO|ÕES)\*\*\s*$",
    )
    discussao = between_headings(
        text,
        r"^#\s+\*\*6\s+DISCUSS(?:ÃO|ÕES)\*\*\s*$",
        r"^#\s+\*\*7\s+CONCLUSÕES\*\*\s*$",
    )
    conclusoes = between_headings(
        text,
        r"^#\s+\*\*7\s+CONCLUSÕES\*\*\s*$",
        r"^#\s+(?:\*\*)?Apêndice A-\s+Link repositório GitHub(?:\*\*)?\s*$",
    )
    apendice_d = between_headings(
        text,
        r"^#\s+(?:\*\*)?APÊNDICE D\b.*$",
        r"^#\s+(?:\*\*)?REFERÊNCIAS(?:\*\*)?\s*$",
    )
    RESULTADOS.write_text(parse_markdown(resultados, images), encoding="utf-8")
    DISCUSSAO.write_text(
        parse_markdown(discussao, images, chapter_number="6", chapter_title="Discussões"),
        encoding="utf-8",
    )
    CONCLUSOES.write_text(
        parse_markdown(conclusoes, images, chapter_number="7", chapter_title="Conclusões"),
        encoding="utf-8",
    )
    APENDICE_D.write_text(parse_markdown(apendice_d, images, appendix=True), encoding="utf-8")
    print(f"Resultados: {RESULTADOS}")
    print(f"Discussão: {DISCUSSAO}")
    print(f"Conclusões: {CONCLUSOES}")
    print(f"Apêndice D: {APENDICE_D}")


if __name__ == "__main__":
    main()
