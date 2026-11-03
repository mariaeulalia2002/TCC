from __future__ import annotations

import base64
import re
import shutil
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "docs" / "template-tcc-geologia-ufsc.docx"
MARKDOWN = ROOT / "docs" / "google-docs-latest.md"
RAW_MARKDOWN = ROOT / "tmp" / "google-docs-final-2026-11-03-pos-revisao.md"
RESULTADOS = ROOT / "latex" / "capitulos" / "05-resultados.tex"
DISCUSSAO = ROOT / "latex" / "capitulos" / "06-discussao.tex"
CONCLUSOES = ROOT / "latex" / "capitulos" / "07-conclusoes.tex"
APENDICE_D = ROOT / "latex" / "pos-textuais" / "apendice-d.tex"
FIGDIR = ROOT / "latex" / "figuras" / "resultados"
REFERENCIAS = ROOT / "latex" / "referencias.bib"

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
        r"\b([A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ][A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ’' .-]*?"
        r"(?:\s*&\s*[A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ][A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ’' .-]*|\s+\*?et al\.\*?)?)"
        r"\s*\(((?:19|20)\d{2})\)"
    )

    def narrative_repl(match: re.Match[str]) -> str:
        key = lookup_citation(match.group(1), match.group(2))
        return f"\\citeonline{{{key}}}" if key else match.group(0)

    return narrative.sub(narrative_repl, text)


def inline(text: str, crossrefs: bool = True) -> str:
    text = text.strip().replace("\\.", ".").replace("\\_", "_")
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
        found: list[str] = []
        body = root.find(".//w:body", NS)
        if body is None:
            return {}
        for p in body.findall(".//w:p", NS):
            for blip in p.findall(".//a:blip", NS):
                rid = blip.attrib.get(f"{{{NS['r']}}}embed")
                if rid and rid in rels:
                    found.append(rels[rid])
        out = {}
        for i, target in enumerate(found, 1):
            member = "word/" + target.replace("\\", "/")
            out[i] = (Path(target).suffix.lower(), zf.read(member))
        return out


def image_targets_markdown(markdown: Path) -> dict[int, tuple[str, bytes]]:
    """Read Google Docs image definitions without relying on DOCX relationship order."""
    out: dict[int, tuple[str, bytes]] = {}
    pattern = re.compile(
        r"^\[image(\d+)\]:\s*<data:image/(png|jpe?g);base64,([^>]+)>\s*$"
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


def table_tex(rows: list[list[str]], caption: str | None, label: str | None, environment: str = "table") -> list[str]:
    if not rows:
        return []
    if environment == "quadro" and len(rows) > 4:
        header, body = rows[0], rows[1:]
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
                r"  \begin{tabularx}{\linewidth}{>{\raggedright\arraybackslash}p{0.15\linewidth} >{\raggedright\arraybackslash}X >{\raggedright\arraybackslash}p{0.20\linewidth} >{\centering\arraybackslash}p{0.25\linewidth}}",
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
        mfig = re.match(r"Figura\s+(\d+)\s+[—–-]\s*(.*)", s)
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
                landscape_figure = 30 <= num <= 46
                if landscape_figure:
                    out.append(r"\begin{landscape}")
                out += [
                    r"\begin{figure}[p]" if landscape_figure else r"\begin{figure}[htbp]",
                    r"  \centering",
                    (f"  \\includegraphics[width=\\linewidth,height=0.70\\textheight,keepaspectratio]{{figuras/resultados/{filename}}}"
                     if landscape_figure else
                     f"  \\includegraphics[width=\\textwidth,height=0.78\\textheight,keepaspectratio]{{figuras/resultados/{filename}}}"),
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
    if len(sys.argv) == 2:
        source_md = Path(sys.argv[1])
        MARKDOWN.write_text(source_md.read_text(encoding="utf-8"), encoding="utf-8")
    text = MARKDOWN.read_text(encoding="utf-8")
    images = image_targets_markdown(RAW_MARKDOWN) if RAW_MARKDOWN.exists() else image_targets(DOCX)
    resultados = between_headings(
        text,
        r"^#\s+\*\*5\s+RESULTADOS\*\*\s*$",
        r"^#\s+\*\*6\s+DISCUSSÃO\*\*\s*$",
    )
    discussao = between_headings(
        text,
        r"^#\s+\*\*6\s+DISCUSSÃO\*\*\s*$",
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
        parse_markdown(discussao, images, chapter_number="6", chapter_title="Discussão"),
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
