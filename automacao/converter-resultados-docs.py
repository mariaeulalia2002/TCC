from __future__ import annotations

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
RESULTADOS = ROOT / "latex" / "capitulos" / "05-resultados.tex"
DISCUSSAO = ROOT / "latex" / "capitulos" / "06-discussao.tex"
APENDICE_D = ROOT / "latex" / "pos-textuais" / "apendice-d.tex"
FIGDIR = ROOT / "latex" / "figuras" / "resultados"

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
    text = text.strip("*").strip()
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


def inline(text: str, crossrefs: bool = True) -> str:
    text = text.strip().replace("\\.", ".").replace("\\_", "_")
    text = re.sub(r"\\([()\[\],!<>])", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\(https?://[^)]+\)", r"\1", text)
    text = text.replace("(Konno & Ohmachi, 1998)", r"\cite{konno1998}")
    text = text.replace("SESAME (2004)", r"SESAME \cite{sesame2004}")
    text = text.replace("(Vantassel et al., 2021)", r"\cite{vantassel2021}")
    text = text.replace("Ford et al. (2021)", r"\citeonline{ford2021}")
    text = text.replace("Ford et al. (2023)", r"\citeonline{ford2023}")
    text = text.replace("Hercos et al. (2023)", r"\citeonline{hercos2023}")
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


def table_tex(rows: list[list[str]], caption: str | None, label: str | None, environment: str = "table") -> list[str]:
    if not rows:
        return []
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
            raw_title = re.sub(r"^#+\s*", "", s).strip().strip("*").strip()
            title = clean_heading(s)
            chapter_heading = f"{chapter_number} {chapter_title.upper()}"
            if not (raw_title == chapter_heading or re.match(rf"^{re.escape(chapter_number)}(?:\.\d+)*\s+\S", raw_title)):
                out += [inline(title), ""]
                i += 1
                continue
            if level == 1:
                if appendix:
                    title = re.sub(r"^APÊNDICE D\s*[—-]\s*", "", title, flags=re.IGNORECASE)
                    out += [f"\\chapter{{{inline(title, crossrefs=False)}}}\\label{{ap:dados-morfometria}}", ""]
                else:
                    label = "resultados" if chapter_number == "5" else "discussao"
                    out += [f"\\chapter{{{chapter_title}}}\\label{{cap:{label}}}", ""]
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
            if pending_image is None or pending_image not in image_map:
                out.append(f"% PENDÊNCIA: imagem da Figura {num} não encontrada na exportação.")
            else:
                ext, data = image_map[pending_image]
                if ext not in {".png", ".jpg", ".jpeg", ".pdf"}:
                    ext = ".png"
                filename = f"figura-{num:02d}-{slug(caption)}{ext}"
                FIGDIR.mkdir(parents=True, exist_ok=True)
                (FIGDIR / filename).write_bytes(data)
                out += [
                    r"\begin{figure}[htbp]",
                    r"  \centering",
                    f"  \\includegraphics[width=\\textwidth,height=0.78\\textheight,keepaspectratio]{{figuras/resultados/{filename}}}",
                    f"  \\caption{{{inline(caption, crossrefs=False)}}}",
                    f"  \\label{{fig:resultado-{num}}}",
                    *([r"  \label{fig:velocidade-profundidade-poco}"] if num == 30 else []),
                    r"\end{figure}",
                    "",
                ]
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
    images = image_targets(DOCX)
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
    apendice_d = between_headings(
        text,
        r"^#\s+\*\*APÊNDICE D\b.*$",
        r"^#\s+(?:\*\*)?REFERÊNCIAS(?:\*\*)?\s*$",
    )
    RESULTADOS.write_text(parse_markdown(resultados, images), encoding="utf-8")
    DISCUSSAO.write_text(
        parse_markdown(discussao, images, chapter_number="6", chapter_title="Discussão"),
        encoding="utf-8",
    )
    APENDICE_D.write_text(parse_markdown(apendice_d, images, appendix=True), encoding="utf-8")
    print(f"Resultados: {RESULTADOS}")
    print(f"Discussão: {DISCUSSAO}")
    print(f"Apêndice D: {APENDICE_D}")


if __name__ == "__main__":
    main()
