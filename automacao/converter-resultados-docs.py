from __future__ import annotations

import base64
import io
import json
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
MARKDOWN = ROOT / "docs" / "google-docs-current.md"
IMAGE_MANIFEST = ROOT / "docs" / "google-docs-image-manifest.json"
IMAGE_CACHE = ROOT / "docs" / ".google-docs-images-current"
CHAPTERS = {
    "1": ROOT / "latex" / "capitulos" / "01-introducao.tex",
    "2": ROOT / "latex" / "capitulos" / "02-fundamentacao-teorica.tex",
    "3": ROOT / "latex" / "capitulos" / "03-area-estudo.tex",
    "4": ROOT / "latex" / "capitulos" / "04-materiais-metodos.tex",
    "5": ROOT / "latex" / "capitulos" / "05-resultados.tex",
    "6": ROOT / "latex" / "capitulos" / "06-discussao.tex",
    "7": ROOT / "latex" / "capitulos" / "07-conclusoes.tex",
}
RESULTADOS = ROOT / "latex" / "capitulos" / "05-resultados.tex"
DISCUSSAO = ROOT / "latex" / "capitulos" / "06-discussao.tex"
CONCLUSOES = ROOT / "latex" / "capitulos" / "07-conclusoes.tex"
APENDICES = {
    letter: ROOT / "latex" / "pos-textuais" / f"apendice-{letter.lower()}.tex"
    for letter in "ABCDEF"
}
FIGDIR = ROOT / "latex" / "figuras" / "google-docs"
RECOVERED_FIGURE_26 = (
    ROOT
    / "latex"
    / "figuras"
    / "metodologia"
    / "figura-28-altura-inclinacao-cicatriz.png"
)
REFERENCIAS = ROOT / "latex" / "referencias.bib"
LANDSCAPE_FIGURES: set[int] = set()
LANDSCAPE_QUADROS: set[str] = {"qua:resultado-1"}

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# Conteúdo matemático confirmado no texto. A API do Google Docs informa a
# posição de objetos de equação, mas não expõe a expressão interna; por isso a
# sincronização usa esta tabela explícita em vez de tentar adivinhar símbolos.
EQUATIONS: dict[int, tuple[str, str, str]] = {
    1: (r"C(t)=\cos\bigl(\phi(t)\bigr)", "eq:cosseno-fase", "Cosseno da fase instantânea"),
    2: (r"FS=\frac{\tau_f}{\tau_m}", "eq:fator-seguranca", "Fator de segurança"),
    3: (r"\tau_f=c'+\sigma'_n\tan\varphi'", "eq:mohr-coulomb", "Critério de ruptura de Mohr--Coulomb"),
    4: (r"\sigma'_n=\sigma_n-u", "eq:tensao-efetiva", "Tensão normal efetiva"),
    5: (r"\tau_f=c'+(\sigma_n-u)\tan\varphi'", "eq:mohr-coulomb-poropressao", "Resistência ao cisalhamento em função da poropressão"),
    6: (r"FS=\frac{c'+(\sigma_n-u)\tan\varphi'}{\tau_m}", "eq:fator-seguranca-completo", "Fator de segurança em função da poropressão"),
    7: (r"\Delta t(x,y)=t_{\mathrm{topo}}(x,y)-t_{\mathrm{base}}(x,y)", "eq:espessura-temporal", "Espessura temporal da isócrona"),
    8: (r"t_t=t_C-t_A", "eq:diferenca-altura-total", "Diferença temporal da altura total de queda"),
    9: (r"H_t=\lvert z_C-z_A\rvert", "eq:altura-total", "Altura total de queda"),
    10: (r"H_s=z_{B_s}-z_{A_s}", "eq:altura-cicatriz", "Altura da cicatriz"),
    11: (r"S_s=\arctan\!\left(\frac{\lvert z_Q-z_P\rvert}{L_{P-Q}}\right)\frac{180}{\pi}", "eq:inclinacao-cicatriz", "Inclinação máxima da cicatriz"),
    12: (r"S=\arctan\!\left(\frac{\lvert z_{P2}-z_{P1}\rvert}{L_{P1-P2}}\right)\frac{180}{\pi}", "eq:inclinacao-declive", "Inclinação do declive não deformado"),
    13: (r"S_t=\arctan\!\left(\frac{\lvert z_{P4}-z_{P3}\rvert}{L_{P3-P4}}\right)\frac{180}{\pi}", "eq:inclinacao-superficie-receptora", "Inclinação da superfície receptora"),
    14: (r"\Delta t_i=\frac{2000\,\Delta z_i}{\overline{v}_i}", "eq:acrescimo-twt", "Acréscimo de TWT no modelo de velocidade do poço 137"),
    15: (r"v_{\mathrm{int}}=\frac{2000(z_2-z_1)}{t_2-t_1}", "eq:velocidade-intervalar", "Velocidade intervalar no poço 127"),
    16: (r"\overline{v}_e=\frac{\sum_i v_{i,e}/d_{i,e}}{\sum_i 1/d_{i,e}}", "eq:velocidade-ponderada", "Velocidade ponderada pelo inverso da distância"),
    17: (r"h=\frac{\lvert t_{\mathrm{base}}-t_{\mathrm{topo}}\rvert\,\overline{v}_e}{2000}", "eq:conversao-espessura", "Conversão da espessura temporal em profundidade"),
    18: (r"z_{\mathrm{agua}}=\frac{1500\,t_{\mathrm{agua}}}{2000}", "eq:conversao-agua", "Conversão do TWT na coluna d'água em profundidade"),
    19: (r"z_P=\frac{1500\,t_{\mathrm{fundo},P}+\overline{v}_e(t_P-t_{\mathrm{fundo},P})}{2000}", "eq:profundidade-ponto", "Profundidade de ponto soterrado"),
    20: (r"u_{\mathrm{pick}}=\frac{s_d}{\sqrt{2}}", "eq:repetibilidade", "Incerteza de repetibilidade da interpretação"),
    21: (r"u_{\mathrm{interp},j}=\sqrt{K_j}", "eq:interpolacao", "Incerteza de interpolação"),
    22: (r"u_v=\sqrt{\sum_i p_i(v_i-\overline{v}_e)^2}", "eq:dispersao-velocidade", "Dispersão ponderada da velocidade"),
    23: (r"u_F=\sqrt{\sum_i\left(\frac{\partial F}{\partial x_i}\right)^2u_{x_i}^2}", "eq:propagacao-geral", "Propagação geral de incertezas"),
    24: (r"u_{t,\mathrm{esp}}=\sqrt{2u_{\mathrm{pick}}^2+K_{\mathrm{topo}}+K_{\mathrm{base}}}", "eq:incerteza-temporal-espessura", "Incerteza temporal da espessura"),
    25: (r"u_{t,\mathrm{sot}}=\sqrt{u_{\mathrm{pick}}^2+K_{\mathrm{topo}}+u_f^2}", "eq:incerteza-temporal-soterramento", "Incerteza temporal do soterramento"),
    26: (r"u_z=\sqrt{\left(\frac{\overline{v}_e u_t}{2000}\right)^2+\left(\frac{t u_v}{2000}\right)^2}", "eq:incerteza-profundidade", "Incerteza da medida convertida em profundidade"),
    27: (r"u_I=A\,u_{t,\mathrm{esp}}", "eq:incerteza-integral-temporal", "Incerteza da integral temporal do volume"),
    28: (r"u_V=\sqrt{\left(\frac{\overline{v}_e u_I}{2\times10^{12}}\right)^2+\left(\frac{Iu_v}{2\times10^{12}}\right)^2}", "eq:incerteza-volume", "Incerteza do volume"),
    29: (r"u_{\Delta z}=\sqrt{\left(\frac{\overline{v}_e u_{\Delta t_h}}{2000}\right)^2+\left(\frac{(1500-\overline{v}_e)u_{\Delta t_f}}{2000}\right)^2+\left(\frac{(\Delta t_h-\Delta t_f)u_v}{2000}\right)^2}", "eq:incerteza-desnivel", "Incerteza do desnível"),
    30: (r"u_\theta=\frac{180}{\pi}\frac{u_{\Delta z}}{L\left[1+(\Delta z/L)^2\right]}", "eq:incerteza-inclinacao", "Incerteza da inclinação"),
    31: (r"R_{\mathrm{TWT}}=\frac{500}{f_{\mathrm{dom}}}", "eq:resolucao-vertical", "Resolução vertical temporal"),
    32: (r"\operatorname{HVSR}(f)=\frac{\sqrt{\left([NS(f)]^2+[EW(f)]^2\right)/2}}{V(f)}", "eq:hvsr", "Razão espectral horizontal/vertical"),
    33: (r"H\approx\frac{V_s}{4f_0}", "eq:quarto-comprimento", "Profundidade da interface pelo modelo de quarto de comprimento de onda"),
    34: (r"f_{\pm1\sigma}=\widetilde{f}_0\exp\!\left[\pm\sigma_{\ln(f_0)}\right]", "eq:intervalo-lognormal", "Limites lognormais de frequência"),
    35: (r"V_s=V_p\sqrt{\frac{1-2\nu}{2(1-\nu)}}", "eq:vp-vs", "Conversão de $V_p$ para $V_s$ pelo coeficiente de Poisson"),
    36: (r"H_c=\frac{H_{\min}+H_{\max}}{2}", "eq:ponto-medio-envelope", "Ponto médio do envelope de profundidade"),
    37: (r"\Delta H_{\mathrm{abs}}=\frac{H_{\max}-H_{\min}}{2}", "eq:semi-amplitude-envelope", "Semi-amplitude do envelope de profundidade"),
    38: (r"\Delta H_{\mathrm{rel}}(\%)=\frac{\Delta H_{\mathrm{abs}}}{H_c}\times100", "eq:semi-amplitude-relativa", "Semi-amplitude relativa do envelope de profundidade"),
}
EQUATION_LABELS = {number: values[1] for number, values in EQUATIONS.items()}


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
    # A exportação do Google Docs pode fragmentar o negrito/itálico do título
    # nas últimas letras (por exemplo, ``poço***s***``). O destaque do título é
    # definido pela hierarquia no LaTeX, portanto esses marcadores não devem
    # ser interpretados como ênfase interna.
    text = text.replace("*", "").replace("_", "")
    text = re.sub(r"^(?:\d+(?:\.\d+)*|[A-F]\.\d+(?:\.\d+)*)\s+", "", text)
    return text.strip()


def slug(text: str) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    text = re.sub(r"[^a-zA-Z0-9]+", "-", text).strip("-").lower()
    return text[:48]


def short_caption(text: str) -> str:
    """Retorna somente o titulo para as listas de figuras e tabelas."""
    value = text.strip()
    panel = re.search(r"(?:\.\s*|:\s*|\s+)\([a-z]\)\s+", value, re.IGNORECASE)
    if panel and panel.start() >= 12:
        value = value[:panel.start()]
    else:
        sentence = re.search(r"\.\s+(?=[A-ZÁÀÂÃÉÊÍÓÔÕÚÜÇ])", value)
        if sentence and sentence.start() >= 12:
            value = value[:sentence.start()]
    if len(value) > 150:
        cut = max(value.rfind(mark, 0, 145) for mark in (":", ";", ","))
        if cut >= 40:
            value = value[:cut]
    return value.rstrip(" .;:")


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
        ("SMITH", "2011"): "smith2011",
        ("RIBEIRO", "2001"): "severianoribeiro2001",
        ("ORTIZ-KARPF", "2016"): "ortizkarpf2016",
        ("ORTIZ KARPF", "2016"): "ortizkarpf2016",
        ("QGIS DEVELOPMENT TEAM", "2025"): "qgis2025",
        ("NOAA", "2024"): "noaa2024",
        ("NATIONAL OCEANIC AND ATMOSPHERIC ADMINISTRATION", "2024"): "noaa2024",
        ("DAVIDOVICI", "1985"): "davidovici1985",
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
        postnote: str | None = None
        current_author: str | None = None
        for part in [p.strip() for p in content.split(";")]:
            item = re.fullmatch(
                r"(.+?),\s*((?:19|20)\d{2})(?:,\s*((?:19|20)\d{2}))?(?:,\s*(.+))?",
                part,
            )
            if not item:
                return match.group(0)
            current_author = item.group(1).strip()
            years = [item.group(2)] + ([item.group(3)] if item.group(3) else [])
            if item.group(4):
                if postnote is not None or len(content.split(";")) > 1:
                    return match.group(0)
                postnote = item.group(4).strip()
            for year in years:
                key = lookup_citation(current_author, year)
                if not key:
                    return match.group(0)
                keys.append(key)
        optional = f"[{postnote.replace(' ', '~')}]" if postnote else ""
        return "\\cite" + optional + "{" + ",".join(keys) + "}"

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


def convert_object_references(text: str) -> str:
    """Transforma chamadas numéricas do Docs em referências cruzadas LaTeX."""

    pattern = re.compile(
        r"\b(Figuras?|Tabelas?|Quadros?|Equação|Equações)\s+"
        r"((?:[A-F]\.)?\d+(?:[.\-]?[a-z])?"
        r"(?:(?:\s*(?:,|e|a|--|–|-)\s*)"
        r"(?:(?:[A-F]\.)?\d+(?:[.\-]?[a-z])?|[a-z]))*)"
    )

    def replacement(match: re.Match[str]) -> str:
        kind, sequence = match.group(1), match.group(2)
        sequence = re.sub(r"((?:[A-F]\.)?\d+)[.\-]([a-z])", r"\1\2", sequence)

        def one_ref(item: re.Match[str]) -> str:
            raw_number, panel = item.group(1), item.group(2) or ""
            label_number = raw_number.replace(".", "-")
            if kind.startswith("Figura"):
                command = rf"\ref{{fig:resultado-{label_number}}}"
            elif kind.startswith("Tabela"):
                command = rf"\ref{{tab:resultado-{label_number}}}"
            elif kind.startswith("Quadro"):
                command = rf"\ref{{qua:resultado-{label_number}}}"
            else:
                label = EQUATION_LABELS.get(int(raw_number)) if raw_number.isdigit() else None
                command = rf"\eqref{{{label}}}" if label else raw_number
            return command + (f"({panel})" if panel else "")

        converted = re.sub(r"\b((?:[A-F]\.)?\d+)([a-z])?", one_ref, sequence)
        return f"{kind}~{converted}"

    return pattern.sub(replacement, text)


def inline(text: str, crossrefs: bool = True) -> str:
    text = text.strip().replace("\\.", ".").replace("\\_", "_").replace("\\=", "=")
    # A API do Google Docs ainda devolve alguns objetos de equação nativa sem
    # o conteúdo da fórmula. Restaura os símbolos em linha a partir das
    # próprias equações numeradas e das definições mantidas no documento.
    inline_equation_repairs = {
        "em que  é o acréscimo de TWT no intervalo , em milissegundos;  é a sua espessura , em metros; e é a média das velocidades nos dois extremos do intervalo.":
            r"em que $\Delta t_i$ é o acréscimo de TWT no intervalo $i$, em milissegundos; $\Delta z_i$ é a sua espessura, em metros; e $\overline{v}_i$ é a média das velocidades nos dois extremos do intervalo.",
        "Em que h é a espessura, em metros; t<sub>base</sub> e t<sub>topo</sub> são os tempos duplos de percurso da <sub>base</sub> e do <sub>topo</sub>, em milissegundos; e v̄ₑ é a velocidade ponderada do evento":
            r"Em que $h$ é a espessura, em metros; $t_{\mathrm{base}}$ e $t_{\mathrm{topo}}$ são os tempos duplos de percurso da base e do topo, em milissegundos; e $\overline{v}_e$ é a velocidade ponderada do evento",
        "Em que  é a profundidade em metros e , o TWT em milissegundos.":
            r"Em que $z_{\mathrm{agua}}$ é a profundidade em metros e $t_{\mathrm{agua}}$, o TWT em milissegundos.",
        "Em que  é a profundidade do ponto interpretado, em metros;  o tempo do ponto interpretado; e , o tempo do fundo marinho na mesma posição, ambos expressos como TWT positivo em milissegundos; e  é a velocidade ponderada do evento":
            r"Em que $z_P$ é a profundidade do ponto interpretado, em metros; $t_P$, o tempo do ponto interpretado; e $t_{\mathrm{fundo},P}$, o tempo do fundo marinho na mesma posição, ambos expressos como TWT positivo em milissegundos; e $\overline{v}_e$ é a velocidade ponderada do evento",
        "em que  é o desvio-padrão das diferenças entre as duas interpretações, em milissegundos. Nos cálculos, utilizou-se  ms":
            r"em que $s_d$ é o desvio-padrão das diferenças entre as duas interpretações, em milissegundos. Nos cálculos, utilizou-se $u_{\mathrm{pick}}=0{,}38$ ms",
        "em que  é a variância média de krigagem da superfície , em ms², e  é sua componente temporal, em milissegundos.":
            r"em que $K_j$ é a variância média de krigagem da superfície $j$, em ms², e $u_{\mathrm{interp},j}$ é sua componente temporal, em milissegundos.",
        "Para o fundo marinho, adotou-se  ms":
            r"Para o fundo marinho, adotou-se $u_f=0{,}14$ ms",
        "em que  é a velocidade equivalente no poço ,  é a velocidade ponderada do evento e  é o peso normalizado do poço.":
            r"em que $v_i$ é a velocidade equivalente no poço $i$, $\overline{v}_e$ é a velocidade ponderada do evento e $p_i$ é o peso normalizado do poço.",
        "em que  é a medida calculada,  são suas variáveis de entrada e  são as respectivas componentes de incerteza.":
            r"em que $F$ é a medida calculada, $x_i$ são suas variáveis de entrada e $u_{x_i}$ são as respectivas componentes de incerteza.",
        "em que  e  são as variâncias médias de krigagem, em ms², e  é a componente temporal do fundo marinho":
            r"em que $K_{\mathrm{topo}}$ e $K_{\mathrm{base}}$ são as variâncias médias de krigagem, em ms², e $u_f$ é a componente temporal do fundo marinho",
        "Para uma medida convertida pela relação , a incerteza em metros":
            r"Para uma medida convertida pela relação $z=\overline{v}_e t/2000$, a incerteza em metros",
        "em que  é o intervalo de TWT utilizado no cálculo, em milissegundos, e  corresponde a  ou , conforme a medida.":
            r"em que $t$ é o intervalo de TWT utilizado no cálculo, em milissegundos, e $u_t$ corresponde a $u_{t,\mathrm{esp}}$ ou $u_{t,\mathrm{sot}}$, conforme a medida.",
        "conforme a expressão de  apresentada anteriormente.":
            r"conforme a expressão de $u_{t,\mathrm{esp}}$ apresentada anteriormente.",
        "em que  é a área do depósito, em metros quadrados, e  é a componente de incerteza da integral temporal":
            r"em que $A$ é a área do depósito, em metros quadrados, e $u_I$ é a componente de incerteza da integral temporal",
        "em que  é a integral da espessura temporal, em ms·m²;  é a velocidade ponderada do evento; e  é sua dispersão ponderada":
            r"em que $I$ é a integral da espessura temporal, em ms·m²; $\overline{v}_e$ é a velocidade ponderada do evento; e $u_v$ é sua dispersão ponderada",
        "O fator  reúne a conversão do tempo duplo em milissegundos e a expressão do volume em quilômetros cúbicos.":
            r"O fator $2\times10^{12}$ reúne a conversão do tempo duplo em milissegundos e a expressão do volume em quilômetros cúbicos.",
        "em que  é a diferença de TWT entre os pontos do horizonte interpretado e  é a diferença de TWT do fundo marinho nas mesmas posições, ambas em milissegundos e calculadas na mesma ordem entre os extremos;  é a velocidade sedimentar adotada para o evento, em metros por segundo; e  é sua dispersão ponderada. As componentes temporais foram definidas por  e , considerando independência entre os extremos.":
            r"em que $\Delta t_h$ é a diferença de TWT entre os pontos do horizonte interpretado e $\Delta t_f$ é a diferença de TWT do fundo marinho nas mesmas posições, ambas em milissegundos e calculadas na mesma ordem entre os extremos; $\overline{v}_e$ é a velocidade sedimentar adotada para o evento, em metros por segundo; e $u_v$ é sua dispersão ponderada. As componentes temporais foram definidas por $u_{\Delta t_h}=\sqrt{2}\;u_{\mathrm{pick}}$ e $u_{\Delta t_f}=\sqrt{2}\;u_f$, considerando independência entre os extremos.",
        "Para a altura total de queda, adotou-se .":
            r"Para a altura total de queda, adotou-se $\Delta t_h=t_C-t_A$ e $\Delta t_f=t_{\mathrm{fundo},C}-t_{\mathrm{fundo},A}$.",
        "Para as inclinações, calculadas por , a propagação foi realizada mantendo a distância horizontal  fixa:":
            r"Para as inclinações, calculadas por $\theta=\arctan(|\Delta z|/L)$, a propagação foi realizada mantendo a distância horizontal $L$ fixa:",
        "em que  é a incerteza angular, em graus;  é a incerteza do desnível, em metros; e  é a distância horizontal entre os pontos, em metros. Essa expressão foi aplicada às inclinações do declive não deformado () e da superfície receptora ().":
            r"em que $u_\theta$ é a incerteza angular, em graus; $u_{\Delta z}$ é a incerteza do desnível, em metros; e $L$ é a distância horizontal entre os pontos, em metros. Essa expressão foi aplicada às inclinações do declive não deformado ($S$) e da superfície receptora ($S_t$).",
        "Adotou-se o critério de um quarto do comprimento de onda dominante, apresentado por":
            "Adotou-se o critério de um quarto do comprimento de onda dominante, apresentado por",
        "segundo o qual a resolução vertical em profundidade é estimada por . Expressando essa espessura em tempo duplo de percurso, pela relação , e convertendo segundos para milissegundos, obtém-se:":
            r"segundo o qual a resolução vertical em profundidade é estimada por $R_z=v/(4f_{\mathrm{dom}})$. Expressando essa espessura em tempo duplo de percurso, pela relação $R_{\mathrm{TWT}}=2000R_z/v$, e convertendo segundos para milissegundos, obtém-se:",
        "em que  é a frequência dominante, em hertz, e  é a resolução vertical temporal":
            r"em que $f_{\mathrm{dom}}$ é a frequência dominante, em hertz, e $R_{\mathrm{TWT}}$ é a resolução vertical temporal",
        "Nas anotações desses perfis, os limites superior e inferior foram chamados de A e B. Para distingui-los dos pontos do eixo longitudinal, são indicados aqui como  e :":
            r"Nas anotações desses perfis, os limites superior e inferior foram chamados de A e B. Para distingui-los dos pontos do eixo longitudinal, são indicados aqui como $A_s$ e $B_s$:",
        "Dentro do intervalo –, foram selecionados":
            r"Dentro do intervalo $A_s$--$B_s$, foram selecionados",
        "a inclinação máxima .":
            r"a inclinação máxima $S_s$.",
        "A inclinação do declive, , foi determinada":
            r"A inclinação do declive, $S$, foi determinada",
        "O parâmetro , identificado na tabela como inclinação da base":
            r"O parâmetro $S_t$, identificado na tabela como inclinação da base",
        "em que  e  representam os espectros de amplitude das componentes horizontais norte–sul e leste–oeste, respectivamente;  representa o espectro da componente vertical; e  é a frequência.":
            r"em que $NS(f)$ e $EW(f)$ representam os espectros de amplitude das componentes horizontais norte--sul e leste--oeste, respectivamente; $V(f)$ representa o espectro da componente vertical; e $f$ é a frequência.",
        "determinação da frequência fundamental  foram":
            r"determinação da frequência fundamental $f_0$ foram",
        "possíveis frequências secundárias  e , sem modificar a frequência fundamental  determinada":
            r"possíveis frequências secundárias $f_1$ e $f_2$, sem modificar a frequência fundamental $f_0$ determinada",
        "A frequência fundamental  foi mantida":
            r"A frequência fundamental $f_0$ foi mantida",
        "classificados como  e .":
            r"classificados como $f_1$ e $f_2$.",
        "Os valores mínimo () e máximo () delimitaram":
            r"Os valores mínimo ($H_{\min}$) e máximo ($H_{\max}$) delimitaram",
        "O pico secundário mais consistente ocorreu em torno de  Hz":
            r"O pico secundário mais consistente ocorreu em torno de $2{,}68$ Hz",
        "( Hz para 01/02/2020 e  Hz para 29/02/2020)":
            r"($2{,}57$ Hz para 01/02/2020 e $2{,}62$ Hz para 29/02/2020)",
        "motivo pelo qual  foi interpretado":
            r"motivo pelo qual $f_1$ foi interpretado",
        "classificado como , também":
            r"classificado como $f_2$, também",
        "indica que  deve ser interpretado":
            r"indica que $f_2$ deve ser interpretado",
    }
    for old, new in inline_equation_repairs.items():
        text = text.replace(old, new)
    text = re.sub(r"<br\s*/?>", r"\\newline ", text, flags=re.IGNORECASE)
    text = re.sub(r"<sub>(.*?)</sub>", r"$_{\1}$", text)
    text = re.sub(r"<sup>(.*?)</sup>", r"$^{\1}$", text)
    # O Google Docs exporta o nome do campo do QGIS como ``$length``. O
    # cifrão é parte do nome do campo, não um delimitador matemático. Tokens
    # temporários evitam que o protetor de expressões confunda esse texto com
    # o início de uma fórmula.
    text = text.replace("$length / 1000", "@@QLENGTHKM@@")
    text = text.replace("$length", "@@QLENGTH@@")
    text = re.sub(
        r"==(.+?)==",
        lambda match: r"\colorbox{yellow}{\strut " + match.group(1) + "}",
        text,
    )
    text = text.replace("Figura Iv.6", "Figura IV.6")
    text = text.replace("sísmicas.Para", "sísmicas. Para")
    text = text.replace("*et al*.", "*et al.*")
    text = text.replace("*et al.,*", "*et al.*,")
    text = re.sub(r"\*\s+et al\.\*", "*et al.*", text, flags=re.IGNORECASE)
    text = re.sub(r"(?<=\w)\*et al\.\*", " *et al.*", text, flags=re.IGNORECASE)
    text = re.sub(r"(\*?et al\.\*?),\s*\(", r"\1 (", text, flags=re.IGNORECASE)
    text = re.sub(r"\\([()\[\],!<>])", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\(https?://[^)]+\)", r"\1", text)
    text = text.replace(
        "Comissão Internacional de Estratigrafia (2026)",
        r"Comissão Internacional de Estratigrafia \cite{ics2026}",
    )
    text = text.replace(
        "(Güralp Systems Limited, s.d.)",
        r"\cite{guralp}",
    )
    text = text.replace(
        "Vail *et al.* (1987 apud Ribeiro, 2001)",
        r"\citeonline{vail1987} apud \citeonline{severianoribeiro2001}",
    )
    text = text.replace(
        "MOREIRA *et al.*, 2007, apud Diório (2024, p. 34)",
        r"\citeonline{moreira2007} apud \citeonline[p.~34]{diorio2024}",
    )
    text = text.replace(
        "Kearey, Brooks & Hill (2009)",
        r"\citeonline{kearey2009}",
    )
    text = text.replace(
        "Os três termos da Equação 27 representam",
        r"Os três termos da Equação~\eqref{eq:incerteza-desnivel} representam",
    )
    # Corrige duas chamadas herdadas da renumeração das figuras do MTD 1.
    text = text.replace("Na seção inline (Fig 30-a)", "Na seção inline (Figura 29a)")
    text = text.replace("Na seção crossline (Fig 30-d)", "Na seção crossline (Figura 29d)")
    text = text.replace("seção inline 11986 (Fig 30a--b)", "seção inline 11986 (Figura 29a--b)")
    text = text.replace("seção crossline 6656 (Fig 30d--e)", "seção crossline 6656 (Figura 29d--e)")
    text = text.replace("seção inline 11986 (Figura 30a–b)", "seção inline 11986 (Figura 29a–b)")
    text = text.replace("seção crossline 6656 (Figura 30d–e)", "seção crossline 6656 (Figura 29d–e)")
    text = convert_citations(text)
    text, maths = protect_math(text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\\textbf{\1}", text)
    text = re.sub(r"\*([^*]+)\*", r"\\textit{\1}", text)
    protected_italics: dict[str, str] = {}
    for match in re.finditer(r"\\textit\{[^{}]+\}", text):
        token = f"@@ITALIC{len(protected_italics)}@@"
        protected_italics[token] = match.group(0)
        text = text.replace(match.group(0), token, 1)
    english_terms = [
        "two-way travel time",
        "Post-Stack Time Migration",
        "Southern Ocean Current",
        "Horizontal-to-Vertical Spectral Ratio",
        "sediment bypassing",
        "en masse",
        "mass-transport deposit",
        "mass-transport complex",
        "Volume Below Surface",
        "Surfer Binary Grid",
        "Guided Autotracking",
        "Make Surface",
        "binary header",
        "trace headers",
        "bin size",
        "Operations",
        "Petrel",
        "Surfer",
        "Grapher",
        "QGIS",
        "Python",
        "hvsrweb",
        "hvsrpy",
        "Raster",
        "raster",
        "levee",
        "paleopockmarks",
        "pockmarks",
        "pockmark",
        "bright spot",
        "channel-levee",
        "Manual Interpretation",
        "loop tie",
        "mis-ties",
        "crossline",
        "crosslines",
        "inline",
        "inlines",
        "offshore",
        "rollovers",
        "rollover",
        "drifts",
        "drift",
        "drapes",
        "drape",
        "downlap",
        "toplap",
        "onlap",
        "onlapante",
        "slumping",
        "attached",
        "detached",
        "chirp",
        "notebooks",
        "notebook",
        "taper",
        "script",
        "software",
        "grid",
        "pool",
    ]
    for term in english_terms:
        text = re.sub(
            rf"(?<![\w@]){re.escape(term)}(?![\w@])",
            lambda match: rf"\textit{{{match.group(0)}}}",
            text,
            flags=re.IGNORECASE,
        )
    for token, value in protected_italics.items():
        text = text.replace(token, value)
    # Mantém o símbolo de marca registrada no mesmo destaque tipográfico do
    # nome do software, conforme o padrão adotado no texto do trabalho.
    text = re.sub(
        r"\\textit\{(Petrel|Surfer|Grapher|QGIS|Python)\}®",
        r"\\textit{\1®}",
        text,
    )
    text = re.sub(r"(?<![\w@])find_peaks(?![\w@])", r"\\texttt{find\_peaks}", text)
    if crossrefs:
        text = re.sub(r"\bFig(?:\.\s*|\s+)", "Figura ", text, flags=re.IGNORECASE)
        text = convert_object_references(text)
        appendix_labels = {
            "A": "ap:codigos",
            "B": "ap:localizacao-medidas",
            "C": "ap:medicoes-cicatrizes",
            "D": "ap:distancias-velocidades",
            "E": "ap:dados-morfometria",
            "F": "ap:hvsr",
        }
        for letter, label in appendix_labels.items():
            text = text.replace(f"Apêndice {letter}", rf"Apêndice~\ref{{{label}}}")
        text = re.sub(
            r"\bTabela\s+([A-F])\.(\d+)",
            r"Tabela~\\ref{tab:resultado-\1-\2}",
            text,
        )
        # O parágrafo do MTD 1 ainda conserva duas chamadas antigas à Figura
        # 30 no Docs; os painéis correspondentes pertencem à Figura 29.
        if text.startswith("O MTD 1, destacado em rosa"):
            text = text.replace("fig:resultado-30", "fig:resultado-29")
    placeholders = {}
    command_pattern = r"\\(?:cite|ref|textbf|textit|texttt)(?:online)?\{[^}]+\}|\\includegraphics(?:\[[^\]]+\])?\{[^}]+\}"
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
    text = text.replace("@@QLENGTHKM@@", r"\texttt{\$length / 1000}")
    text = text.replace("@@QLENGTH@@", r"\texttt{\$length}")
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


def api_image_targets(manifest_path: Path, cache_dir: Path) -> dict[int, tuple[str, bytes]]:
    """Load the current Docs images and apply the crop stored by Google Docs."""
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    result: dict[int, tuple[str, bytes]] = {}
    for item in manifest:
        object_id = item["objectId"]
        candidates = [
            path for path in cache_dir.glob(f"{object_id}.*")
            if path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp"}
        ]
        if not candidates:
            continue
        source_path = candidates[0]
        crop = item.get("crop") or {}
        with Image.open(source_path) as source:
            source.load()
            left = round(source.width * float(crop.get("offsetLeft", 0)))
            right = round(source.width * (1 - float(crop.get("offsetRight", 0))))
            top = round(source.height * float(crop.get("offsetTop", 0)))
            bottom = round(source.height * (1 - float(crop.get("offsetBottom", 0))))
            left = max(0, min(left, source.width - 1))
            right = max(left + 1, min(right, source.width))
            top = max(0, min(top, source.height - 1))
            bottom = max(top + 1, min(bottom, source.height))
            cropped = source.crop((left, top, right, bottom))
            buffer = io.BytesIO()
            if source_path.suffix.lower() in {".jpg", ".jpeg"}:
                extension = ".jpg"
                cropped.convert("RGB").save(
                    buffer, format="JPEG", quality=98, subsampling=0, optimize=True
                )
            else:
                extension = ".png"
                cropped.save(buffer, format="PNG", optimize=True)
        result[int(item["order"])] = (extension, buffer.getvalue())
    return result


def replace_gdoc_images(text: str, manifest_path: Path) -> str:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    by_id = {item["objectId"]: int(item["order"]) for item in manifest}

    def replacement(match: re.Match[str]) -> str:
        object_id = match.group(1)
        number = by_id.get(object_id)
        return f"![Imagem][image{number}]" if number is not None else match.group(0)

    return re.sub(
        r"!\[[^\]]*\]\(gdoc://([^)]+)\)",
        replacement,
        text,
    )


def landscape_figure_numbers_api(text: str) -> set[int]:
    result: set[int] = set()
    landscape = False
    for line in text.splitlines():
        section = re.match(r"<!-- SECTION (portrait|landscape)", line.strip())
        if section:
            landscape = section.group(1) == "landscape"
            continue
        if landscape:
            figure = re.match(r"Figura\s+(\d+)\s*[—–-]", line.strip())
            if figure:
                result.add(int(figure.group(1)))
    return result


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


def table_tex(
    rows: list[list[str]],
    caption: str | None,
    label: str | None,
    environment: str = "table",
    note: str | None = None,
) -> list[str]:
    if not rows:
        return []
    is_appendix_table = bool(label and re.fullmatch(r"tab:resultado-[A-F]-\d+", label))
    # O Quadro 2 tem apenas cinco linhas de conteúdo e cabe inteiro em uma
    # página. Mantê-lo unido evita um título de continuação desnecessário.
    if (
        environment == "quadro"
        and len(rows) > 4
        and label not in {"qua:resultado-2", "qua:resultado-A-1"}
    ):
        quadro_landscape = bool(label and label in LANDSCAPE_QUADROS)
        header, body = rows[0], rows[1:]
        if len(header) == 5:
            column_spec = (
                r"|>{\raggedright\arraybackslash}p{0.10\linewidth}"
                r"|>{\raggedright\arraybackslash}p{0.14\linewidth}"
                r"|>{\raggedright\arraybackslash}p{0.18\linewidth}"
                r"|>{\raggedright\arraybackslash}X"
                r"|>{\centering\arraybackslash}p{0.18\linewidth}|"
            )
        else:
            if label == "qua:resultado-1":
                # O último campo contém os exemplos sísmicos. Dar-lhe mais
                # largura padroniza visualmente as imagens e impede que elas
                # cubram as linhas verticais do quadro.
                column_spec = (
                    r"|>{\raggedright\arraybackslash}p{0.13\linewidth}"
                    r"|>{\raggedright\arraybackslash}X"
                    r"|>{\raggedright\arraybackslash}p{0.17\linewidth}"
                    r"|>{\centering\arraybackslash}p{0.34\linewidth}|"
                )
            else:
                column_spec = (
                    r"|>{\raggedright\arraybackslash}p{0.15\linewidth}"
                    r"|>{\raggedright\arraybackslash}X"
                    r"|>{\raggedright\arraybackslash}p{0.20\linewidth}"
                    r"|>{\centering\arraybackslash}p{0.25\linewidth}|"
                )
        midpoint = (len(body) + 1) // 2
        parts = [body[:midpoint], body[midpoint:]]
        out: list[str] = []
        if quadro_landscape:
            # Um único ambiente paisagem para as duas partes evita a página
            # em branco que surgia ao fechar e reabrir ``landscape``.
            out.append(r"\begin{landscape}")
        for part_index, part in enumerate(parts):
            out += [r"\begin{quadro}", r"  \centering"]
            if part_index == 0:
                if caption:
                    out.append(f"  \\caption[{inline(short_caption(caption))}]{{{inline(caption)}}}")
                if label:
                    out.append(f"  \\label{{{label}}}")
            else:
                continuation = inline(caption or "", crossrefs=False)
                out.append(
                    f"  \\begin{{center}}\\footnotesize Quadro~\\thequadro{{}} — {continuation} (continuação).\\end{{center}}\\smallskip"
                )
            out += [
                r"  \small",
                r"  \setlength{\tabcolsep}{3pt}",
                r"  \renewcommand{\arraystretch}{1.08}",
                f"  \\begin{{tabularx}}{{\\linewidth}}{{{column_spec}}}",
                r"    \hline",
                "    " + " & ".join(f"\\textbf{{{inline(c, crossrefs=False)}}}" for c in header) + " \\\\",
                r"    \hline",
            ]
            for row in part:
                cells = [inline(c, crossrefs=False) for c in row]
                if label == "qua:resultado-1":
                    cells = [
                        c.replace(r"\includegraphics[width=\linewidth]", r"\includegraphics[width=0.98\linewidth,height=2.05cm,keepaspectratio]")
                        for c in cells
                    ]
                else:
                    cells = [
                        c.replace(r"\includegraphics[width=\linewidth]", r"\includegraphics[width=\linewidth,height=2.05cm,keepaspectratio]")
                        for c in cells
                    ]
                # Na configuração interna 4, o recorte do cosseno de fase
                # possui proporção mais baixa que o recorte de amplitude e
                # ficava visualmente menor. Um acréscimo discreto equaliza os
                # dois exemplos sem alterar as demais imagens do quadro.
                if label == "qua:resultado-1":
                    # O exemplo de cosseno da configuração 4 tem proporção
                    # mais baixa; o pequeno acréscimo o aproxima visualmente
                    # do exemplo de amplitude, como solicitado na revisão.
                    cells = [
                        c.replace(
                            r"height=2.05cm,keepaspectratio]{figuras/google-docs/quadro-imagem-050.png}",
                            r"height=2.2cm,keepaspectratio]{figuras/google-docs/quadro-imagem-050.png}",
                        )
                        for c in cells
                    ]
                else:
                    cells = [
                        c.replace(
                            r"height=2.05cm,keepaspectratio]{figuras/google-docs/quadro-imagem-050.png}",
                            r"height=2.2cm,keepaspectratio]{figuras/google-docs/quadro-imagem-050.png}",
                        )
                        for c in cells
                    ]
                out.append("    " + " & ".join(cells) + r" \\ \hline")
            out += [
                r"  \end{tabularx}",
                *([f"  \\nota{{{inline(note)}}}"] if note and part_index == len(parts) - 1 else []),
                r"\end{quadro}",
            ]
        if quadro_landscape:
            out.append(r"\end{landscape}")
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
            r"{\small",
            r"\setlength{\tabcolsep}{2pt}",
            r"\renewcommand{\arraystretch}{1.08}",
            f"\\begin{{longtable}}{{{spec}}}",
        ]
        if caption:
            out.append(f"  \\caption[{inline(short_caption(caption))}]{{{inline(caption)}}}" + (f"\\label{{{label}}}" if label else "") + r" \\")
        elif label:
            out.append(f"  \\label{{{label}}}\\")
        out += [r"    \toprule", header_line, r"    \midrule", r"    \endfirsthead"]
        if caption:
            out.append(f"  \\caption[]{{{continuation} (continuação)}}" + r" \\")
        out += [r"    \toprule", header_line, r"    \midrule", r"    \endhead"]
        out += [
            rf"    \midrule \multicolumn{{{cols}}}{{r}}{{\textit{{continua}}}} \\",
            r"    \endfoot",
            r"    \bottomrule",
            r"    \endlastfoot",
        ]
        for row_index, row in enumerate(rows[1:]):
            cells = [inline(c, crossrefs=False) for c in row] + [""] * (cols - len(row))
            if label in {"tab:resultado-F-1", "tab:resultado-F-2"} and 6 <= row_index <= 11:
                cells = [f"\\textbf{{{cell}}}" for cell in cells]
            out.append("    " + " & ".join(cells) + r" \\")
            if is_appendix_table and row_index < len(rows[1:]) - 1:
                out.append(r"    \midrule")
        out.append(r"\end{longtable}")
        if note:
            out.append(f"\\nota{{{inline(note)}}}")
        out.append(r"}")
        return out
    cols = max(len(r) for r in rows)
    if environment == "quadro" and label == "qua:resultado-A-1":
        spec = (
            r"|>{\raggedright\arraybackslash}p{0.34\textwidth}"
            r"|>{\raggedright\arraybackslash}X|"
        )
    elif environment == "quadro":
        spec = "|>{\\raggedright\\arraybackslash}X|" + "".join([">{\\centering\\arraybackslash}X|"] * (cols - 1))
    else:
        spec = ">{\\raggedright\\arraybackslash}X " + " ".join([">{\\centering\\arraybackslash}X"] * (cols - 1))
    force_here = bool(label and re.fullmatch(r"tab:resultado-[5-7]", label))
    placement = "[H]" if environment == "table" and (is_appendix_table or force_here) else "[htbp]" if environment == "table" else ""
    out = [f"\\begin{{{environment}}}{placement}", r"  \centering"]
    if caption:
        out.append(f"  \\caption[{inline(short_caption(caption))}]{{{inline(caption)}}}")
    if label:
        out.append(f"  \\label{{{label}}}")
    first_rule = r"    \hline" if environment == "quadro" else r"    \toprule"
    out += [r"  \small", r"  \setlength{\tabcolsep}{2.5pt}", f"  \\begin{{tabularx}}{{\\textwidth}}{{{spec}}}", first_rule]
    for ri, row in enumerate(rows):
        cells = [inline(c, crossrefs=False) for c in row] + [""] * (cols - len(row))
        if label == "qua:resultado-A-1" and cells:
            cells[0] = cells[0].replace(r"\_", r"\_\allowbreak{}")
        if ri == 0:
            cells = [f"\\textbf{{{c}}}" for c in cells]
            cells = [c.replace("Depósito/evento", "Depósito/\\newline evento") for c in cells]
        out.append("    " + " & ".join(cells) + r" \\")
        if environment == "quadro":
            out.append(r"    \hline")
        elif ri == 0:
            out.append(r"    \midrule")
        elif is_appendix_table and ri < len(rows) - 1:
            out.append(r"    \midrule")
    if environment != "quadro":
        out.append(r"    \bottomrule")
    out.append(r"  \end{tabularx}")
    if note:
        out.append(f"  \\nota{{{inline(note)}}}")
    out.append(f"\\end{{{environment}}}")
    return out


def parse_markdown(
    section: str,
    image_map: dict[int, tuple[str, bytes]],
    appendix_letter: str | None = None,
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
        # Marcadores internos usados para registrar as seções do Google Docs
        # servem apenas à conversão e nunca devem aparecer no texto final.
        if s.startswith("<!--") and s.endswith("-->"):
            i += 1
            continue
        mimg = re.search(r"\[image(\d+)\]", s)
        if mimg:
            pending_image = int(mimg.group(1))
            remainder = re.sub(r"!?\[[^\]]*\]\[image\d+\]", "", s).strip()
            if not remainder:
                i += 1
                continue
            # Algumas legendas do Docs (notadamente a Figura 26) contêm o
            # objeto flutuante no fim do mesmo parágrafo. Conserva a imagem
            # pendente e continua interpretando o texto da legenda.
            s = remainder
        heading_prefix = appendix_letter or chapter_number
        mbold_heading = re.fullmatch(
            rf"\*\*({re.escape(heading_prefix)}(?:\.\d+)+)\s+(.+?)\*\*",
            s,
        )
        mplain_heading = re.fullmatch(
            rf"({re.escape(heading_prefix)}(?:\.\d+)+)\s+([^|]+)",
            s,
        )
        numbered_heading = mbold_heading or mplain_heading
        if numbered_heading:
            depth = numbered_heading.group(1).count(".")
            title = inline(numbered_heading.group(2), crossrefs=False)
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
            is_appendix_heading = bool(
                appendix_letter
                and re.match(rf"^APÊNDICE\s+{re.escape(appendix_letter)}\b", raw_title, re.IGNORECASE)
            )
            is_acknowledgements = raw_title == "8 AGRADECIMENTOS"
            if not (
                is_appendix_heading
                or is_acknowledgements
                or raw_title == chapter_heading
                or re.match(rf"^{re.escape(heading_prefix)}(?:\.\d+)*\s+\S", raw_title)
            ):
                out += [inline(title), ""]
                i += 1
                continue
            if level == 1:
                if appendix_letter:
                    title = re.sub(
                        rf"^APÊNDICE\s+{re.escape(appendix_letter)}\s*(?:\\?-|—|–)\s*",
                        "",
                        title,
                        flags=re.IGNORECASE,
                    )
                    appendix_labels = {
                        "A": "ap:codigos",
                        "B": "ap:localizacao-medidas",
                        "C": "ap:medicoes-cicatrizes",
                        "D": "ap:distancias-velocidades",
                        "E": "ap:dados-morfometria",
                        "F": "ap:hvsr",
                    }
                    out += [
                        f"\\chapter{{{inline(title, crossrefs=False)}}}"
                        f"\\label{{{appendix_labels[appendix_letter]}}}",
                        f"\\setcounter{{figure}}{{0}}\\renewcommand{{\\thefigure}}{{{appendix_letter}.\\arabic{{figure}}}}",
                        f"\\setcounter{{table}}{{0}}\\renewcommand{{\\thetable}}{{{appendix_letter}.\\arabic{{table}}}}",
                        f"\\setcounter{{quadro}}{{0}}\\renewcommand{{\\thequadro}}{{{appendix_letter}.\\arabic{{quadro}}}}",
                        "",
                    ]
                elif is_acknowledgements:
                    out += [r"\chapter{Agradecimentos}\label{cap:agradecimentos}", ""]
                else:
                    label = {
                        "1": "introducao",
                        "2": "fundamentacao",
                        "3": "area-estudo",
                        "4": "metodos",
                        "5": "resultados",
                        "6": "discussao",
                        "7": "conclusoes",
                    }.get(chapter_number, slug(chapter_title))
                    out += [f"\\chapter{{{chapter_title}}}\\label{{cap:{label}}}", ""]
            elif level == 2:
                out += [f"\\section{{{inline(title, crossrefs=False)}}}", ""]
            elif level == 3:
                out.append(f"\\subsection{{{inline(title, crossrefs=False)}}}")
                if title.casefold() == "velocidades utilizadas na análise hvsr".casefold():
                    out.append(r"\label{subsec:resultados-velocidades-hvsr}")
                out.append("")
            elif level == 4:
                out += [f"\\subsubsection{{{inline(title, crossrefs=False)}}}", ""]
            else:
                out += [f"\\paragraph{{{inline(title, crossrefs=False)}}}", ""]
            i += 1
            continue
        if s.startswith("* "):
            items: list[str] = []
            while i < len(lines):
                current = lines[i].strip()
                if current.startswith("* "):
                    items.append(current[2:].strip())
                    i += 1
                    continue
                # A exportação do Google Docs separa visualmente os itens com
                # uma linha vazia. Ela não deve encerrar a lista no LaTeX.
                if not current and i + 1 < len(lines) and lines[i + 1].strip().startswith("* "):
                    i += 1
                    continue
                break
            out += [r"\begin{itemize}"]
            out += [f"  \\item {inline(item)}" for item in items]
            out += [r"\end{itemize}", ""]
            continue
        mequation = re.fullmatch(r"(.*?),?\s*\\?\[(\d+)\\?\]", s)
        if mequation:
            number = int(mequation.group(2))
            if number in EQUATIONS:
                expression, label, description = EQUATIONS[number]
                out += [
                    r"\begin{equation}",
                    f"  {expression}",
                    f"  \\label{{{label}}}",
                    r"\end{equation}",
                    f"\\equationdescription{{{description}}}",
                    "",
                ]
                i += 1
                continue
        mfig = re.match(r"Figura\s+((?:[A-F]\.)?\d+)\s*[—–-]\s*(.*)", s)
        if mfig:
            number, caption = mfig.group(1), mfig.group(2).replace("\\.", ".")
            source_line: str | None = None
            note_line: str | None = None
            source_index = i + 1
            consumed_index = i + 1
            while source_index < len(lines):
                candidate = lines[source_index].strip().strip("*")
                if not candidate:
                    source_index += 1
                    continue
                if candidate.startswith("Fonte:"):
                    source_line = re.sub(r"^Fonte:\s*", "", candidate).rstrip(".")
                    consumed_index = source_index + 1
                    source_index += 1
                    continue
                if candidate.startswith("Nota:"):
                    note_line = re.sub(r"^Nota:\s*", "", candidate).rstrip(".")
                    consumed_index = source_index + 1
                    source_index += 1
                    continue
                break
            numeric_figure = number.isdigit()
            num = int(number) if numeric_figure else number
            label_number = str(number).replace(".", "-")
            figure_data = image_map.get(pending_image) if pending_image is not None else None
            # A Figura 26 é o único objeto flutuante do Google Docs. A exportação
            # Markdown/DOCX usada pelo sincronizador não a associa ao parágrafo
            # da legenda, embora o objeto continue presente no documento. Usa-se
            # a cópia original já versionada para impedir que a figura volte a
            # ser substituída por um marcador em sincronizações futuras.
            if figure_data is None and num == 26 and RECOVERED_FIGURE_26.exists():
                figure_data = (".png", RECOVERED_FIGURE_26.read_bytes())
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
                    f"  \\caption[{inline(short_caption(caption))}]{{{inline(caption)}}}",
                    f"  \\label{{fig:resultado-{label_number}}}",
                    *([f"  \\FonteTCC{{{inline(source_line, crossrefs=False)}}}"] if source_line else []),
                    *([f"  \\nota{{{inline(note_line)}}}"] if note_line else []),
                    r"\end{figure}",
                    "",
                ]
            else:
                ext, data = figure_data
                if ext not in {".png", ".jpg", ".jpeg", ".pdf"}:
                    ext = ".png"
                filename = (
                    f"figura-{num:02d}-{slug(caption)}{ext}"
                    if numeric_figure
                    else f"figura-{label_number}-{slug(caption)}{ext}"
                )
                FIGDIR.mkdir(parents=True, exist_ok=True)
                target = FIGDIR / filename
                if not target.exists() or target.read_bytes() != data:
                    target.write_bytes(data)
                landscape_figure = numeric_figure and num in LANDSCAPE_FIGURES
                if landscape_figure:
                    out.append(r"\begin{landscape}")
                max_height = "0.62\\textheight" if num == 48 else "0.80\\textheight"
                # As Figuras 47 e 48 precisam permanecer na ordem textual;
                # deixar apenas a 48 fixa permitiria que ela ultrapassasse a
                # 47. As figuras de domínios também ficam juntas ao texto e
                # com espaçamento previsível.
                fixed_placement = num in {47, 48} or (numeric_figure and 77 <= num <= 82)
                out += [
                    r"\begin{figure}[p]" if landscape_figure else (r"\begin{figure}[H]" if num in {5, 27, 52} or fixed_placement else r"\begin{figure}[htbp]"),
                    r"  \centering",
                    (f"  \\includegraphics[width=1.03\\linewidth,height={max_height},keepaspectratio]{{figuras/google-docs/{filename}}}"
                     if landscape_figure else
                     f"  \\includegraphics[width=1.03\\textwidth,height={max_height},keepaspectratio]{{figuras/google-docs/{filename}}}"),
                    f"  \\caption[{inline(short_caption(caption))}]{{{inline(caption)}}}",
                    f"  \\label{{fig:resultado-{label_number}}}",
                    *([r"  \label{fig:velocidade-profundidade-poco}"] if num == 52 else []),
                    *([f"  \\FonteTCC{{{inline(source_line, crossrefs=False)}}}"] if source_line else []),
                    *([f"  \\nota{{{inline(note_line)}}}"] if note_line else []),
                    r"\end{figure}",
                ]
                if landscape_figure:
                    out.append(r"\end{landscape}")
                out.append("")
            pending_image = None
            i = consumed_index
            continue
        caption_line = s.strip("*").strip()
        mtab = re.match(r"(Tabela|Quadro)\s+((?:[A-F]\.)?\d+)\s+[—–-]\s*(.*)", caption_line)
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
                cells = [
                    c.replace(r"\|", "|").strip()
                    for c in re.split(r"(?<!\\)\|", lines[i].strip().strip("|"))
                ]
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
                            target = FIGDIR / filename
                            if not target.exists() or target.read_bytes() != data:
                                target.write_bytes(data)
                            return rf"\includegraphics[width=\linewidth]{{figuras/google-docs/{filename}}}"
                        converted.append(re.sub(r"!\[[^\]]*\]\[image(\d+)\]", table_image, cell))
                    cells = converted
                    rows.append(cells)
                i += 1
            cap, lab, environment = pending_caption if pending_caption else (None, None, "table")
            note: str | None = None
            note_index = i
            while note_index < len(lines) and not lines[note_index].strip():
                note_index += 1
            if note_index < len(lines):
                candidate = lines[note_index].strip()
                if re.match(r"^\*{0,2}Nota(?: específica)?:\*{0,2}\s*", candidate):
                    note = re.sub(
                        r"^\*{0,2}Nota(?: específica)?:\*{0,2}\s*",
                        "",
                        candidate,
                    )
                    i = note_index + 1
            out += table_tex(rows, cap, lab, environment, note=note) + [""]
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
    text = MARKDOWN.read_text(encoding="utf-8")
    if re.search(r"(?m)^\[image\d+\]:\s*<data:image/", text):
        images = high_resolution_image_targets(DOCX, MARKDOWN)
    else:
        images = image_targets(DOCX)
    # A orientação vem das quebras de seção da própria exportação do Docs.
    # Assim, apenas as figuras que a autora colocou em páginas paisagem são
    # giradas no LaTeX, sem depender de uma lista manual que fica obsoleta.
    LANDSCAPE_FIGURES = landscape_figure_numbers(DOCX)
    LANDSCAPE_FIGURES.update(landscape_figure_numbers_api(text))

    chapter_specs = {
        "1": ("Introdução", r"^#\s+(?:\*\*)?1\s+INTRODUÇÃO(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?2\s+FUNDAMENTAÇÃO TEÓRICA(?:\*\*)?\s*$"),
        "2": ("Fundamentação teórica", r"^#\s+(?:\*\*)?2\s+FUNDAMENTAÇÃO TEÓRICA(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?3\s+ÁREA DE ESTUDO(?:\*\*)?\s*$"),
        "3": ("Área de estudo", r"^#\s+(?:\*\*)?3\s+ÁREA DE ESTUDO(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?4\s+MATERIAIS E MÉTODOS(?:\*\*)?\s*$"),
        "4": ("Materiais e métodos", r"^#\s+(?:\*\*)?4\s+MATERIAIS E MÉTODOS(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?5\s+RESULTADOS(?:\*\*)?\s*$"),
        "5": ("Resultados", r"^#\s+(?:\*\*)?5\s+RESULTADOS(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?6\s+DISCUSS(?:ÃO|ÕES)(?:\*\*)?\s*$"),
        "6": ("Discussões", r"^#\s+(?:\*\*)?6\s+DISCUSS(?:ÃO|ÕES)(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?7\s+CONCLUSÕES(?:\*\*)?\s*$"),
        "7": ("Conclusões", r"^#\s+(?:\*\*)?7\s+CONCLUSÕES(?:\*\*)?\s*$", r"^#\s+(?:\*\*)?REFERÊNCIAS(?:\*\*)?\s*$"),
    }
    for chapter_number, (chapter_title, start, end) in chapter_specs.items():
        section = between_headings(text, start, end)
        CHAPTERS[chapter_number].write_text(
            parse_markdown(
                section,
                images,
                chapter_number=chapter_number,
                chapter_title=chapter_title,
            ),
            encoding="utf-8",
        )
    appendix_sections: dict[str, str] = {}
    for index, letter in enumerate("ABCDEF"):
        start_pattern = rf"^#\s+(?:\*\*)?APÊNDICE\s+{letter}\b.*?(?:\*\*)?\s*$"
        if letter != "F":
            next_letter = chr(ord(letter) + 1)
            appendix_sections[letter] = between_headings(
                text,
                start_pattern,
                rf"^#\s+(?:\*\*)?APÊNDICE\s+{next_letter}\b.*?(?:\*\*)?\s*$",
            )
        else:
            start = re.search(start_pattern, text, flags=re.MULTILINE)
            if not start:
                raise ValueError(f"Cabeçalho inicial não encontrado: {start_pattern}")
            appendix_sections[letter] = text[start.start():]
    for letter, section in appendix_sections.items():
        appendix_tex = parse_markdown(section, images, appendix_letter=letter)
        if letter == "A":
            appendix_tex = re.sub(
                r"Link:\s+https:?\\?//github\.com/mariaeulalia2002/codigos\\?_tcc",
                r"Link: \\url{https://github.com/mariaeulalia2002/codigos_tcc}",
                appendix_tex,
            )
        APENDICES[letter].write_text(appendix_tex, encoding="utf-8")
    for number, path in CHAPTERS.items():
        print(f"Capítulo {number}: {path}")
    for letter, path in APENDICES.items():
        print(f"Apêndice {letter}: {path}")


if __name__ == "__main__":
    main()
