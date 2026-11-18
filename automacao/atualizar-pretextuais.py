"""Sincroniza os resumos do Google Docs com os arquivos pre-textuais."""

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def replace(path: Path, old: str, new: str) -> None:
    text = path.read_text(encoding="utf-8")
    if old not in text:
        if new in text:
            return
        raise SystemExit(f"Trecho esperado nao encontrado em {path}")
    path.write_text(text.replace(old, new, 1), encoding="utf-8")


replace(
    ROOT / "latex" / "pre-textuais" / "resumo.tex",
    "Este trabalho teve como objetivo identificar, mapear e caracterizar depósitos e complexos de transporte de massa no talude norte da Bacia de Santos, com ênfase na análise morfométrica dos eventos. A análise também abordou a distribuição espacial, a organização estratigráfica, a configuração interna, o confinamento e as relações com o relevo do Canal de Santos e a tectônica salífera.",
    "Este trabalho teve como objetivo identificar, mapear e caracterizar depósitos e complexos de transporte de massa no talude norte da Bacia de Santos. As análises abordaram a distribuição espacial, a organização estratigráfica, a morfometria, a configuração interna, o confinamento e as relações com o relevo do Canal de Santos e a tectônica salífera.",
)

replace(
    ROOT / "latex" / "pre-textuais" / "abstract.tex",
    "This study aimed to identify, map, and characterize mass-transport deposits and complexes on the northern slope of the Santos Basin, with emphasis on the morphometric analysis of the events. The analysis also addressed their spatial distribution, stratigraphic organization, internal configuration, confinement conditions, and relationships with the relief associated with the Santos Channel and salt tectonics.",
    "This study aimed to identify, map, and characterize mass-transport deposits and complexes on the northern slope of the Santos Basin. The analysis addressed their spatial distribution, stratigraphic organization, morphometric characterization, internal configuration, confinement conditions, and relationships with the relief associated with the Santos Channel and salt tectonics.",
)
