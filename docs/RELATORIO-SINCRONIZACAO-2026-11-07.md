# Relatório de sincronização e revisão — 7 de novembro de 2026

## Fontes e preservação

- Documento principal: Google Docs `1ZoLTqc8Ihwv2FqDgTXCC3GTA_VLcLR78uw_x9cwu1ME`.
- Revisão exportada: `AHj4eMRV26O5k9q2Wsl1tSj2RCKmnlUlkVj8JGV9SSYlNSxVjhIphRwr2YrMJozoDEHLWPP5QGWQGcc00fLIqfaT7SmHQm8gT8pQlo-Nh5E`.
- Cópia de segurança anterior à revisão: `https://docs.google.com/document/d/1MkHlDKWlwPUPcmyCcGketRuKhHgyPDYiyULQI80I21s/edit`.
- Os destaques, observações e anotações da autora foram preservados.

## Ajustes no Google Docs

- Aplicação de Times New Roman 12 no corpo, cor preta, alinhamento justificado, recuo de primeira linha de 1,5 cm e espaçamento 1,5, sem espaço extra entre parágrafos.
- Hierarquia de títulos normalizada; quatro títulos numerados que estavam como corpo foram promovidos ao nível correto. Seções primárias permanecem com quebra de página.
- Legendas, fontes, notas, tabelas e equações foram mantidas em espaçamento simples e tamanho reduzido conforme a função editorial.
- Equações centralizadas.
- Padronização de 370 ocorrências de sobrenomes em citações para a grafia da NBR 10520:2023; manutenção da convenção acordada de `&` para dois autores e de *et al.* em itálico para mais de dois autores.
- Correção de 10 ocorrências de `Jr..`, de `Figura62`, de `Figura Iv.6` e de uma citação narrativa com dois autores.
- 145 ocorrências de *et al.* receberam itálico.
- 56 tabelas tiveram a primeira linha fixada para repetir os cabeçalhos de coluna nas páginas seguintes. O Google Docs não repete automaticamente uma legenda externa à tabela; no LaTeX, título e cabeçalho são repetidos.
- Referências mantidas em sequência, com espaçamento simples interno e 12 pt entre entradas, sem uma referência por página.

## Sincronização com o LaTeX

- Resultados, Discussão, Conclusões e Apêndice D foram regenerados a partir da exportação final do Google Docs.
- Alterações recentes de Área de estudo e Metodologia foram incorporadas ao repositório.
- Referências de Kayen e Lee (1991), Li et al. (2025) e Petrobras (2022) foram incluídas no banco BibTeX com autoria completa.
- A auditoria final não encontrou grupos autor-data crus nos capítulos gerados nem citações sem chave bibliográfica.
- Tabelas extensas usam `longtable`, com título de continuação e cabeçalho repetido. Os quadros extensos foram divididos em páginas em paisagem com título e cabeçalho repetidos.
- A orientação passou a ser obtida das seções do DOCX. As Figuras 29–42, 44 e 45 estão em paisagem; a Figura 43 e a figura de *pockmarks* permanecem em retrato.
- O PDF foi compilado com XeLaTeX + BibTeX e estabilizou em 232 páginas, sem referências ou citações indefinidas.

## Imagens em alta resolução

- A causa da baixa qualidade foi identificada: o conversor priorizava imagens reduzidas do Markdown, muitas com cerca de 605 px.
- O conversor agora extrai as mídias originais do DOCX, aplica os recortes visuais registrados no documento e associa cada imagem à sua posição no Markdown.
- O conjunto usado pelo LaTeX passou de aproximadamente 22 MB para cerca de 128 MB; diversas imagens chegam a 2.048 px.
- As figuras usam a maior largura disponível e limite de altura seguro de 0,80 da altura do texto em retrato e 0,78 em paisagem, preservando proporção e espaço para legendas.
- Foi feita inspeção visual das 232 páginas em miniaturas e verificação ampliada de figuras em paisagem, quadros continuados e apêndices.

## Estado final e ressalvas

- A automação de versionamento estava ativa e criou commits intermediários no `main` remoto durante o trabalho.
- Permanecem apenas avisos tipográficos não fatais: algumas linhas longas na lista de figuras/sumário e no texto, além de avisos internos de divisão de caixas em tabelas. Não há conteúdo cortado nas páginas inspecionadas.
- As figuras foram preservadas como estão no Docs; futuras substituições por arquivos-fonte de maior resolução continuarão recomendáveis para as que ainda tenham baixa resolução de origem.
