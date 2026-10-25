# Histórico de sincronização

## 18 de outubro de 2026 — equações, símbolos, tabelas e hierarquia

- Versão Git anterior ao ciclo de revisão: `010079e` (`Padroniza tabelas e atualiza lista de tabelas`).
- Revisão inicial do Google Docs: `AHj4eMSE_cH1ebJBFfQnlA0j1Yg9rMMmnju33jWAZda51wSGSTtEpHPF5KypXsaAUngePAeBVo26mWm3pyKBlCzxr4qZxhuL3RbfP2gM8y0`.
- Revisão do Google Docs apó a padronização: `AHj4eMRXCQuV_UxeWHqMzdpv93u7fJt78WDBE74PDqI9QZsRRwCIYHsMVQ0UeqYMhrNWPOUYWhNH7vUOTOl21yUURQ1bLrIR5O5ChhdDL_E`.

Alterações realizadas:

- conferência das 46 equações, com centralização, vírgula antes da numeração e sequência equivalente no Docs e no LaTeX;
- correção de índices e símbolos matemáticos nas equações e no texto corrido;
- reorganização dos resultados de velocidades e HVSR, mantendo os procedimentos na metodologia;
- padronização das 18 tabelas e reconstrução da lista de tabelas;
- correção da hierarquia dos títulos, com fonte preta;
- correção da ordem das Figuras 29 e 30;
- preservação das observações e dos destaques amarelos do Google Docs.

Valores que permanecem pendentes de confirmação científica pela autora, sem alteração automática: `1.528,37`, `1.6556` e `214,–301,1`.

## 25 de outubro de 2026 — resultados, figuras, tabelas e Apêndice D

- O Google Docs foi mantido como documento principal e exportado integralmente para `docs/template-tcc-geologia-ufsc.docx` e `docs/google-docs-latest.md`.
- O capítulo de resultados foi sincronizado com o LaTeX, incluindo as Figuras 29 a 49 e as Tabelas 5 a 43.
- As tabelas foram convertidas para o ambiente de tabela, com largura limitada à área útil da página e espaçamento interno compacto.
- O Apêndice D foi incluído no documento LaTeX com os dados de entrada e as memórias de cálculo das alturas e inclinações.
- Para o evento D do MTC 2, o ponto A originalmente definido para o eixo 1 passou a ser a origem comum dos eixos 1 e 2; foram preservadas terminações distintas C1 e C2.
- Os valores do eixo 2 foram sincronizados como `594,82 ± 0,54 ms`, `429,71 ± 0,14 ms`, `451,2 ± 1,2 m`, `0,28%` e `23,48 km`.
- Foram preservadas as observações e os destaques amarelos ainda pendentes no documento principal.
- Foi adicionado um conversor reproduzível em `automacao/converter-resultados-docs.py` para atualizar os resultados, as imagens e o Apêndice D a partir das exportações do Google Docs.

Pendências mantidas para preenchimento posterior pela autora: resultados da análise sísmica/estratigráfica, compartimentação em domínios e observações destacadas sobre valores que ainda exigem confirmação científica.
