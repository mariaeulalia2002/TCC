# Templates de TCC em Geologia da UFSC

> Repositório de trabalho de Maria Eulália. O histórico principal é mantido em
> `mariaeulalia2002/TCC`; o template original permanece configurado como remoto
> `upstream`.

Este repositório reúne duas versões do mesmo modelo de Trabalho de Conclusão de Curso para o Curso de Graduação em Geologia da Universidade Federal de Santa Catarina e um guia para manter as versões sincronizadas.

## Estrutura do repositório

```text
.
├── latex/      projeto modular para Overleaf, XeLaTeX ou Tectonic
├── docs/       modelo DOCX editável no Word e importável no Google Docs
└── guia-ia/    instruções para converter o texto do Docs para LaTeX
```

## Qual versão usar

- Use [`latex/`](latex/) para escrever diretamente no Overleaf ou manter a versão final tipograficamente controlada.
- Use [`docs/`](docs/) para escrever de forma colaborativa no Google Docs ou no Microsoft Word.
- Use [`guia-ia/`](guia-ia/) quando um agente de IA for responsável por levar o conteúdo escrito no Docs para o projeto LaTeX.

As duas versões seguem a mesma organização geral: elementos pré-textuais, introdução, área de estudo e contexto geológico, materiais e métodos, resultados, discussão, conclusões, referências, apêndices e anexos.

## Fluxo recomendado

1. Escreva e revise o conteúdo no Google Docs usando o arquivo de `docs/` como ponto de partida.
2. Exporte uma cópia em `.docx` quando houver uma versão pronta para conversão.
3. Entregue o `.docx`, a pasta `latex/` e o guia de `guia-ia/` ao agente de IA.
4. Compile o LaTeX e confira visualmente texto, citações, figuras, tabelas, sumário e referências.
5. Mantenha o Docs como fonte editorial até definir que o LaTeX passou a ser a versão principal.

> Este é um modelo de apoio, não um documento oficial da UFSC. Antes da entrega, confirme as normas vigentes com a Coordenação do Curso de Geologia e com a Biblioteca Universitária.

## Sincronização neste computador

A pasta é sincronizada automaticamente com o GitHub a cada dois minutos pela
tarefa do Windows **TCC - sincronização GitHub**. A automação:

1. cria um commit quando encontra arquivos alterados;
2. incorpora os commits existentes no GitHub com `rebase`;
3. envia a versão local para a branch `main`;
4. registra ocorrências em `%LOCALAPPDATA%\TCC-REAL\sync.log`.

Para sincronizar imediatamente, execute no PowerShell:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\automacao\sincronizar-tcc.ps1"
```

O MiKTeX está instalado neste computador. Para gerar `latex/main.pdf`, execute:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\automacao\compilar-latex.ps1"
```

Arquivos nativos do Google Docs não ficam dentro de uma pasta Git. Evite editar
simultaneamente o mesmo trecho em dois computadores; conflitos ficam
preservados para revisão manual e nunca são descartados pela automação.

## Documento editorial no Google Docs

O documento editorial oficial deste TCC é
[`template-tcc-geologia-ufsc`](https://docs.google.com/document/d/1ZoLTqc8Ihwv2FqDgTXCC3GTA_VLcLR78uw_x9cwu1ME/edit).

O fluxo de trabalho adotado é:

1. escrever e revisar o conteúdo no Google Docs;
2. solicitar a conversão quando uma etapa estiver pronta;
3. usar o Google Docs como autoridade para o conteúdo e `latex/` como autoridade
   para estrutura, comandos e formatação;
4. converter títulos, texto, citações, tabelas, figuras e legendas para os
   arquivos modulares do projeto;
5. compilar e revisar `latex/main.pdf`;
6. registrar e enviar a conversão ao GitHub.

A conversão do Google Docs para LaTeX é feita sob demanda. Isso evita que
rascunhos intermediários ou edições ainda incompletas sejam propagados
automaticamente para a versão destinada à banca.

