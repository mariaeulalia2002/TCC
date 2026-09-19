[CmdletBinding()]
param(
    [string]$Repositorio
)

$ErrorActionPreference = 'Stop'
if (-not $Repositorio) {
    $pastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path
    $Repositorio = Split-Path -Parent $pastaScript
}
$miktex = Join-Path $env:LOCALAPPDATA 'Programs\MiKTeX\miktex\bin\x64'
$xelatex = Join-Path $miktex 'xelatex.exe'
$bibtex = Join-Path $miktex 'bibtex.exe'
$pastaLatex = Join-Path $Repositorio 'latex'

if (-not (Test-Path -LiteralPath $xelatex)) {
    throw "XeLaTeX não encontrado. Instale o MiKTeX antes de compilar."
}

Push-Location $pastaLatex
try {
    & $xelatex --enable-installer -interaction=nonstopmode -halt-on-error main.tex
    if ($LASTEXITCODE -ne 0) { throw 'A primeira execução do XeLaTeX falhou.' }

    & $bibtex main
    if ($LASTEXITCODE -ne 0) { throw 'A geração das referências com BibTeX falhou.' }

    1..2 | ForEach-Object {
        & $xelatex --enable-installer -interaction=nonstopmode -halt-on-error main.tex
        if ($LASTEXITCODE -ne 0) { throw "A execução final $_ do XeLaTeX falhou." }
    }

    Write-Output "PDF gerado: $(Join-Path $pastaLatex 'main.pdf')"
}
finally {
    Pop-Location
}
