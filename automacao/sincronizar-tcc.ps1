[CmdletBinding()]
param(
    [string]$Repositorio
)

$ErrorActionPreference = 'Stop'
if (-not $Repositorio) {
    $pastaScript = Split-Path -Parent $MyInvocation.MyCommand.Path
    $Repositorio = Split-Path -Parent $pastaScript
}
$git = 'C:\Program Files\Git\cmd\git.exe'
$pastaLog = Join-Path $env:LOCALAPPDATA 'TCC-REAL'
$arquivoLog = Join-Path $pastaLog 'sync.log'
$mutex = [Threading.Mutex]::new($false, 'Local\TCC-REAL-Git-Sync')

New-Item -ItemType Directory -Path $pastaLog -Force | Out-Null

function Registrar([string]$Mensagem) {
    Add-Content -LiteralPath $arquivoLog -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Mensagem" -Encoding UTF8
}

function Git {
    # O Windows PowerShell 5 transforma o stderr de programas nativos em
    # ErrorRecord. Capture-o sem deixar ErrorActionPreference=Stop interromper
    # comandos bem-sucedidos como `git fetch`.
    $preferenciaAnterior = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $saida = & $git -C $Repositorio @args 2>&1
    $codigo = $LASTEXITCODE
    $ErrorActionPreference = $preferenciaAnterior
    $saida | ForEach-Object { Registrar "git: $_" }
    if ($codigo -ne 0) {
        throw "O comando git falhou (código $codigo): git $($args -join ' ')"
    }
}

try {
    if (-not $mutex.WaitOne(0)) { exit 0 }
    if (-not (Test-Path -LiteralPath $git)) { throw "Git não encontrado em $git" }

    Registrar 'Início da sincronização.'

    # Um repositório recém-criado pode ainda não ter a branch main remota.
    & $git -C $Repositorio ls-remote --exit-code --heads origin main *> $null
    $mainRemotaExiste = ($LASTEXITCODE -eq 0)
    if ($LASTEXITCODE -notin 0, 2) {
        throw 'Não foi possível consultar a branch main no GitHub.'
    }

    if ($mainRemotaExiste) {
        Git fetch origin main --prune
    }

    $alteracoes = & $git -C $Repositorio status --porcelain=v1
    if ($LASTEXITCODE -ne 0) { throw 'Não foi possível consultar o estado do repositório.' }

    if ($alteracoes) {
        Git add --all
        $mensagem = "Atualização automática $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Git commit -m $mensagem
    }

    if ($mainRemotaExiste) {
        Git rebase origin/main
    }
    Git push origin main
    Registrar 'Sincronização concluída.'
}
catch {
    Registrar "ERRO: $($_.Exception.Message)"
    exit 1
}
finally {
    try { $mutex.ReleaseMutex() } catch { }
    $mutex.Dispose()
}
