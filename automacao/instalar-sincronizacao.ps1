[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$nomeTarefa = 'TCC - sincronização GitHub'
$scriptSync = Join-Path $PSScriptRoot 'sincronizar-tcc.ps1'
$powershell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

if (-not (Test-Path -LiteralPath $scriptSync)) {
    throw "Script de sincronização não encontrado: $scriptSync"
}

$acao = New-ScheduledTaskAction `
    -Execute $powershell `
    -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptSync`""

$gatilho = New-ScheduledTaskTrigger `
    -Once `
    -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes 2) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$configuracao = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName $nomeTarefa `
    -Action $acao `
    -Trigger $gatilho `
    -Settings $configuracao `
    -Description 'Versiona e sincroniza automaticamente o TCC com o GitHub.' `
    -Force | Out-Null

Write-Output "Tarefa instalada: $nomeTarefa"
