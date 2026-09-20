Option Explicit

Dim shell, comando, codigoSaida
Set shell = CreateObject("WScript.Shell")

comando = """C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""C:\Users\maria_ym81dto\Documents\TCC REAL\automacao\sincronizar-tcc.ps1"" -Repositorio ""C:\Users\maria_ym81dto\Documents\TCC REAL"""

' Janela 0 mantém a execução invisível; True aguarda e propaga o resultado.
codigoSaida = shell.Run(comando, 0, True)
WScript.Quit codigoSaida
