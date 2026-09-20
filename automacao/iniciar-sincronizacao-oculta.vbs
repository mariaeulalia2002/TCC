Option Explicit

Dim shell, comando
Set shell = CreateObject("WScript.Shell")

comando = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""C:\Users\maria_ym81dto\Documents\TCC REAL\automacao\sincronizar-tcc.ps1"" -Repositorio ""C:\Users\maria_ym81dto\Documents\TCC REAL"""

' O segundo argumento igual a 0 executa o PowerShell sem abrir uma janela.
shell.Run comando, 0, False
