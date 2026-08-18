' Lanca o auto_push.ps1 sem NENHUMA janela aparecer na tela (nem por uma
' fracao de segundo). O "-WindowStyle Hidden" do PowerShell sozinho nao
' resolve isso 100% -- o Windows chega a desenhar a janela preta antes de
' escondê-la, e é esse "pisca" que aparece a cada 3 minutos. Rodar o
' PowerShell por dentro do WScript.Shell (com o parametro de janela = 0)
' evita esse flash porque a janela nunca chega a ser criada de verdade.
Set objShell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
comando = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptDir & "\auto_push.ps1"""
objShell.Run comando, 0, False
