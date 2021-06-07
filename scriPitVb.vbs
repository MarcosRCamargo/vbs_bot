Dim objFile, strLine, botwpp
Set Fso = CreateObject("Scripting.FileSystemObject")

set WshShell = WScript.CreateObject("WScript.Shell")


WshShell.run "https://web.whatsapp.com/"

WScript.sleep 5000
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Users\marco\Desktop\py\contatos.txt", 1)
Do While Not objFile.AtEndOfStream
	strContato = objFile.readline
    strLine = "Ola " & strContato  &  " sou um Bot de mensagens automaticas para whatsapp"
    WshShell.SendKeys "{TAB}"
    WScript.sleep 5000
    WshShell.SendKeys"" & strContato
    WScript.sleep 3000
    WshShell.SendKeys "{ENTER}"
    WScript.sleep 2000
    WshShell.SendKeys strLine
    WScript.sleep 3000
    WshShell.SendKeys "{ENTER}"
    WScript.sleep 2000
    WshShell.SendKeys "{TAB}"
    WScript.sleep 2000
    WshShell.SendKeys "{TAB}"
    WScript.sleep 2000
    WshShell.SendKeys "{TAB}"
    WScript.sleep 2000
    WshShell.SendKeys "{TAB}"
Loop
objFile.Close