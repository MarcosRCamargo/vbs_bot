Dim objFile, strLine, botwpp

Set Fso = CreateObject("Scripting.FileSystemObject")

set WshShell = WScript.CreateObject("WScript.Shell")


' WshShell.run "https://web.whatsapp.com/"

WScript.sleep 5000
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile_amigos = objFSO.OpenTextFile("C:\Users\marco\Desktop\vbs\vbs_bot\contatos\contatos_amigos.txt", 1)
Set objFile_especial = objFSO.OpenTextFile("C:\Users\marco\Desktop\vbs\vbs_bot\contatos\contatos_especial.txt", 1)
Set objFile_familia = objFSO.OpenTextFile("C:\Users\marco\Desktop\vbs\vbs_bot\contatos\contatos_familia.txt", 1)
Set objFile_colegas = objFSO.OpenTextFile("C:\Users\marco\Desktop\vbs\vbs_bot\contatos\contatos_colegas.txt", 1)
Set objFile_teste = objFSO.OpenTextFile("C:\Users\marco\Desktop\vbs\vbs_bot\contatos\contatos_teste.txt", 1)

WScript.Echo (objFile_amigos.Name())

' Function defineMessage(today, listType)
'     Select Case listType
'    Case listType = (objFile_amigos)
'       statement1
'       statement2
'       ....
'       ....
'       statement1n
'    Case expressionlist2
'       statement1
'       statement2
'       ....
'       ....
'    Case expressionlistn
'       statement1
'       statement2
'       ....
'       ....   
'   Case Else
'       elsestatement1
'       elsestatement2
'       ....
'       ....
' End Select
    
' End Function 

' Function sendMessage(listContact)
'   Do While Not listContact.AtEndOfStream
' 	strContato = listContact.readline
'     strLine = "Ola " & strContato  &  " sou um Bot de mensagens automaticas para whatsapp"
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 5000
'     WshShell.SendKeys"" & strContato
'     WScript.sleep 3000
'     WshShell.SendKeys "{ENTER}"
'     WScript.sleep 2000
'     WshShell.SendKeys strLine
'     WScript.sleep 3000
'     WshShell.SendKeys "{ENTER}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
' Loop
' listContact.Close
' End Function


' sendMessage(objFile_teste)
' Do While Not objFile.AtEndOfStream
' 	strContato = objFile.readline
'     strLine = "Ola " & strContato  &  " sou um Bot de mensagens automaticas para whatsapp"
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 5000
'     WshShell.SendKeys"" & strContato
'     WScript.sleep 3000
'     WshShell.SendKeys "{ENTER}"
'     WScript.sleep 2000
'     WshShell.SendKeys strLine
'     WScript.sleep 3000
'     WshShell.SendKeys "{ENTER}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
'     WScript.sleep 2000
'     WshShell.SendKeys "{TAB}"
' Loop
' objFile.Close