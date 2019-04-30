' **************************************************************************
' * Program:		myrecordings.vbs
' * Purpose:		Open the folder depends on the pathings
' *
' * Date:			8/22/2018
' **************************************************************************
On Error Resume Next
dim fso, username
Set fso = CreateObject("Scripting.FileSystemObject")

'--------------------------------------------------
' Get username
Set net = createObject("wscript.network")
username = net.username
'wscript.echo username
'--------------------------------------------------
Set shell = wscript.CreateObject("Shell.Application")
'wscript.echo "1"
shell.Open "C:\Users\" + username + "\Desktop\My Recordings" 