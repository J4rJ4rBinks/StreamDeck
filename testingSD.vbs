' Testing to make the smart button to recall 

Dim Arg, var1, var2
Dim wshShell
Set Arg = WScript.Arguments
'Parameter1, begin with index0
var1 = Arg(0)
dim o 
set o = CreateObject("MSXML2.XMLHTTP")

o.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input="& var1, false
o.send 
'Pass the variable to the registry and allow to write the var1 to registry
msgbox "First parameter passed was " _
       & var1 
	  ' Create a WSH Shell object:
	Set wshShell = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key:
	wshShell.RegWrite "HKCU\TestKey\", ""

	' Create a new DWORD value:
	wshShell.RegWrite "HKCU\TestKey\DWordTestValue", var1, "REG_SZ"

'Clear the objects at the end of your script.
set Arg = Nothing