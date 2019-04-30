' Testing to make the smart button to recall 

Dim Arg, var2
Dim wshShell
Set Arg = WScript.Arguments
'Parameter1, begin with index0
var2 = Arg(0)
if (var2 = "play") then

Dim oShell 
Dim var1 
set OShell = CreateObject("Wscript.Shell")

var1 = oShell.RegRead("HKCU\TestKey\DWordTestValue")

dim o 
set o = CreateObject("MSXML2.XMLHTTP")

msgbox "First parameter passed was " _
       & var1 
o.open "GET", "http://localhost:8088/api/?Function=Play&Input="& var1, false
o.send 
end if 

if (var2 = "pause") then

set OShell = CreateObject("Wscript.Shell")

var1 = oShell.RegRead("HKCU\TestKey\DWordTestValue")
dim i 
set i = CreateObject("MSXML2.XMLHTTP")
msgbox "First parameter passed was " _
       & var1 

i.open "GET", "http://localhost:8088/api/?Function=Pause&Input="& var1, false
i.send 
end if 

if (var2 = "restart") then

set OShell = CreateObject("Wscript.Shell")

var1 = oShell.RegRead("HKCU\TestKey\DWordTestValue")
dim r 
set r = CreateObject("MSXML2.XMLHTTP")
msgbox "First parameter passed was " _
       & var1 

r.open "GET", "http://localhost:8088/api/?Function=Restart&Input="& var1, false
r.send 
end if 



'Clear the objects at the end of your script.
set Arg = Nothing