' Testing to make the smart button to recall 


'=============================================================================================
'Set up variables
Dim Arg, var2
Dim wshShell
Set Arg = WScript.Arguments

'=============================================================================================
'initialize variables
'Parameter1, begin with index
var2 = Arg(0)
'the play button pass the argrument "play" to this code
if (var2 = "play") then
	Dim oShell 
	Dim var1 
	set OShell = CreateObject("Wscript.Shell")
	'Read the variable stored in the VariableofVmixInput then be able to pass that variable to api link
	var1 = oShell.RegRead("HKCU\StoreAPIVariable\VariableofVmixInput")

	dim o 
	set o = CreateObject("MSXML2.XMLHTTP")
	'prep mesage and send
	'msgbox "First parameter passed was " _
	 '      & var1 
	 'sent API to VMix by telling it to play the video
	o.open "GET", "http://localhost:8088/api/?Function=Play&Input="& var1, false
	o.send 
end if 

'=============================================================================================
'the pause button pass the argrument "pause" to this code
if (var2 = "pause") then
	set OShell = CreateObject("Wscript.Shell")
	'Read the variable stored in the VariableofVmixInput then be able to pass that variable to api link
	var1 = oShell.RegRead("HKCU\StoreAPIVariable\VariableofVmixInput")
	dim i 
	set i = CreateObject("MSXML2.XMLHTTP")
	'msgbox "First parameter passed was " _
	 '      & var1 
	'sent API to VMix by telling it to pause the video
	i.open "GET", "http://localhost:8088/api/?Function=Pause&Input="& var1, false
	i.send 
end if 

'=============================================================================================
'the restart button pass the argrument "restart" to this code
if (var2 = "restart") then
	
	set OShell = CreateObject("Wscript.Shell")
	'Read the variable stored in the VariableofVmixInput then be able to pass that variable to api link
	var1 = oShell.RegRead("HKCU\StoreAPIVariable\VariableofVmixInput")
	dim r 
	set r = CreateObject("MSXML2.XMLHTTP")
	'msgbox "First parameter passed was " _
	 '      & var1 
	'sent API to VMix by telling it to restart the video
	r.open "GET", "http://localhost:8088/api/?Function=Restart&Input="& var1, false
	r.send 
end if 
'=============================================================================================
'Clear the objects at the end of your script.
set Arg = Nothing