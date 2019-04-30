'========================================================================================================================================================'
	'only change the variable here 
	dim variable
	variable = "14"
'do not change the code below only if you need to add function or something go ahead. But be careful.
'======================================================================================================================================================='
	dim stoprec
	set stoprec = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api 
	stoprec.open "GET", "http://localhost:8088/api/?Function=StopRecording", 0
	stoprec.send 
' change the after stop recording change the input to how to edit with camtasia input 

	dim changeinput
	set changeinput = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api 
	changeinput.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input=" & variable, 0
	changeinput.send 
	
'and start playing the video. 

	dim playvid
	set playvid = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api 
	playvid.open "GET", "http://localhost:8088/api/?Function=Play&Input=" & variable, 0
	playvid.send 

	
'then change the registry because the registry is 10 so we dont want that. so we rewrite 10 to 14 in the registry for smart button 
	Set wshShell = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	wshShell.RegWrite "HKCU\StoreAPIVariable\", ""

	' Create a new DWORD value to store the string of variable:
	wshShell.RegWrite "HKCU\StoreAPIVariable\VariableofVmixInput", variable, "REG_SZ"
	
	