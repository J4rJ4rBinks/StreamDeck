' Testing to make the smart button to recall 
'declare varable for Arg and var1
	Dim Arg, var1
	Dim wshShell
	'set arg to pass argruments from StreamDeck Console to this code
	Set Arg = WScript.Arguments
	'Get whatever number from the StreamDeck Console to pass the number to this code
	var1 = Arg(0)
	
'================================================================================================================'
	'change only this block 
	dim vmixvariable
	
	Select Case var1 
		Case "howtousestreamdeck"
			vmixvariable = "10"
		Case "howtoeditanddelete"
			vmixvariable = "14"
		Case "Howtoaddannotation"
			vmixvariable = "11"
		Case "Howtoaddcaptions"
			vmixvariable = "12" 
		Case "Howtoaddtransition" 
			vmixvariable = "15"
		Case "Howtoburnincaptions"
			vmixvariable = "13"
		Case "Howtouploadyoutube"
			vmixvariable = "16"
	End Select
' do not change anything below 
'================================================================================================================'
	dim o 
	set o = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api 
	o.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input="& vmixvariable, 0
	o.send 
'testing the message to confirm what number it was sending
'msgbox "First parameter passed was " _
 '      & var1 
	  ' Create a WSH Shell object:
	Set wshShell = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	wshShell.RegWrite "HKCU\StoreAPIVariable\", ""

	' Create a new DWORD value to store the string of variable:
	wshShell.RegWrite "HKCU\StoreAPIVariable\VariableofVmixInput", vmixvariable, "REG_SZ"

'Clear the objects at the end of your script.
set Arg = Nothing