	' do not change anything below this is only write to vmix input of view to the registry. 
	
	'write to registry for all view inputs 
	Dim Arg, var1
	Dim wshShell
	'set arg to pass argruments from StreamDeck Console to this code
	Set Arg = WScript.Arguments
	'Get whatever number from the StreamDeck Console to pass the number to this code
	var1 = Arg(0)
	'====================================================================================================================================================='
	dim vmixvariable
	
	Select Case var1 
		Case "PicinPic"
			vmixvariable = "6"
		Case "Singleview"
			vmixvariable = "1"
		Case "DesktopView"
			vmixvariable = "2"
		Case "Weathermanview"
			vmixvariable = "4" 
		Case "SidebySide" 
			vmixvariable = "5"
		Case "CircularStudio" 
			vmixvariable ="7"
		Case "Latenightnews"
			vmixvariable= "8" 
		Case "NewsStudio" 
			vmixvariable ="9"
	End Select
	
	'=========================================================================================================================================================='
	dim moveinputs
	set moveinputs = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api and show this input
	moveinputs.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input=" & vmixvariable, 0
	moveinputs.send 
	
	' Create a WSH Shell object:
	Set wshShell = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	wshShell.RegWrite "HKCU\Storeviewinputs\", ""

	' Create a new DWORD value to store the string of variable:
	wshShell.RegWrite "HKCU\Storeviewinputs\inputsofviews", vmixvariable, "REG_SZ"

'Clear the objects at the end of your script.
set Arg = Nothing