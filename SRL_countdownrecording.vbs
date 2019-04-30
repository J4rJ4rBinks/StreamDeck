'==============================================================================================================================='
' change the variable here for the input of vmix that has countdown number on the screen. 
	dim vmixinput 
	vmixinput = "17"
	'do not change the code below only if you need to add function or something go ahead. But be careful.
'=================================================================================================================================='
	dim s
	set s = CreateObject("MSXML2.XMLHTTP")
	
	'Pass the variable to var1 to be able to call Vmix API and run this api and show this input
	s.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input=" & vmixinput, 0
	s.send 
	
'declare varable for Arg and var1
	dim o 
	set o = CreateObject("MSXML2.XMLHTTP")
	
	'this will start the countdown of the timer
	o.open "GET", "http://localhost:8088/api/?Function=StartCountdown&Input=" & vmixinput, 0
	o.send 
'======================================================================================================================================================================'
' wait for five sec then start recording	
	Wscript.Sleep(5000)
	Dim oShell 
	Dim var1 
	set OShell = CreateObject("Wscript.Shell")
	'Read the variable stored in the VariableofVmixInput then be able to pass that variable to api link
	var1 = oShell.RegRead("HKCU\Storeviewinputs\inputsofviews")
	dim c
	set c = CreateObject("MSXML2.XMLHTTP")
	'switch to the orignal input that the user choose 
	c.open "GET", "http://localhost:8088/api/?Function=ActiveInput&Input=" & var1, false
	c.send 
'========================================================================================================================================================================='
	' start recording right away after the timer ended. 	
	dim rec
	set rec = CreateObject("MSXML2.XMLHTTP")
	rec.open "GET", "http://localhost:8088/api/?Function=StartRecording", 0
	rec.send 
'======================================================================================================================================================================='
	
	dim r 
	set r = CreateObject("MSXML2.XMLHTTP")
	
	'reset the timer when switched out so it wont show 00 timer when pressed again.  
	r.open "GET", "http://localhost:8088/api/?Function=StopCountdown&Input=" & vmixinput, 0
	r.send 
	
	