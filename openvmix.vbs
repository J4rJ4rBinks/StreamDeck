'------------------------------------------------------------------------------------------------------------------
' | This script runs by Streamdeck to open the vMix only one time instead of overloading with many applications |
'-----------------------------------------------------------------------------------------------------------------
Dim Network
Set Network = CreateObject("Wscript.network")
Hostname = Network.ComputerName
	Username = Network.UserName					'Stores the username person logged in with. This is used to complete a directory path (Path_vMixConfiguration)
	'Directory of the presets for the vMix
	TEMP_Path_VideoLab    = "C:\Users\" & username & "\AppData\Local\Temp\VideoLab.vmix"		'vMix have presets that remembers which inputs were opened and what overlays are being used. For Video Lab.		
	TEMP_Path_VideoStudio = "C:\Users\" & username & "\AppData\Local\Temp\VideoStudio.vmix"		'vMix have presets that remembers which inputs were opened and what overlays are being used. For Video Studi
	
	'BOTH NETWORK PATH NEED TO BE IN SHARE FOLDER 	
	Set WSHShell = WScript.CreateObject("WScript.Shell")
'==== 
set objWMIService = GetObject("winmgmts:")
foundProc = false
procName = "vMix64.exe"
procNameFriend = "Word"

for each Process in objWMIService.InstancesOf("Win32_Process")
If StrComp(Process.Name,procName,vbTextCompare) = 0 then
foundProc=true
End If
Next
If foundProc = True Then
' WScript.Echo "Found Process"
else  
'WScript.Echo "Did not find it" 
Select Case Hostname 
		Case "NLCSRL1"
			WSHShell.Run TEMP_Path_VideoStudio, 2, False		'Runs the preset designed for SRL-Sup-01
		Case "NLCSRLE"
			WSHShell.Run TEMP_Path_VideoLab, 2, False		'Runs the preset designed for Video Lab.
		Case "NLCVIDEOSTUDIO"
			WSHShell.Run TEMP_Path_VideoStudio, 2, False		'Runs the preset designed for Video Studio.
	End Select
End If