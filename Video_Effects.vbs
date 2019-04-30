'--------------------------------
' | This script runs by vMix.vbs |
'--------------------------------

On Error Resume Next

Dim Network, debug
debug = TRUE
Set Network = CreateObject("Wscript.network")

Set fso1 = CreateObject("Scripting.FileSystemObject")
Hostname = Network.ComputerName ' To be able to get the computer unique name to focus on only Video lab and Video Studio and Sprint relay lab(for testing purpose) 
' -------------------------------------------------------------------------------
If(Hostname = "NLCVIDEOLAB") or (Hostname = "NLCVIDEOSTUDIO") or (Hostname = "NLCSRL1")  Then ' MAKE SURE TO LOOK AT THIS FIRST BEFORE RUNNING ANYTHING
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	' Global Variables
	' In some of the variable you may have noticed there is triple quote marks enclosing a string. It's because of the spaces in the directory itself, the triple quotes ensure those spaces are included.	
	
	'Serial Number of Elgato Stream Deck (as shown on back of device labelled "Serial #"
	'Serial Number of Elgato Stream Deck is important to us because in the manifest.json which is user.config that contain information of stream deck profile
	'In the manifest.json will have unique Serial number to point to specific StreamDeck if we tried to match a different streamdeck device into nonmatchable serial number
	'It will automatically create a new folder which we do not want that, it will lead to the orignal default profile for stream deck therefore make sure to match the Videolab and Video Studio serial number to the manifest
	'To find the manifest go to (A:\SRL\01 - Projects\StreamDeck\01 - Projects\ElgatoStudio\StreamDeck\ProfilesV2\7B479231-A9FC-4726-AF67-2CB2D3A2BA85.sdProfile) there is a lot of manifest json in each folder to save the information of streamdeck profile
	'To find the serial number it always at the end of the line with those Serial number below, you can open json.manifest by Notepad++ 
	'That why I save the Stream deck ElagtoStudio to AL26H1A06611 and ElagtoVideoLab to AL01H1A03602(A:\SRL\01 - Projects\StreamDeck\01 - Projects) make sure they match the device too!
	'For the stream deck profile you can edit it in the Stream deck app. For start,stop, pic in pic, Single view, desktop view, Weatherman view, side by side view I used api(A:\SRL\01 - Projects\StreamDeck\01 - Projects\shortcutapiforvmix) 
	'Look the streamdeck profile I made for some examples, for applications you can make a directory to run vbs script in the streamdeck such as View my video, and reset vmix 
	SD_SN_VideoLab    = "AL26H1A06611"
	SD_SN_VideoStudio = "AL01H1A03602"
	
	'Names of various objects such as username, folder names, and file names.
	Username = Network.UserName					'Stores the username person logged in with. This is used to complete a directory path (Path_vMixConfiguration)
	
	'Names of the folders that leads to the vMix's default configuration file location
	Folder_vMixCompany     = "\StudioCoast_Pty_Ltd"									'Parent company of vMix is the name of the first folder (Double check the actual directory before changing this.)
	Folder_vMixUniqueID    = "\vMix64.exe_Url_5l40qvre1y15spmggwt1gelpnmwemdfj"	'vMix's unique identifer is the name of the second folder (Double check the actual directory before changing this.)
	Folder_vMixVersion 	   = "\21.0.0.57"											'vMix's version number is the name of the third folder (Double check the actual directory before changing this.)
	File_vMixConfiguration = "\user.config"										'The configuration file itself (Double check the actual directory before changing this.) 
	vMix_Default_Settings  = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\user.config"		'vMix have a configuration file that saves all of the personalized settings a user made to the software. 
	'FOR VMIX DEFAULT SETTINGS IN THE SHARE FOLDER(\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects) and if you have any issues with the vmix user config it is possible it is corrupted so therefore 
	'You have to create a new user config in the vmix by going in the vmix and go to settings to change to your preference. To export the vmix user config go to settings in the vmix and find export to Share folder(A:\SRL\01 - Projects\StreamDeck\01 - Projects)
	'IF YOU WANT TO TEST THIS CODE IN NLCSRL1 MAKE SURE YOU CHANGE USER.CONFIG Desktop to DOCUMENTS FOR MY Recordings!(C:\Users\[USERNAME]\Desktop\My Recordings) since the many desktop is linked with 6 other monitors will cause some issues so therefore 
	'its easier to change the user.config by open it with notepad++ and change desktop to documents. 
	'"\\cias-files.rit.edu\my_shares\SRL\01 - Projects\StreamDeck\01 - Projects\user.config"
	
	'Directory of vMix's saved configuration file
	Path_VMixCompany       = "C:\Users\" & Username & "\AppData\Local" & Folder_vMixCompany		'Starting location of the vMix configuration dir.
	Path_VMixUniqueID      = Path_VMixCompany & Folder_vMixUniqueID								'Second Folder of the vMix configuration dir.
	Path_VMixVersion       = Path_VMixUniqueID & Folder_vMixVersion								'Third Folder of the vMix configuration dir.
	Path_vMixConfiguration = Path_VMixVersion & File_vMixConfiguration							'Location of the configuration file.

	'BOTH NETWORK PATH NEED TO BE IN SHARE FOLDER 	
	'Videolab and Video studio present is tiny bit difference so you need to change the file in Video lab and Video Studio
	Network_Path_VideoLab = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\VideoLab.vmix" ' remember to update VideoLab.vmix if make any changes(make sure to save the changes in the A:\SRL\01 - Projects\StreamDeck\01 - Projects)
	Network_Path_VideoStudio = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\VideoStudio.vmix" ' remember to update VideoLab.vmix if make any changes (make sure to save the change in the A:\SRL\01 - Projects\StreamDeck\01 - Projects)
	
	
	'Directory of the presets for the vMix
	' Only reason to use this to be able to read and overwrite to the Temp folder
	TEMP_Path_VideoLab    = "C:\Users\" & username & "\AppData\Local\Temp\VideoLab.vmix"		'vMix have presets that remembers which inputs were opened and what overlays are being used. For Video Lab.		
	TEMP_Path_VideoStudio = "C:\Users\" & username & "\AppData\Local\Temp\VideoStudio.vmix"		'vMix have presets that remembers which inputs were opened and what overlays are being used. For Video Studio.

	'Audio config for the Yeti microphone since they have unique id 
	VideoStudiomicrophoneid = "@device:cm:{33D9A762-90C8-11D0-BD43-00A0C911CE86}\wave:{DCC1AC61-3AA0-4822-9AF0-22596210FB6B}"
	VideoLabmicrophoneid = "@device:cm:{33D9A762-90C8-11D0-BD43-00A0C911CE86}\wave:{DDBB5272-BF49-4D09-8442-8B1F9A725E1C}"
	
	'PowerPoint Version
	Dim Powerpointversion
	Powerpointversion = "16.0"

	'Uses Sleep to give vMix software time to fully opens before the webpage is loaded to ensure the port (8088) is opened by then.
	intLengthOfSleep = 1000*10		'Length of sleep 1000 millisecond equals 1 second
	
	
	'*****************************************************************************************************************
	' Variables used for DIRECTORY 
	'*****************************************************************************************************************	
	'====================================================================================================================================================================
	'Directory of Actual Window Manager
	Execute_Actual = """C:\Program Files (x86)\Actual Window Manager\ActualWindowManagerCenter.exe"""   'Location of Actual Window Manager's executable.
		
	'Directory of Actual's configuration FROM SHARE FOLDER
	Path_ActualScript = """\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\actualScript.bat"""
	
	'Directory of Actual's configuration FROM SHARE FOLDER
	Path_ActualScriptVS = """\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\actualScriptVS.bat"""
		
	'Directory of StreamDeck
	StreamDeckdir = """C:\Program Files\Elgato\StreamDeck\StreamDeck.exe"""
	
	'Directory of My Recordings, used for able to create My Recordings Folder in the Desktop 
	MyRecordingsfolder = "C:\Users\"& username & "\Desktop\My Recordings"
	
	'Directory of Elgato folder, Used for delete folder in the Roaming in app data, copy to that directory
	ElgatoFolder = "C:\Users\" & username & "\AppData\Roaming\Elgato"
	
	'Directory of Elgato Folder that we make copy from the SHARE FOLDER for Video studio --------
	shareElgatoFolder = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\ElgatoStudio"
	
	'Directory of Elgato Folder that we make copy from the SHARE FOLDER for Video Lab 
	shareElgatoFolderVL = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\ElgatoVideoLab" 
	
	'Directory of Actual Tools, used for delete folder in the Roaming in app data, copy to that directory 
	ActualToolsfolder = "C:\Users\" & username & "\AppData\Roaming\Actual Tools"
	
	'Directory of Actual Tools Folder that we make copy from SHARE FOLDER 
	shareActualFolder = "\\ntidlabs\nlc system files$\Lab\Scripts\Video_Effects\Actual Tools"
	
	'Directory of Actual Tools, to copy the share folder to that directory
	userActualTools = "C:\Users\" & username & "\AppData\Roaming\"
	'===================================================================================================================================================================
	
	' Code start below
	
	'=================================================================================================================================
	' Now we start running actual codes from here on...
	'  1. If not yet exists, create folder titled "My Recordings" on user desktop
	'  2. Delete Elgato StreamDeck user profile "manifest.json"
	'  3. Overwrite Actual Software user profile
	'  4. Run Actual Tools
	'  5. Delete vmix profile folder in appdata/local
	'  6. Recreate vmix user folder in appdata/local
	'  7. Recreate user.config for vmix profile(usually have error if vmix is not open because user.config is corrupted so need new user.config) and read the file to make some change to username 
	'  8. Load Vmix profile and run Vmix 

	'=================================================================================================================================
	'Create folder titled "My Recordings" 
	' 1. If not yet exists, create folder titled "My Recordings" on user desktop, MyRecordingsfolder variable located here 
	' Check to see if recording folder exists
	If Not fso.FolderExists(MyRecordingsfolder) Then
		'Create My Recordings Folder in desktop 
		fso1.CreateFolder MyRecordingsfolder
	End If

'=================================================================================================================================
	' 2. Delete Elgato StreamDeck user profile "manifest.json" 
	
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	'To be able to delete everything in the folder and the folder included 
	objFSO.DeleteFile(ElgatoFolder)
	objFSO.DeleteFolder(ElgatoFolder)
	'call debugmsg("To begin recording, please use Stream Deck Console") 

'=================================================================================================================================
'Copy & Paste Actual Windows Manager Configuration Folder in AppData/Roaming
	' 3. Overwrite Actual Software user profile
	'To be able to delete everything in the folder and the folder included 
	objFSO.DeleteFile(ActualToolsfolder)
	objFSO.DeleteFolder(ActualToolsfolder)
	'call debugmsg("check the actual tools folder in roaming folder") 
	
	'then copy actual window folder to user appdata
	' If you want make some changes to Actual Window, Make sure to open the Actual Tools Window configuration to make some changes to users preference. I made two specific settings for the Streamdeck and Vmix to be able 
	' to minimize the screen to the tray however for the brand new user it will show the streamdeck and then minimize to taskbar but for the second user it wont show the streamdeck. To make some changes for the Actual tools window you need to export 
	' tools folder then click on configuration, click on backup and make sure to overwrite in (A:\SRL\01 - Projects\StreamDeck\01 - Projects\Actual Configure.zip) then extract the Actual Configure.zip to Actual tools folder then make sure to change the name
	' to Actual Window Manager
	'A:\SRL\01 - Projects\StreamDeck\01 - Projects\Actual Tools
	Set copyActualManagerFolder = CreateObject("Scripting.FileSystemObject")
	copyActualManagerFolder.GetFolder(shareActualFolder).Copy userActualTools
	Set copyActualManagerFolder = Nothing
	

'=================================================================================================================================
	
	'Actual Window Manager software
	' 4. Run Actual Tools ( this is not running in the background so we have to call it) 
		Select Case Hostname 
		case "NLCSRL1" 
			WSHShell.run Path_ActualScriptVS, 0, False ' In the acutalscriptVS I put the register license for Video Studio since it wont remember automatically for each user  
		case "NLCVIDEOLAB"	
			WSHShell.run Path_ActualScript, 0, False ' In the acutalscript I put the register license for Videolab since it wont remember automatically for each user  
		case "NLCVIDEOSTUDIO"
			WSHShell.run Path_ActualScriptVS, 0, False 'In the acutalscriptVS I put the register license since it wont remember automatically for each user  
	end select 
	
	WSHShell.run Execute_Actual, 0, False ' Will start running the actual window manager
	
'*****************************************************************************************************************
	
'=================================================================================================================================
	'Before opening vMix, replace user.config file. This is done to assure consistent environment and user's experience
	Dim fso, MyFile, objShell
	'Creates a variable that contains the configuration file for copying operation.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.GetFile(vMix_Default_Settings)
	Set objShell =  CreateObject("Shell.Application")

'  5. Delete vmix profile folder in appdata/local
	objFSO.DeleteFolder(Path_VMixCompany)	

'  6. Recreate vmix user folder in appdata/local
	Call PreppingDirectoryForVMixConfigurationFile()
	MyFile.Copy Path_vMixConfiguration, True				   	'Copies the configuration file to the directory	 
	'call debugmsg("LOOK AT THE CONFIG FILE")
	
'  7. Recreate user.config for vmix profile
	Call writeVMixConfig(username) 'IN THE FUNCTION CODE BELOW
	'call debugmsg("please check the userconfig.")
	Select Case Hostname 
		case "NLCSRL1" 
			call WriteVmixpresetstudio(username) 'to the function below 
		case "NLCVIDEOLAB"	
			call WriteVmixpreset(username) 'to the function below 
		case "NLCVIDEOSTUDIO"
			call WriteVmixpresetstudio(username)'to the function below 
	end select
	
	call debugmsg("To begin recording, please use Stream Deck Console") 

'  8. Load vMix preset file and run vMix
	'Each computer require a different preset file which is dealt with by Select Case
	Select Case Hostname 
		Case "NLCSRL1"
			WSHShell.Run TEMP_Path_VideoStudio, 2, False		'Runs the preset designed for SRL-Sup-01
		Case "NLCVIDEOLAB"
			WSHShell.Run TEMP_Path_VideoLab, 2, False		'Runs the preset designed for Video Lab.
		Case "NLCVIDEOSTUDIO"
			WSHShell.Run TEMP_Path_VideoStudio, 2, False		'Runs the preset designed for Video Studio.
	End Select
End If

'	9. Pass the variable 10 to the registry to be able to store the variable for SRL_buttonsactions.vbs.
	  ' Create a WSH Shell object:
	Set registry = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	registry.RegWrite "HKCU\StoreAPIVariable\", ""

	' Create a new DWORD value to store the string of variable:
	registry.RegWrite "HKCU\StoreAPIVariable\VariableofVmixInput", "10", "REG_SZ"
	
	'10. write the variable 6 as a default to the registry to be able store the variable for writetoregistryforvmixinput
	
	Set registry = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	registry.RegWrite "HKCU\Storeviewinputs\", ""

	' Create a new DWORD value to store the string of variable:
	registry.RegWrite "HKCU\Storeviewinputs\inputsofviews", "6", "REG_SZ"
	
	'11. Prevent the PowerPoint to interfer the Vmix Software
	Dim Powerpointreg
	' Create a WSH Shell object:
	Set Powerpointreg = Wscript.CreateObject( "WScript.Shell" )
	'
	' Create a new key for the Registry to be able to store the variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\16.0\PowerPoint\Options", ""

	' Create a new DWORD value to store the string of variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\" + Powerpointversion +"\PowerPoint\Options\UseAutoMonSelection", "0", "REG_DWORD"
	'
	' Create a new key for the Registry to be able to store the variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\" + Powerpointversion +"\PowerPoint\Options", ""

	' Create a new DWORD value to store the string of variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\" + Powerpointversion +"\PowerPoint\Options\UseMonMgr", "0", "REG_DWORD"
	
	' Create a new key for the Registry to be able to store the variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\" + Powerpointversion +"\PowerPoint\Options", ""

	' Create a new DWORD value to store the string of variable:
	Powerpointreg.RegWrite "HKCU\Software\Microsoft\Office\" + Powerpointversion +"\PowerPoint\Options\DisplayMonitor", "\\.\DISPLAY1", "REG_SZ"
	

'*****************************************************************************************************************
' Subroutines
'*****************************************************************************************************************

'Creates the Configuration Directory and copies the configuration file to it
Function PreppingDirectoryForVMixConfigurationFile()
	fso.CreateFolder(Path_VMixCompany)
	fso.CreateFolder(Path_VMixUniqueID)
	fso.CreateFolder(Path_VMixVersion) 
End Function

'*****************************************************************************************************************
'Debugging message(USE TO FIND ERROR) 
'you need to add " Call desbugmsg("whatever you want to put here") " 
'you need to add them in before or after the line code is running to verify that line code works  
Function debugmsg(msg)
	IF debug = True THEN
		wscript.echo msg
	End IF
	
'	If Err.Number <> 0 Then
'		'error handling:
'		WScript.Echo Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
'		Err.Clear
'	End If
End Function

'*****************************************************************************************************************
'Write user.config to vmix folder in AppData\Local

Function writeVMixConfig(username)
	Const ForReading = 1
	Const ForWriting = 2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'call debugmsg("make it to userconfig")
	Set objFile = objFSO.OpenTextFile(vMix_Default_Settings, ForReading) ' reading the file from the ntid share
	'call debugmsg("read userconfig")
	strText = objFile.ReadAll ' read all the file 
	objFile.Close
	strNewText = Replace(strText, "[USERNAME]", username) ' Try to find the USERNAME only , make change to username
	'call debugmsg("make it to userconfig")
	select case Hostname
		case "NLCVIDEOLAB"
		strNewText = Replace(strText, "[AUDIOCONFIG]", VideoLabmicrophoneid)
		case "NLCVIDEOSTUDIO"
		strNewText = Replace(strText, "[AUDIOCONFIG]", VideoStudiomicrophoneid)
	end select 
		
	
	Set objFile = objFSO.CreateTextFile(Path_vMixConfiguration, ForWriting) ' write to user.config to (C:\Users\(random people username)\AppData\Local\StudioCoast_Pty_Ltd\vMix64.exe_Url_5l40qvre1y15spmggwt1gelpnmwemdfj\21.0.0.55\user.config) 
	
	objFile.WriteLine strNewText
	'call debugmsg("write it to userconfig")
	
	objFile.Close 
End Function

Function WriteVmixpreset(username)
	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'debugmsg("ERROR")
	Set objFile = objFSO.OpenTextFile(Network_Path_VideoLab, ForReading) ' reading the vmix present file 
		'debugmsg("pass")
		strText = objFile.ReadAll
		objFile.Close
		strNewText = Replace(strText, "[USERNAME]", username) ' change the USERNAME depends on the person login
	'debugmsg("pass")
	Set objFile = objFSO.CreateTextFile(TEMP_Path_VideoLab, ForWriting)
	
	objFile.WriteLine strNewText
	'debugmsg("pass")
	objFile.Close 
End Function

Function WriteVmixpresetstudio(username)
	Const ForReading = 1
	Const ForWriting = 2
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'debugmsg("ERROR")
	Set objFile = objFSO.OpenTextFile(Network_Path_VideoStudio, ForReading) 'reading the vmix present file 
		'debugmsg("pass")
		strText = objFile.ReadAll
		objFile.Close
		strNewText = Replace(strText, "[USERNAME]", username) ' change the USERNAME depends on the person login
	'debugmsg("pass")
	Set objFile = objFSO.CreateTextFile(TEMP_Path_VideoStudio, ForWriting)
	
	objFile.WriteLine strNewText
	'debugmsg(strNewText)
	objFile.Close 
End Function
