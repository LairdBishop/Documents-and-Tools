' //******************************************************************************************************************
' // Author: 	Amar Maouche - Dell IMS
' // Version:		22.01.2017
' //
' // Notes: 	 ***
' // 
' //*******************************************************************************************************************

' initialize the objects
	
	Public Const ForReading = 1
	Public Const ForWriting = 2
	Public Const ForAppending = 8

	Public Const Success = 0
	Public Const Failure = 1
	
	Dim sIsLaptop, sIsDesktop, sIsServer, sIsOnBattery, sAssetTag, sSerialNumber, sMake, sModel, sProduct, sUUID, sMemory
	Dim sArchitecture, sProcessorSpeed, sCapableArchitecture, bIsUEFI, bSupportsSLAT, bSupportsX64, bSupportsX86
	Dim sAppsID(), dic_AppsNameCmd, dic_AppsWrkgDirBootFlag, dic_DeviceIDs, IsInstallPerPNPId
	Dim TasksFile, sRebootFlag, bLastAppInstDone, TotalApp
	
	bLastAppInstDone="false"
	sRebootFlag="false"

	Dim UACFlagFile
	

'set a dictionary object to hold the device IDs of current system
	
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set oShell = WScript.CreateObject("WScript.Shell")
	Set objWMI = Nothing
	Set objExplorer = Nothing
	Set dic_AppsNameCmd=CreateObject("Scripting.Dictionary")
	Set dic_AppsWrkgDirBootFlag=CreateObject("Scripting.Dictionary")
	Set dic_DeviceIDs = CreateObject("Scripting.Dictionary")
	'On error resume next
	Set objWMI = GetObject("winmgmts:")
	On Error Goto 0
	
	
	strComputer="."

' set the script directory

	'On error resume next
	bDebug= False
	strComputer= "."
	Dim iRetVal
	iRetVal= Success
	
	
' set the script directory
	sScriptDir = empty
	sScriptDir = WScript.ScriptFullName
	sScriptName= objFSO.GetFileName(sScriptDir)
	sScriptDir = Left(sScriptDir, InStrRev(sScriptDir, "\"))
	
	oShell.CurrentDirectory = sScriptDir
	
' get the root system drive letter
	Set oEnv = oShell.Environment("PROCESS")
	RootDrv = oEnv("SystemDrive")
	s_WinDir = oEnv("WINDIR")
	
	oEnv("SEE_MASK_NOZONECHECKS") = 1
	
'set Tasks.xml file
	TasksFile =sScriptDir & "Tasks.xml"
	
'set splash HTa file
	sSplashHTA =sScriptDir & "hta\Splash.HTA"
	sSplashErrorHTA	=sScriptDir & "hta\SplashErr.HTA"
	
' process arguments	
	s_Args=""
	Set objArgs = WScript.Arguments	
		For I = 0 to objArgs.Count - 1
		   If I=0 Then
		   		s_Args=objArgs(I)
		   Else
		   		s_Args=s_Args & " " & objArgs(I)
		   		
		   End If
	Next
	
	
	If Instr(1, UCase(s_Args), "/POST", 1) > 0 Then
		
		Set NamedArgs = Wscript.Arguments.Named 
		
		sArgPost=Trim(NamedArgs("POST"))
		
		If Len(sArgPost) > 0 Then
			sPhase="POST"
			
			'get the OSDProfile.ini from /POST argument						    			
			sOSDProfileIniFile= sArgPost
			
			If Not Len(Trim(sOSDProfileIniFile))> 0 Then sOSDProfileIniFile = sScriptDir & "OSDProfile.ini"	
		
		Else
		
			sOSDProfileIniFile = sScriptDir & "OSDProfile.ini"	
		End If
		
	ElseIf Instr(1, UCase(s_Args), "/AFTER_REBOOT", 1) > 0 Then
		sPhase="AFTER_REBOOT"
	ElseIf Instr(1, UCase(s_Args), "/AFTER_DISABLEUA", 1) > 0 Then
		sPhase="AFTER_DISABLEUA"
	Else
	
		sPhase="STANDALONE"
		
		' set the default OSDProfile.ini file path
		sOSDProfileIniFile = sScriptDir & "OSDProfile.ini"
		    
	End If

'set log file and initiate logging

	LogFile ="OSDCustomizer.log"
	sLogPath=sScriptDir		
	LogFile = sLogPath & LogFile

	'create a new log file if not exist
	If objFSO.FileExists (LogFile) Then
			WriteLog "Script " & sScriptName & " execution started."
	Else
		'create a new log file
			Set oLog= objFSO.CreateTextFile(LogFile, True)
			oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "Script " & sScriptName & " execution.")
			oLog.Close
	End If
	
	If sPhase="POST" Then
		WriteLog "Execution Phase:Post sysprep First execution."
		RemoveOSDCustomizerFromStartRun
		
	ElseIf sPhase="AFTER_REBOOT" Then
		WriteLog "Execution Phase: AFTER_REBOOT execution. Script re-executed due to a previous reboot required by Dell OSDCustomizer."
	ElseIf sPhase="AFTER_DISABLEUA" Then
		WriteLog "Execution Phase: AFTER_DISABLEUA execution. Script re-executed due to a previous reboot required by Dell OSDCustomizer."
	Else
		WriteLog "Execution Phase:Initial standalone execution."
	End if

	
	If sPhase="AFTER_REBOOT" Then
			
			'continue execution of each command line within Tasks.xml file
			Install_Apps TasksFile
			
			If iRetval=Failure Then
			
				WriteLog "Error occured during task execution." 
				WriteLog "====================================================================================="
				WriteLog "Script " & sScriptName & " execution is aborted."
				WriteLog "====================================================================================="
				
				MsgBox "Error occured during task execution. Process aborted.",vbSystemModal, "Error"
				'close current splash
				
				'iRetval= killProcess ("mshta.exe")	
				WScript.Sleep 100
				WScript.Quit
				
			End If
			  
			'check if reboot required
			
			If lcase(sRebootFlag)="true" Then
		
				If Not bLastAppInstDone="true" Then
				
						'add a shortcut at startup folder if execution not complete
						UpdateStartupLink
						
						WriteLog "A reboot is required. Script " & sScriptName & " execution will continue at next reboot."
							
						'do a system reboot 
						sCmd="shutdown.exe /r /t 5 /f /c " & Chr(34) & "System reboot required by Dell OSDCustomizer tool." & Chr(34)
						
						Result= RunWithHeartbeat(sCmd)
							
						If Err Then
							
							Result = Err.number
							sError = Err.Description
						    WriteLog "Error with execution of command:" & sCmd & ": " & sError
						        
						ElseIf Result = 0 Then
							
					        WriteLog "Command:" & sCmd & " executed successfully"
						Else
					        WriteLog "Command:"& sCmd & " returned an unexpected return code: " & Result
						End If
						
						
						'exit script
						 WScript.Quit
				
				End If
							
			End If
			
	Else
	
		If sPhase="AFTER_DISABLEUA" Then
			'bypass the check of UA and do not reboot
			
		Else
		
					'Check if current user has administrator ritghs	
				Set objNetwork = CreateObject("Wscript.Network")
				strUser = objNetwork.UserName
			 
				If IsCurrentUserAdmin(strUser) = Failure Then
						Writelog strUser & " user account is not a local administrator account. Administrator rights required for task execution."
						WriteLog "====================================================================================="
						WriteLog "Script " & sScriptName & " execution is aborted."
						WriteLog "====================================================================================="
						
						'remove shortcut from Startup and remove credentials from autologon registry
						CleanupStartItems
						WScript.Quit
						
				End If
				
				'check UAC and disable it if enabled
				UACFlagFile=sScriptDir	& "UAC.flg"
				sUACKeyPath ="HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA"
				
				If customRegRead(sUACKeyPath) = 0 Then
					WriteLog "Disabling UAC and rebooting system."
					
					'save existing UAC value
						
						If Not objFSO.FileExists (UACFlagFile) Then
							Set oUACflg= objFSO.CreateTextFile(UACFlagFile, True)
							oUACflg.WriteLine("0")
							oUACflg.Close
						End If
						
					'disable UAC
					oShell.RegWrite sUACKeyPath, "1", "REG_DWORD"
				
					' reboot
					'add a shortcut at startup folder if execution not complete
									UpdateStartupLink_DISABLE_UA
									WriteLog "A reboot is required. Script " & sScriptName & " execution will continue at next reboot."
										
									'do a system reboot 
									sCmd="shutdown.exe /r /t 5 /f /c " & Chr(34) & "System reboot required by Dell OSDCustomizer tool." & Chr(34)
									
									Result= RunWithHeartbeat(sCmd)
										
									If Err Then
										
										Result = Err.number
										sError = Err.Description
									    WriteLog "Error with execution of command:" & sCmd & ": " & sError
									        
									ElseIf Result = 0 Then
										
								        WriteLog "Command:" & sCmd & " executed successfully"
									Else
								        WriteLog "Command:"& sCmd & " returned an unexpected return code: " & Result
									End If
									
									'exit script
									 WScript.Quit
				Else
					WriteLog "UAC already disabled."
					
				End If
		End If
				
		' set the Apps.xml file path
			sAppsXmlFile ="Apps.xml"
			If objFSO.FileExists(sScriptDir & "cfg\" & "Apps.xml") Then
				sAppsXmlFile =sScriptDir & "cfg\" & "Apps.xml"
			Else
				WriteLog "Error: File not found: " &  sAppsXmlFile 
				WriteLog "====================================================================================="
				WriteLog "Script " & sScriptName & " execution is aborted."
				WriteLog "====================================================================================="
				
				MsgBox "Error: File not found: " &  sAppsXmlFile & ". Process aborted.",vbSystemModal, "File not found"
				
				WScript.Quit  
			End If
				
		
		'get asset info
		iRetVal = GetAssetInfo
		
		' check if running in Windows PE

		If UCase(RootDrv)="X:" Then
			WriteLog "Running Application installation is not supported in Windows PE. Script execution aborted."
	    	MsgBox "Running Application installation from Windows PE is not supported. Script execution aborted.",vbSystemModal, "Warning"
	    	
		Else

			'get Application or command entries from OSDProfile.ini
			WriteLog "Calling Sub: getAppIdFromProfile" 
			If getAppIdFromProfile = True Then
				
				For i=0 To ubound(sAppsID)

					WriteLog  "Application(" & i & ")=" & sAppsID(i)
					
				Next
					
				'get PNP Ids for current system if PNPId is used in apps.xml
				Set oAppsDoc = CreateObject("Microsoft.XMLDOM") 
				oAppsDoc.async = False 
				oAppsDoc.load(sAppsXmlFile)
				IsInstallPerPNPId=False
					   
				Set nodesList = oAppsDoc.documentelement.childNodes
				For each node in nodesList	
					'MsgBox "node.selectSingleNode(""PNPId"").text=" & node.selectSingleNode("PNPId").text
					
					If Len(node.selectSingleNode("PNPId").text) > 0 And node.selectSingleNode("PNPId").text <> "*" Then
						GetDeviceIDs dic_DeviceIDs
						IsInstallPerPNPId=True
						Exit For
					End If
					
				Next
				
				'Populate apps dictionnary and create Tasks.xml file
				Populate_Apps_Dic sAppsID
			
			If TotalApp > 0 Then
				
				' Create Tasks.xml
				Create_Tasks_Xml sAppsID
				
				'start execution of each command line within Tasks.xml file
				Install_Apps TasksFile
				
				If iRetval=Failure Then
				
					WriteLog "Error occured during task execution." 
					WriteLog "====================================================================================="
					WriteLog "Script " & sScriptName & " execution is aborted."
					WriteLog "====================================================================================="
					
					MsgBox "Error occured during task execution. Process aborted.",vbSystemModal, "Error"
					'close current splash
				
					'iRetval= killProcess ("mshta.exe")	
					WScript.Sleep 100
					
					WScript.Quit
					
				End If
				
				
				'check if reboot required
				If lcase(sRebootFlag)="true" Then
					
						If Not bLastAppInstDone="true" Then
								WriteLog "A reboot is required. Script " & sScriptName & " execution will continue at next reboot."
								
								'add a shortcu at startup folder
								UpdateStartupLink
			
								'Restart Computer
									sCmd="shutdown.exe /r /t 5 /f /c " & Chr(34) & "System reboot required by Dell OSDCustomizer tool." & Chr(34)
									
									Result= RunWithHeartbeat(sCmd)
					
									If Err Then
									
										Result = Err.number
										sError = Err.Description
								        writeLog "Error with execution of command:" & sCmd & ": " & sError
								        
									ElseIf Result = 0 Then
									
							                writeLog "Command:" & sCmd & " executed successfully"
									Else
							               writeLog "Command:"& sCmd & " returned an unexpected return code: " & Result
									End If
							
								'exit script
								 WScript.Quit
						End If
							
				End If
			Else
				WriteLog  "No Application or command to execute for this system."
				bLastAppInstDone="true"
			End If
							
			Else
					
				WriteLog  "No Application or command found on OSDProfile.ini."
				bLastAppInstDone="true"
					
			End If
					
		End If
				
	End If
	
'script completed

	If bLastAppInstDone="true" Then
		
		'remove shortcut from Startup and remove credentials from autologon registry
			CleanupStartItems
			
		're-apply UAC saved value
	
		If objFSO.FileExists (UACFlagFile) Then
			Set oUACFlg = objFSO.OpenTextFile(UACFlagFile, 1, False)
				
			Do While (not oUACFlg.AtEndOfStream)
				line = Trim(oUACFlg.ReadLine)
			Loop
		
			oUACFlg.Close
			
			'set UAC saved value
			oShell.RegWrite sUACKeyPath, line, "REG_DWORD"
			objFSO.DeleteFile UACFlagFile
		End If
		
		WriteLog "Script " & sScriptName & " execution is completed."
		
		Writelog "====================================================================================="
		WriteLog "POST Sysprep task execution completed."
		WriteLog "OSDCustomizer process execution is completed."
		WriteLog "====================================================================================="
		
		
		'Cleanup files and copy logs to c:\windows\temp\OSDCustomizer
		iRetval= killProcess ("SetTop.exe")	
		oShell.CurrentDirectory = s_WinDir

		sCmd="cmd /c XCOPY /y " & sScriptDir & "*.log" & " " & s_WinDir & "\Panther\DTRI\OSDCustomizer\"
		Result= RunWithHeartbeat(sCmd)
				If Err Then
					Result = Err.number
					sError = Err.Description
			        writeLog "Error with execution of command:" & sCmd & ": " & sError
			        
				ElseIf Result = 0 Then
				
		                writeLog "Command:" & sCmd & " executed successfully"
				Else
		               writeLog "Command:"& sCmd & " returned an unexpected return code: " & Result
				End If
		oEnv.Remove("SEE_MASK_NOZONECHECKS")
		
		'Restart Computer
		sCmd="shutdown.exe /r /t 10 /f /c " & Chr(34) & "BUILD COMPLETE. Restarting system. Please wait..." & Chr(34)
		oShell.Run sCmd,0,False
			
		
	End if

'exit
	WScript.Quit


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    SUB & FUNCTIONS
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//---------------------------------------------------------------------------
	'//  Function:	GetAssetInfo()
	'//  Purpose:	Get asset information using WMI
	'//---------------------------------------------------------------------------
Function GetAssetInfo

	 	Dim bFoundBattery, bFoundAC
		Dim objResults, objInstance
		Dim i, scmd
		Dim bisX64
		
		WriteLog "Begin Function: GetAseetInfo" 
		
		writelog "Getting local computer information: Asset information"


		' Get the SMBIOS asset tag from the Win32_SystemEnclosure class

		Set objResults = objWMI.InstancesOf("Win32_SystemEnclosure")
		bIsLaptop = false
		bIsDesktop = false
		bIsServer = false
		For each objInstance in objResults

			If objInstance.ChassisTypes(0) = 12 or objInstance.ChassisTypes(0) = 21 then
				' Ignore docking stations
			Else

				If not IsNull(objInstance.SMBIOSAssetTag) then
					sAssetTag = Trim(objInstance.SMBIOSAssetTag)
				End if
				Select Case objInstance.ChassisTypes(0)
				Case "8", "9", "10", "11", "12", "14", "18", "21"
					bIsLaptop = true
				Case "3", "4", "5", "6", "7", "15", "16"
					bIsDesktop = true
				Case "23"
					bIsServer = true
				Case Else
					' Do nothing
				End Select

			End if

		Next
		If sAssetTag = "" then
			writelog "Unable to determine asset tag via WMI."
		End if


		' Get the serial number from the Win32_BIOS class.

		Set objResults = objWMI.InstancesOf("Win32_BIOS")
		For each objInstance in objResults

	                ' Get the serial number

			If not IsNull(objInstance.SerialNumber) then
				sSerialNumber = Trim(objInstance.SerialNumber)
			End if

		Next
		If sSerialNumber = "" then
			writelog "Unable to determine serial number via WMI."
		End if


		' Figure out the architecture from the environment

		If oEnv("PROCESSOR_ARCHITEW6432") <> "" then
			If UCase(oEnv("PROCESSOR_ARCHITEW6432")) = "AMD64" then
				sArchitecture = "X64"
			Else
				sArchitecture = UCase(oEnv("PROCESSOR_ARCHITEW6432"))
			End if
		ElseIf UCase(oEnv("PROCESSOR_ARCHITECTURE")) = "AMD64" then
			sArchitecture = "X64"
		Else
			sArchitecture = UCase(oEnv("PROCESSOR_ARCHITECTURE"))
		End if
		
		

		' Get the processor speed from the Win32_Processor class.

		bSupportsX86 = false
		bSupportsX64 = false
		bSupportsSLAT = false
		Set objResults = objWMI.InstancesOf("Win32_Processor")
		For each objInstance in objResults

			' Get the processor speed

			If not IsNull(objInstance.MaxClockSpeed) then
				sProcessorSpeed = Trim(objInstance.MaxClockSpeed)
			End if


			' Determine if the machine supports SLAT (only supported with Windows 8)

			On error resume next
			bSupportsSLAT = objInstance.SecondLevelAddressTranslationExtensions
			On Error Goto 0
			

			' Get the capable architecture

			If not IsNull(objInstance.Architecture) then
				Select Case objInstance.Architecture
				Case 0
					sCapableArchitecture = "X86"
					bSupportsX86 = true
				Case 6
					sCapableArchitecture = "IA64"
				Case 9
					sCapableArchitecture = "AMD64 X64 X86"
					bSupportsX86 = true
					bSupportsX64 = true
				Case Else
					SCapableArchitecture = "Unknown"
				End Select
			End if


			' Stop after first processor since all should match

			Exit For

		Next
		If sProcessorSpeed = "" then
			writelog "Unable to determine processor speed via WMI."
		End if
		If sCapableArchitecture = "" then
			writelog "Unable to determine capable architecture via WMI."
		End if


		' Get the make, model, and memory from the Win32_ComputerSystem class

		Set objResults = objWMI.InstancesOf("Win32_ComputerSystem")
		For each objInstance in objResults

			If not IsNull(objInstance.Manufacturer) then
				sMake = Trim(objInstance.Manufacturer)
			End if
			If not IsNull(objInstance.Model) then
				sModel = Trim(objInstance.Model)
			End if
			If not IsNull(objInstance.TotalPhysicalMemory) then
				sMemory = Trim(Int(objInstance.TotalPhysicalMemory / 1024 / 1024))
			End if

		Next
		If sMake = "" then
			writelog "Unable to determine make via WMI."
		End if
		If sModel = "" then
			writelog "Unable to determine model via WMI."
		End if


		' Get the UUID from the Win32_ComputerSystemProduct class

		Set objResults = objWMI.InstancesOf("Win32_ComputerSystemProduct")
		For each objInstance in objResults

			If not IsNull(objInstance.UUID) then
				sUUID = Trim(objInstance.UUID)
			End if

		Next
		If sUUID = "" then
			writelog "Unable to determine UUID via WMI."
		End if


		' Get the product from the Win32_BaseBoard class

		Set objResults = objWMI.InstancesOf("Win32_BaseBoard")
		For each objInstance in objResults

			If not IsNull(objInstance.Product) then
				sProduct = Trim(objInstance.Product)
			End if

		Next
		If sProduct = "" then
			writelog "Unable to determine product via WMI."
		End if

		' Determine if we are running UEFI

		bIsUEFI = False

		If oEnv("SystemDrive") = "X:" Then   'running from WinPE
			'check if UEFI from registry 
			scmd="cmd /c wpeutil UpdateBootInfo"
			sReturn = oShell.Run(scmd , 0, True)
			On Error Goto 0
			If sReturn = 0 Then 
					'read the value of reg HKLM\System\CurrentControlSet\Control\PEFirmwareType
					
					sregPEFirmwareType="HKLM\System\CurrentControlSet\Control\PEFirmwareType"
					
					writelog "determine UEFI or BIOS mode by reading registry key HKLM\System\CurrentControlSet\Control\PEFirmwareType."
					
					sFirmware = oShell.RegRead(sregPEFirmwareType)
					writelog "PEFirmwareType value is:" & sFirmware
					
					If sFirmware="0x1" Then bIsUEFI = False
					If sFirmware="0x2" Then bIsUEFI = True

			Else
					writelog "Unable to determine if running UEFI via registry." & ". Error description :" & Err.Description
				
			End If
			
		Else
		
			'On error resume next
			scmd="cmd /c BCDEDIT.exe /ENUM >" & Chr(34) & oEnv("tmp") & "\BcdeditEnum.txt" & Chr(34)
			sReturn = oShell.Run(scmd , 0, True)
			On Error Goto 0
			If sReturn = 0 Then 
				If objFSO.FileExists (oEnv("tmp") & "\BcdeditEnum.txt") Then
					
					Set ini = objFSO.OpenTextFile( oEnv("tmp") & "\BcdeditEnum.txt", 1, False)
					Do While (not ini.AtEndOfStream)
						line = ini.ReadLine
						line = Trim(line)
						
						If InStr(1, UCase (line), ucase("Path"),1) > o And InStr(1, UCase (line), ucase("\EFI\Microsoft\Boot\bootmgfw.efi"),1) > o Then 
							bIsUEFI = True
							
							Exit Do 
						End If
					Loop
					ini.Close
					writelog "deleting temp file " &  oEnv("tmp") & "\BcdeditEnum.txt"
					objFSO.DeleteFile oEnv("tmp") & "\BcdeditEnum.txt", True
					
				Else
					writelog "NOT found " &  oEnv("tmp") & "\BcdeditEnum.txt"
					writelog "Unable to determine if running UEFI via command BCDEDIT.exe /ENUM."
				End If
			Else
				writelog "Unable to determine if running UEFI via command BCDEDIT.exe /ENUM." & ". Error description :" & Err.Description
				
			End If 
		End If
			
		' See if we are running on battery

		If oEnv("SystemDrive") = "X:" and objFSO.FileExists("X:\Windows\Inf\Battery.inf") then
			
			' Load the battery driver

			oShell.Run "drvload X:\Windows\Inf\Battery.inf", 0, true
			
		End if

		bFoundAC = False
		bFoundBattery = False
		Set objResults = objWMI.InstancesOf("Win32_Battery")
		For each objInstance in objResults
			bFoundBattery = True
			If objInstance.BatteryStatus = 2 then
				bFoundAC = True
			End if 
		Next
		If bFoundBattery and (not bFoundAC) then
			bOnBattery = True
		Else
			bOnBattery = False
		End if
		
		WriteLog "sMake = " &  sMake
		WriteLog "sModel = " &  sModel
		WriteLog "sAssetTag = " &  sAssetTag
		WriteLog "sSerialNumber = " &  sSerialNumber	
		WriteLog "sProduct = " &  sProduct
		WriteLog "sUUID = " &  sUUID
		WriteLog "sMemory = " &  sMemory
		WriteLog "sArchitecture = " &  sArchitecture
		WriteLog "sProcessorSpeed = " &  sProcessorSpeed
		WriteLog "sCapableArchitecture = " &  sCapableArchitecture
		sIsLaptop = ConvertBooleanToString(bIsLaptop)
		WriteLog "sIsLaptop = " &  sIsLaptop
		sIsDesktop = ConvertBooleanToString(bIsDesktop)
		WriteLog "sIsDesktop = " &  sIsDesktop
		sIsServer = ConvertBooleanToString(bIsServer)
		WriteLog "sIsServer = " &  sIsServer
		sIsUEFI = ConvertBooleanToString(bIsUEFI)
		WriteLog "sIsUEFI = " &  ConvertBooleanToString(bIsUEFI)
		sIsOnBattery = ConvertBooleanToString(bOnBattery)
		WriteLog "sIsOnBattery = " &  sIsOnBattery
		sSupportsX86 = ConvertBooleanToString(bSupportsX86)
		WriteLog "sSupportsX86 = " &  sSupportsX86
		sSupportsX64 = ConvertBooleanToString(bSupportsX64)
		WriteLog "sSupportsX64 = " &  sSupportsX64
		
		If bSupportsSLAT or sSupportsSLAT = "" Then
			sSupportsSLAT = ConvertBooleanToString(bSupportsSLAT)
			WriteLog "sSupportsSLAT = " &  ConvertBooleanToString(bSupportsSLAT)
		Else
			writelog "Property SupportsSLAT = " & sSupportsSLAT
		End if

		writelog "Finished getting asset info"

		GetAssetInfo = Success
		
		WriteLog "End Function: GetAseetInfo"
		
End Function

Function getAppIdFromProfile
	
	Dim sSelSection
	Dim vFound
	Dim i
	
  	writelog "Begin execution of Function: getAppIdFromProfile"
  	writelog "Getting selected settings from " & sOSDProfileIniFile
  	getAppIdFromProfile=False
  
	if objFSO.FileExists(sOSDProfileIniFile) = True Then
	
		WriteLog "Getting selected settings from OSDProfile.ini file..."
		i=0
		
		'read  the main selected section
		sSelSection=ReadIni(sOSDProfileIniFile,"Main","Selected")
		WriteLog "Selected=" & sSelSection
		
		'read Application or command lines from selected section of OSDProfile
		' and populate the array sAppsID()
			
		Set oEntries = SectionContents(sOSDProfileIniFile, sSelSection)
			
			For each sEntry in oEntries.Keys
				
				If InStr(1, UCase(sEntry), "Application",1) > 0 Then
					
						If i=0 Then
							ReDim sAppsID(i)
						Else
							ReDim Preserve sAppsID(i)
						End If
						
						sAppsID(i)=oEntries(sEntry)
						If Len(sAppsID(i)) > 0 Then
						 	vFound=True
						 	i=i+1
						End If		
				End If
				
			Next
			
			If vFound=True Then
				getAppIdFromProfile=True
				
				
			Else
				getAppIdFromProfile=False
				
			End If
		
	Else
		
		WriteLog "Error: OSDProfile.ini file not found."
		getAppIdFromProfile=False
	End If	
	
	writelog "End execution of Function: getAppIdFromProfile"		
End Function


Sub Populate_Apps_Dic (sArrayCmd)
'update with the apps Install 
	writelog "Begin execution of Sub: Populate_Apps_Dic"
	Dim tmp_dic
	Dim temparray, node1, text1
	Dim c, i, j, k, x, counter, appFound
	Dim IsSupportedSystem, IsModelName, IsPNPId
	
	
	Set tmp_dic = CreateObject("Scripting.Dictionary")						 
	Set oAppsDoc = CreateObject("Microsoft.XMLDOM") 
	
	oAppsDoc.async = False 
	oAppsDoc.load(sAppsXmlFile)
	TotalApp = 0
	
	For k=0 To ubound(sArrayCmd)
	
		  WriteLog "parsing Application(" & k & ")= " & sArrayCmd(k)
		   
		 Set nodesList = oAppsDoc.documentelement.childNodes
		 counter = nodesList.length  
		 i = 1 
		 
		 For x = 1 To counter 
			  appFound=False
			  
			  Set currNode=nodesList.nextNode  
			  
			  CurrNodeName = currNode.nodename 
			
			  Set cnode = currNode.childnodes 
			  clength = cnode.length  
			   
			  id1=currNode.Attributes(0).value
			  
			  If UCase(id1) = UCase(sArrayCmd(k)) Then
			  
			   	WriteLog "found matching Application or command on apps.xml which is: " & id1
			   	
				For Each c In cnode   
				  
				   node1 = c.nodename  
				   text1 = c.text 
	
				   Select Case node1
				   
				   	Case "ModelName"
				   		
				   		If Len(tmp_dic(node1)) =0 Then
				   			tmp_dic(node1)= text1
				   		Else
				   			tmp_dic(node1)= tmp_dic(node1) & "," & text1
				   		End If
				   		
				   	Case "PNPId"
				   	
				   		If Len(tmp_dic(node1)) =0 Then
				   			tmp_dic(node1)= text1
				   		Else
				   			tmp_dic(node1)= tmp_dic(node1) & "," & text1
				   		End If
				   		
				   	Case Else
				   	
				   		tmp_dic(node1)= text1
				   
				   End Select 'end of Select Case node1
					   
				   If i = clength Then 
					    
					    'WriteLog "Application ID=" & id1 & " found in apps.xml"
					    
					    For each sEntry in tmp_dic.Keys
										    
							'check if apps to be executed the target system
							 
							 Select Case UCase(sEntry)
							 	Case UCase ("Name")
							 		cmdName=tmp_dic(sEntry)
							 	
							 	Case UCase ("Description")
							 		cmdDesc=tmp_dic(sEntry)
							 	
							 	Case UCase ("WorkingDirectory")
							 		cmdWorkDir=tmp_dic(sEntry)
							 	
							 	Case UCase ("Commandline")
							 		cmdLine=tmp_dic(sEntry)
							 	
							 	Case UCase ("Reboot")
							 		rebootFlag=tmp_dic(sEntry)
							 		
							 	
							 	Case UCase("SupportedSystems")
							 	
							 		If tmp_dic(sEntry)="*" Or Len(tmp_dic(sEntry))=0 Then
							 			IsSupportedSystem=True
							 			
							 		Else
							 			
							 			Select Case UCase(tmp_dic(sEntry))
							 			Case "NOTEBOOKS"
							 				If sIsLaptop Then
							 					IsSupportedSystem=True
							 				End If
							 			
							 			Case "DESKTOPS"
							 				If sIsDesktop Then
							 					IsSupportedSystem=True
							 				End If
							 			
							 			Case "SERVERS"
							 				If sIsServer Then
							 					IsSupportedSystem=True
							 				End If
							 			End Select
							 			
							 		End If
							 			
			 		
							 	Case UCase("ModelName")
							 		If tmp_dic(sEntry) = "*" Or Len(tmp_dic(sEntry)) = 0 Then
							 			IsModelName = True
							 			
							 		Else
							 		
							 			'split tmp_dic(sEntry) by ","
							 			'for each element of splitted array check if match with current Model Name
							 			temparray=Split(tmp_dic(sEntry),",")
							 			IsModelName=False
							 			
										For j=0 To UBound(temparray)
							 				If IsMatchModelName(temparray(j)) = True Then
							 					IsModelName=True
							 					Exit for									 					
							 				End If
							 				
							 			Next
							 			
							 		End If
							 		
							 	Case UCase("PNPId")
							 		If tmp_dic(sEntry)="*" Or Len(tmp_dic(sEntry))=0 Then
							 			IsPNPId=True
							 			
							 		Else
							 			
							 			
							 			'split tmp_dic(sEntry) by ","
							 			'for each element of splitted array and check if match with current PNPIds of system
							 			temparray=Split(tmp_dic(sEntry),",")
							 			
							 			
										For j=0 To UBound(temparray)
										
							 				If IsMatchPNPId(temparray(j), dic_DeviceIDs) = True Then
							 					IsPNPId=True
							 					Exit for					 					
							 				End If
							 				
							 			Next
										
							 		End If
							 		
							 End Select	 'end of UCase(sEntry)
									   
					    Next 'end of For each sEntry in tmp_dic.Keys
					    
					    appFound=True
					    
					    If IsSupportedSystem=True And IsModelName=True And IsPNPId=True Then
				 			'add Application or command name , WorkingDirectory\commandline to APPS dictionnary
							cmdWorkDir=Trim(cmdWorkDir)
							
							If Len(cmdWorkDir) > 0 Then
					 
					 			If cmdWorkDir <> "." Then
						
									If InStr(1,UCase(cmdLine), UCase(cmdWorkDir),1)> 0 Then
									'the working directory is already in commandline.
										
									Else
										
										If Right(cmdWorkDir,1)="\" Then
											cmdLine=cmdWorkDir & cmdLine
										Else
											cmdLine=cmdWorkDir &"\" & cmdLine
										End If
									
									End If
							
								Else
									cmdWorkDir=sScriptDir
									cmdLine=sScriptDir &"\" & cmdLine
								End If
								
							Else
								cmdWorkDir=sScriptDir
								cmdLine=sScriptDir &"\" & cmdLine
								
							End If
							
							
							'MsgBox "cmdLine=" &cmdLine
							dic_AppsNameCmd.Add cmdName, cmdLine
							dic_AppsWrkgDirBootFlag.Add cmdWorkDir, rebootFlag
							WriteLog "Application or command " & cmdName & " added to execution list."
						 	TotalApp = TotalApp + 1
	
						 Else
						 	
								WriteLog "Application or command " & cmdName & " execution not applicable for this system."
							
				 		 End If
					    
					    i = 0
					    tmp_dic.RemoveAll()
					    
					 	Exit For
						 	
					  End If 'end of  If i = clength Then 
					   
					  i = i + 1 
					   
				Next 'end of For Each c In cnode  
			  		
			  		
			End If 'end of If UCase(id1) = UCase(sArrayCmd(k)) Then
			  	
		  	If appFound=True Then
		  			Exit For
		  	End If
			  		
				  	
		 Next 'end of For x = 1 To counter 
		
		 If appFound=False Then
		  		WriteLog "NOT Found Application with ID=" & id1 & " in apps.xml"
		  		
		 End If
						
	Next 'end of For i=0 To ubound(sArrayCmd)
	
	WriteLog "End execution of Sub: Populate_Apps_Dic"		
		
End Sub

Sub Create_Tasks_Xml (sArrayCmd)
	WriteLog "Begin Sub: Create_Tasks_Xml"
	Dim arrAppsCmd,arrAppsName, arrAppsWrkgDir, arrAppsRebootFlag, i, y, Result, oNode, node
	'On error resume next
	
	arrAppsCmd=dic_AppsNameCmd.Items
	arrAppsName=dic_AppsNameCmd.Keys
	arrAppsRebootFlag=dic_AppsWrkgDirBootFlag.Items
	arrAppsWrkgDir=dic_AppsWrkgDirBootFlag.Keys
	
	'create a Tasks.xml file
	Set oTasks= objFSO.CreateTextFile(TasksFile, True)
	
	For i = 0 To dic_AppsNameCmd.Count -1
				
		If i=0 Then
			oTasks.WriteLine("<?xml version=" & Chr(34) & "1.0" & Chr(34)& " encoding=" & Chr(34) & "UTF-8" & Chr(34)& " standalone=" & Chr(34) & "yes" & Chr(34) & "?>")
			oTasks.WriteLine("<Tasks>")
		End If
		
		oTasks.WriteLine("	<Task ID=" & Chr(34) & i+1 & Chr(34) & " wasProcessed=" & Chr(34) & "false" & Chr(34) & ">")
		oTasks.WriteLine("		<Name>" & arrAppsName(i) & "</Name>")
		oTasks.WriteLine("		<WorkingDirectory>" & arrAppsWrkgDir(i) & "</WorkingDirectory>")
		oTasks.WriteLine("		<Commandline>" & arrAppsCmd(i) & "</Commandline>")
		oTasks.WriteLine("		<Reboot>" & arrAppsRebootFlag(i) & "</Reboot>")
		oTasks.WriteLine("	</Task>")
		
		WriteLog ("------------------------------------------------------------")
		WriteLog ("Application : " & arrAppsName(i) & " added to Tasks.xml with below settings:")
		WriteLog ("     - WorkingDirectory: " & arrAppsWrkgDir(i) & " added to Tasks.xml")
		WriteLog ("     - Commandline: " & arrAppsCmd(i) & " added to Tasks.xml")
		WriteLog ("     - Reboot: " & arrAppsRebootFlag(i) & " added to Tasks.xml")
		WriteLog ("------------------------------------------------------------")
		
	Next
	
	oTasks.WriteLine("</Tasks>")
	oTasks.Close
	WriteLog "End Sub: Create_Tasks_Xml"	
		
End Sub

Sub GetDeviceIDs (dic_DeviceIDs)

	Dim n
	'On Error Resume Next
	WriteLog "Begin execution of Sub: GetDeviceIDs"
	
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	' get the list of Device IDs using Win32_PnPEntity
	Set colComputerSystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	
	WriteLog "getting the list of PNP IDs using Win32_PnPEntity..."
	WriteLog "******************** Starting List of PnP IDs detected on current system ********************"
		
	For Each objComputer in colComputerSystem

		Set colPnPEntity = objWMIService.ExecQuery("Select * from Win32_PnPEntity ")
		n=0
		For Each objPnP in colPnPEntity
			
	    	dic_DeviceIDs (n) = objPnP.DeviceID
	    	    	WriteLog "Device ID(" & n & ")= " & objPnP.DeviceID
	    	
	    	n=n+1
	    Next
	Next
	WriteLog "******************** Ending List of PnP IDs detected on current system ********************"
	WriteLog "End Sub: GetDeviceIDs"
End Sub


Function IsMatchPNPId(sPnPId, dicDevices)
	
	WriteLog "Begin Function: IsMatchPNPId"
	IsMatchPNPId= False
	For each sElement in dicDevices.Keys  
		
		If InStr (1,UCase(dicDevices(sElement)), UCase(sPnPId), 1) > 0 Then
		
    		WriteLog "Found matching between PnP Id: " & sPnPId & " and : " & dicDevices(sElement)
    		IsMatchPNPId= True
    		Exit For
    	End If
	Next
	WriteLog "Begin Function: IsMatchPNPId"
	
End Function

Function IsMatchModelName(sname)
	WriteLog "Begin Function: IsMatchModelName"
	If InStr(1, UCase(sModel), UCase(sname), 1 ) > 0 Then
		
		IsMatchModelName=True
	Else
		IsMatchModelName=False
	End If
	WriteLog "End Function: IsMatchModelName"
End Function

Sub Install_Apps (sTasksFile)
	
	WriteLog "Begin Sub: Install_Apps"
	
	'loading Taskx.xml file
	
	If sTasksFile <> "" And objFSO.FileExists(sTasksFile) Then
		Set oTasksXml = CreateObject("Microsoft.XMLDOM")
		oTasksXml.async = False
		oTasksXml.load sTasksFile
		iRetVal=Success
		WriteLog "Loaded " & sTasksFile
	else
		WriteLog "File " & sTasksFile & " does not exist."
		sTasksFile=""
		iRetVal=Failure

	end if


	If iRetVal=Success Then
	
		WriteLog "parsing xml file: " & sTasksFile
		Set tmp_dic = CreateObject("Scripting.Dictionary")	  
		Set nodesList = oTasksXml.documentelement.childNodes
		counter = nodesList.length  
	 	i = 1 
		 
		For x = 1 To counter 
			 
			Set currNode=nodesList.nextNode  
			CurrNodeName = currNode.nodename 
			
			Set cnode = currNode.childnodes 
			clength = cnode.length  
			   
			id1=currNode.Attributes(0).value
			sWasProcessed=currNode.Attributes.getNamedItem("wasProcessed").Text
		
			For Each c In cnode   
							  
				node1 = c.nodename  
				text1 = c.text 
	
				Select Case node1
				   	Case "Name"
				   		sAppName=text1
				   	Case "WorkingDirectory"
				   		sWrkgDir=text1
			   		Case "Commandline"
				   		sAppCmd=text1
				   	Case "Reboot"
				   		sRebootFlag=text1
				   		
				End Select
				
										   
				If i = clength Then 
				
					If lcase(Trim(sWasProcessed)) = "false" Then
				
						WriteLog "Executing Application or command name:" & sAppName	
					
						iRetval= Run_Install_App (sAppName, sWrkgDir, sAppCmd)
						
								If iRetval=Failure Then	
										WriteLog "error/warning during execution of :" & sAppName
									
								End If
								
								
								' update wasprocessed ="true" in Tasks.xml
								WriteLog "Updating Taskx.xml file: " & sAppName & " wasProcessed=true"
								currNode.Attributes.getNamedItem("wasProcessed").Text="true"
								oTasksXml.save sTasksFile					    
								
								If LCase(sRebootFlag)="true" Then
										WriteLog "Execution of Application or command name:" & sAppName & " requiring a reboot."
										
								 		Exit sub
								 
								End if
														    
								'reset					    
								i = 0
								tmp_dic.RemoveAll()
								
					Else
						WriteLog "Execution of Application or command name:" & sAppName & " is already complete."
						'reset					    
						i = 0
						tmp_dic.RemoveAll()
							
					End If
	
							 	
				End If 
										   
				i = i + 1 
										
			Next 'end of For Each c In cnode  
			
			if x=Counter then
					
						bLastAppInstDone="true"
						WriteLog "Last task execution is complete."
						
						'remove shortcut from Startup and remove credentials from autologon registry
						CleanupStartItems
						
						
						
			end If
			
			
		Next 'end of For x = 1 To counter 
				  
	Else 'iRetval=Failure
	
	
	End If
	
	WriteLog "End Sub: Install_Apps"
	
End Sub

Function Run_Install_App (sAppName, sWrkgDir, sAppCmd)
	
	
	WriteLog "Begin Function: Run_Install_App"
	
	
	Message = "Please wait while executing task: "  & sAppName & "..."
	WriteLog "Start the execution of " & sAppName & "..."
	
	'run command and display a splash screen
		
	Result= RunWithSplash(sWrkgDir, sAppCmd, Message)
	
		
			If Err Then
					Result = Err.number
					sError = Err.Description
			        writeLog "Error executing Application or command " & sAppName & ": " & sError
			        iRetval= killProcess ("mshta.exe")	
					WScript.Sleep 1000
			        DisplaySplash sSplashErrorHTA, "Error executing Application or command " & sAppName & ": " & sError	
			        oShell.AppActivate sSplashErrorHTA
			        iRetVal=Failure
			        Exit Function
			End If        
			
			If Result = 0 Then
				
		            WriteLog "Application or command " & sAppName & " executed successfully"
		            iRetval= killProcess ("mshta.exe")	
					WScript.Sleep 1000
			Else
		            WriteLog "Application or command " & sAppName & " returned an unexpected return code: " & Result
		            iRetval= killProcess ("mshta.exe")	
					WScript.Sleep 1000
		            DisplaySplash sSplashErrorHTA, "Error executing Application or command " & sAppName & ": " & sError
					oShell.AppActivate sSplashErrorHTA
					Run_Install_App = Failure
					iRetVal=Failure
					WriteLog "Exit Function: Run_Install_App"
					Exit Function
					
			End if
	WriteLog "End Function: Run_Install_App"

End Function 

Function RunWithSplash(sDir, sCmd, msg)

		Dim oExec
		Dim lastHeartbeat
		Dim lastStart
		Dim iHeartbeat
		Dim iMinutes

		' Initialize the last heartbeat time (start the timer) and interval

		lastHeartbeat = Now
		iHeartbeat = 5
	
		
		If objFSO.FolderExists(sDir) Then
		
				'display splash
				DisplaySplash sSplashHTA, msg
				
				'set working directory
				oShell.CurrentDirectory = sDir
				
				
				' Start the command
				
				Writelog "About to run command: " & sAppCmd
				lastStart = Now
				Set oExec = oShell.Exec(sCmd)
				Do While oExec.Status = 0
					
					'activate the HTA page
					oShell.AppActivate sSplashHTA
					
					' Sleep
					WScript.Sleep 500
					
					' See if it is time for a heartbeat
					If iHeartbeat > 0 and DateDiff("n", lastHeartbeat, Now) > iHeartbeat then
						iMinutes = DateDiff("n", lastStart, Now)
										
						writeLog "Heartbeat: command has been running for " & iMinutes & " minutes (process ID " & oExec.ProcessID & ")"
						
						If iMinutes > 60 Then
						
							writeLog "ERROR: command has been running for more than 60 minutes... So assuming that there is a problem, the execution is aborted." 
							
							RunWithSplash = Failure
							Exit Do
							
						End If
						
						lastHeartbeat = Now
						
					End if
		
				Loop
				
				If RunWithSplash = Failure Then
					
					'close current splash
				
					iRetval= killProcess ("mshta.exe")	
					WScript.Sleep 1000
					oShell.CurrentDirectory = sScriptDir
					iRetVal=Failure
					Exit Function
					
				End If
				
		Else
			
			RunWithSplash = Failure
			writeLog "ERROR: File/Folder not found:" & sDir
			iRetVal=Failure
			Exit Function
			
		End If
		
		' Return the exit code to the caller

		writelog "Return code from command = " & oExec.ExitCode
		RunWithSplash = oExec.ExitCode
		
End Function

Sub DisplaySplash (sSplash,msg)
Dim sCmd
	If objFSO.FileExists(sSplash) Then
				
				sCmd="cmd /c mshta.exe " & Chr(34) & sSplash & Chr(34) & " /" & msg
				iRetVal = oShell.Run(sCmd, 0, False)
				
				If iRetVal <> Success then
					WriteLog "ERROR - Execution command:" & sCmd & " returned a non-zero return code, rc = " & iRetVal
					WriteLog "ERROR: " & sSplash & " Form cannot be displayed."
					
				End If
				
			Else
			
				WriteLog "Warning: Unable to find file: " & sSplash
				WriteLog "Skipping the display of Splash Form"
				iRetVal=Failure
				
	End If

End Sub


Function killProcess (strProcessKill)

Dim objWMI,colProcess, objProcess
Dim strComputer 
strComputer = "."


Set objWMI = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _ 
& strComputer & "\root\cimv2") 

Set colProcess = objWMI.ExecQuery _
("Select * from Win32_Process Where Name = '" & strProcessKill & "'" )
For Each objProcess in colProcess
	
	objProcess.Terminate()
	' Sleep
	WScript.Sleep 1000
	
Next 

End Function

Function RunWithHeartbeat(sCmd)

		Dim oExec
		Dim lastHeartbeat
		Dim lastStart
		Dim iHeartbeat
		Dim iMinutes

		' Initialize the last heartbeat time (start the timer) and interval

		lastHeartbeat = Now
		iHeartbeat = 5

		' Start the command

		Writelog "About to run command: " & sCmd
		lastStart = Now
		Set oExec = oShell.Exec(sCmd)
		Do While oExec.Status = 0

			' Sleep
			WScript.Sleep 500
			
			' See if it is time for a heartbeat
			If iHeartbeat > 0 and DateDiff("n", lastHeartbeat, Now) > iHeartbeat then
				iMinutes = DateDiff("n", lastStart, Now)
								
				writeLog "Heartbeat: command has been running for " & iMinutes & " minutes (process ID " & oExec.ProcessID & ")"
				
				If iMinutes > 60 Then
				
					writeLog "ERROR: command has been running for more than 60 minutes... So assuming that there is a problem, the execution is aborted." 
					
					RunWithHeartbeat = Failure
					Exit Do
					
				End If
				
				lastHeartbeat = Now
			End if

		Loop

		If RunWithHeartbeat = Failure Then
			Exit Function
		End If
		
		' Return the exit code to the caller

		writelog "Return code from command = " & oExec.ExitCode
		RunWithHeartbeat = oExec.ExitCode

End Function

Function WriteLog(sLogMsg)

		Dim sTime, sDate, sTempMsg, oLog, oConsole

		'On error resume next		
		' Suppress messages containing password
		If not strDebug then
			If Instr(1, sLogMsg, "password", 1) > 0 then
				sLogMsg = "<Message containing password has been suppressed>"
			End if
		End if

		' Populate the variables to log
			sTempMsg = "[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  sLogMsg
			
		' If debug, echo the message
		If strDebug then
			Set oConsole = objFSO.GetStandardStream(1) 
			oConsole.WriteLine sLogMsg
		End if

		' Create the log entry
		Err.Clear
		Set oLog = objFSO.OpenTextFile(LogFile, ForAppending, True)
		
		If Err then
			Err.Clear
			Exit Function
		End if
		oLog.WriteLine sTempMsg
		oLog.Close
		Err.Clear

End Function


Function SectionContents(file, section)
		Dim oContents
		Dim line, equalpos, leftstring, ini

		Set oContents = CreateObject("Scripting.Dictionary")
		file = Trim(file)
		section = Trim(section)

		'On Error Resume Next
		Set ini = objFSO.OpenTextFile( file, 1, False)
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0

		Do While ini.AtEndOfStream = False
			line = ini.ReadLine
			line = Trim(line)
			If LCase(line) = "[" & LCase(section) & "]" Then
				line = ini.ReadLine
				line = Trim(line)
				Do While Left( line, 1) <> "["
					'If InStr( 1, line, item & "=", 1) = 1 Then
					equalpos = InStr(1, line, "=", 1 )
					If equalpos > 0 Then
						leftstring = Left(line, equalpos - 1 )
						leftstring = Trim(leftstring)
						oContents(leftstring) = Trim(Mid(line, equalpos + 1 ))
					End If

					If ini.AtEndOfStream Then Exit Do
					line = ini.ReadLine
					line = Trim(line)
				Loop
				Exit Do
			End If
		Loop
		ini.Close
		Set SectionContents = oContents

	End Function
	
Function ReadIni(file, section, item)

		Dim line, equalpos, leftstring, ini

		ReadIni = ""
		file = Trim(file)
		item = Trim(item)

		'On error resume next
		Set ini = objFSO.OpenTextFile( file, 1, False)
		If Err then
			Err.Clear
			Exit Function
		End if
		On Error Goto 0

		Do While (not ini.AtEndOfStream)
			line = ini.ReadLine
			line = Trim(line)
			If LCase(line) = "[" & LCase(section) & "]" and (not ini.AtEndOfStream) Then
				line = ini.ReadLine
				line = Trim(line)
				Do While Left( line, 1) <> "["
					'If InStr( 1, line, item & "=", 1) = 1 Then
					equalpos = InStr(1, line, "=", 1 )
					If equalpos > 0 Then
						leftstring = Left(line, equalpos - 1 )
						leftstring = Trim(leftstring)
						If LCase(leftstring) = LCase(item) Then
							ReadIni = Mid( line, equalpos + 1 )
							ReadIni = Trim(ReadIni)
							Exit Do
						End If
					End If

					If ini.AtEndOfStream Then Exit Do
					line = ini.ReadLine
					line = Trim(line)
				Loop
				Exit Do
			End If
		Loop
		ini.Close

End Function

' WRITEINI ( file, section, item, myvalue )
' file = path and name of ini file
' section = [Section] must not be in brackets
' item = the variable to write;
' myvalue = the myvalue to assign to the item.
'
Sub WriteIni( file, section, item, myvalue )
Dim in_section, section_exists, item_exists, wrote
Dim itemtrimmed, temp_ini, read_ini, write_ini
Dim line, linetrimmed, equalpos, TristateFalse, leftstring 
	
	in_section = False
	section_exists = False
	item_exists = ( ReadIni( file, section, item ) <> "" )

	wrote = False
	file = Trim(file)
	itemtrimmed = Trim(item)
	myvalue = Trim(myvalue)

	temp_ini = sScriptDir & objFSO.GetTempName

	Set read_ini = objFSO.OpenTextFile( file, 1, True, TristateFalse )
	Set write_ini = objFSO.CreateTextFile( temp_ini, False)

	While read_ini.AtEndOfStream = False
		line = read_ini.ReadLine
		linetrimmed = Trim(line)
		If wrote = False Then
			If LCase(line) = "[" & LCase(section) & "]" Then
				section_exists = True
				in_section = True
			ElseIf InStr( line, "[" ) = 1 Then
				in_section = False
			End If
		End If

		If in_section Then
			If item_exists = False Then
				write_ini.WriteLine line
				write_ini.WriteLine item & "=" & myvalue
				wrote = True
				in_section = False
			Else
				equalpos = InStr(1, line, "=", 1 )
				If equalpos > 0 Then
					leftstring = Left(line, equalpos - 1 )
					leftstring = Trim(leftstring)
					If LCase(leftstring) = LCase(item) Then
						write_ini.WriteLine itemtrimmed & "=" & myvalue
						wrote = True
						in_section = False
					End If
				End If
				If Not wrote Then
					write_ini.WriteLine line
				End If
			End If
		Else
				equalpos = InStr(1, line, "=", 1 )
				If equalpos > 0 Then
					leftstring = Left(line, equalpos - 1 )
					leftstring = Trim(leftstring)
					rightsring=Right (line, Len(line) - (Len(leftstring)+1) )
					If LCase(leftstring) = LCase(item) then
						If Len(rightsring) =0  Then
							'do nothing
						End if
					Else
						write_ini.WriteLine line
					End If
				Else
					write_ini.WriteLine line
				End If
			
		End If
	Wend

	If section_exists = False Then ' section doesn't exist
		write_ini.WriteLine
		write_ini.WriteLine "[" & section & "]"
		write_ini.WriteLine itemtrimmed & "=" & myvalue

	End If

	read_ini.Close
	write_ini.Close
	If objFSO.FileExists(file) then
		objFSO.DeleteFile file, True
	End if
	objFSO.CopyFile temp_ini, file, true
	objFSO.DeleteFile temp_ini, True

End Sub
	
'//---------------------------------------------------------------------------
'//  Function:	ConvertBooleanToString()
'//  Purpose:	Perform a Cstr operation manually to prevent localization 
'//             from converting True/False to non-english values.
'//---------------------------------------------------------------------------

Function ConvertBooleanToString(bValue)

		Dim iRetVal 

		iRetVal = Failure
			
		If bValue = -1 Then
			iRetVal = "True"
		ElseIf bValue = 0 Then
			iRetVal = "False" 
		End If

		ConvertBooleanToString = iRetVal

End Function

Sub UpdateStartupLink_DISABLE_UA

		Dim oLink

		' Set up to automatically run me, using the appropriate method

		If objFSO.FileExists(oEnv("SystemRoot") & "\Explorer.exe") then

			' If shortcut for OSDCUSTOMIZER.VBS  doesn't exist then create a new shortcut.

			If objFSO.FileExists(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk") then

			    ' Not Server Core, remove any previous link and create a new shortcut
			   
			   writelog "Removing previous OSDCustomizer shortcut from startup folder"

	    		objFSO.DeleteFile oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk", True

	    		WriteLog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" removed."
			   
			    writelog "Creating startup folder item to run Application or commands once the shell is loaded."

			    Set oLink = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk")
			    oLink.TargetPath = "wscript.exe"
			    
			    oLink.Arguments = """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_DISABLEUA"
			    
			    oLink.Save

			    writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" updated."

			Else
			     writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" not found. creating a new shortcut."
			     
			     writelog "Creating startup folder item to run Application or commands once the shell is loaded."

			    Set oLink = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk")
			    oLink.TargetPath = "wscript.exe"
			    
			    oLink.Arguments = """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_DISABLEUA"
			    
			    oLink.Save

			    writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" created."

			     
			End If

		Else

			' Server core or "hidden shell", register a "Run" item

			writelog "Creating Run registry key to run the OSDCustomizer for the next reboot."

			On Error Resume Next
			oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\DELL_OSDCustomizer", """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_DISABLEUA" , "REG_SZ"
			Writelog "Wrote Run registry key"
			
			On Error Goto 0

			' Allow execution to continue (assuming new Run item won't actually be run yet)

		End if

			
End Sub

Sub UpdateStartupLink

		Dim oLink

		' Set up to automatically run me, using the appropriate method

		If objFSO.FileExists(oEnv("SystemRoot") & "\Explorer.exe") then

			' If shortcut for OSDCUSTOMIZER.VBS  doesn't exist then create a new shortcut.

			If objFSO.FileExists(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk") then

			    ' Not Server Core, remove any previous link and create a new shortcut
			   
			   writelog "Removing previous OSDCustomizer shortcut from startup folder"

	    		objFSO.DeleteFile oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk", True

	    		WriteLog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" removed."
			   
			    writelog "Creating startup folder item to run Application or commands once the shell is loaded."

			    Set oLink = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk")
			    oLink.TargetPath = "wscript.exe"
			    
			    oLink.Arguments = """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_REBOOT"
			    
			    oLink.Save

			    writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" updated."

			Else
			     writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" not found. creating a new shortcut."
			     
			     writelog "Creating startup folder item to run Application or commands once the shell is loaded."

			    Set oLink = oShell.CreateShortcut(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk")
			    oLink.TargetPath = "wscript.exe"
			    
			    oLink.Arguments = """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_REBOOT"
			    
			    oLink.Save

			    writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" created."

			     
			End If

		Else

			' Server core or "hidden shell", register a "Run" item

			writelog "Creating Run registry key to run the OSDCustomizer for the next reboot."

			On Error Resume Next
			oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\DELL_OSDCustomizer", """" & sScriptDir & "OSD_Applications.VBS" & """" & " /AFTER_REBOOT" , "REG_SZ"
			Writelog "Wrote Run registry key"
			
			On Error Goto 0

			' Allow execution to continue (assuming new Run item won't actually be run yet)

		End if

			
End Sub


Sub RemoveOSDCustomizerFromStartRun

	' If shortcut for OSDCustomizer.vbs exist then remove the shortcut.

	If objFSO.FileExists(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk") then

	    writelog "Removing OSDCustomizer shortcut from startup folder"

	    objFSO.DeleteFile oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk", True

	    writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" removed."

	Else
	     writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" does not exist."
	End If

	
	sCurrentValuePath ="HKLM\Software\Microsoft\Windows\CurrentVersion\Run\DELL_OSDCustomizer"
	
	If customRegRead(sCurrentValuePath) Then
		writelog "Removing OSDCustomizer from Run registry key if exists."
	    WScript.Echo "Going to delete "& sCurrentValuePath
	    DeleteRegKey sCurrentValuePath
	End If

	On Error Goto 0

End Sub

Function customRegRead(sRegValue)
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    On Error Resume Next
    Err.Clear
    sRegReturn = oShell.RegRead(sRegValue)
    If Err.Number<>0 Then
        customRegRead = False
    Else
        customRegRead = sRegReturn
    End If  
End Function

Sub CleanupStartItems

		' Clean up the run registry entry (if it exists)
		writelog "CleanStartItems started"
		On Error Resume Next
		oShell.RegDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\OSDCustomizer"
		On Error Goto 0


		' Clean up the shortcut (if it exists)

		If objFSO.FileExists(oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk") then
		
			objFSO.DeleteFile oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"
			 writelog "Shortcut """ & oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" removed."
		End if

		On Error Resume Next
		WriteLog "Removing auologon registry keys if exists"
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon", "0", "REG_SZ"
		oShell.RegDelete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword"
		oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoLogonCount", &H00000000, "REG_DWORD"
		On Error Goto 0

		writelog "CleanStartItems Complete"

	End Sub

Function IsCurrentUserAdmin(sadmin)

		Dim colUserAccounts, oAccount
		On Error Resume Next
		
		IsCurrentUserAdmin = Failure
		
		
		'Determine Local Administrator Account
		Set colUserAccounts = objWMIService.ExecQuery("Select * From Win32_UserAccount where LocalAccount = TRUE")
		For each oAccount in colUserAccounts
			If Left(oAccount.SID, 6) = "S-1-5-" and Right(oAccount.SID, 4) = "-500" Then
				If UCase(oAccount.Name)= UCase(sadmin) Then
					Writelog sadmin & " user account is a local administrator account."
					
					IsCurrentUserAdmin = Success
					Exit For
				End If 
				
			End iF
		Next
	
End Function

