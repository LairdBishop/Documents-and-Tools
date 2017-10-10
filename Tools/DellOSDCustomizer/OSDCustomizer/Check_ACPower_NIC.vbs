' // ***************************************************************************
' // File: Check_ACPower_NIC.vbs
' // Version:		09.10.2016
' // Purpose:  	
' //		-	Detect if a system is running from a battery and display a warning.
' //		- 	Check if system is connected a network via Ethernet cable otherwise display a warning
' //		- 	if runing from a task sequence, it set the required OSD variables to initiate a reboot of TS if required
' //
' // Usage:     cscript.exe [//nologo] Check_ACPower_NIC.vbs
' // 
' // ***************************************************************************
Option Explicit

	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFSO, WshShell, objWMIService, ScriptDir, ScriptName
	Dim LogFile, sLogPath, oTaskSequence, oTSProgressUI, WinDir, strComputer
	Dim objNetworkAdapters, objAdapter, colItems, objItem, aStatus, iEthernet, iWireless
	Dim i, strIP, sMsg1, sMsg2, sWarning
	Dim vbool1, vbool2, vFromTS
	
	'on error resume next

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set WshShell = WScript.CreateObject("WScript.Shell")

	ScriptDir = WScript.ScriptFullName
	ScriptName= objFSO.GetFileName(ScriptDir)
	ScriptDir = Left(ScriptDir, InStrRev(ScriptDir, "\"))

	WinDir=WshShell.ExpandEnvironmentStrings ("%WINDIR%")
	
	strComputer="."
	

	sMsg1="AC power not detected. Please plug in to power."
	sMsg2="Wired connection not found. Please plug in to a network cable."


' set Log file
	LogFile =Left(ScriptName, Len(ScriptName)-4) & ".log"

'set log path to point to _SMSTSLogPath if running script from a TS
	If Script_Started_from_TS() = True Then
	
		vFromTS=True
	    sLogPath = oTaskSequence("_SMSTSLogPath") & "\"	
	    
	   'Hide TS progress bar
	    Set oTSProgressUI = CreateObject("Microsoft.SMS.TsProgressUI") 
		oTSProgressUI.CloseProgressDialog
		Set oTSProgressUI = Nothing	
		
	Else
	    sLogPath=WinDir & "\Temp\"
	End If
	 
	LogFile = sLogPath & LogFile
	WriteLog "script: " & ScriptName & " execution started."
	
	
	
'Display a warning if laptop is running from battery

	WriteLog "Checking if Laptop running from a battery and display a warning."
	
	vbool1=False

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")	
	Set colItems = objWMIService.ExecQuery( "Select * from Win32_Battery")
	
	if colItems.count = 1 Then
	
		WriteLog "This system is a laptop."
		
		For Each objItem in colItems
			
			If objItem.BatteryStatus = 1 Then
				
				sWarning=sWarning & sMsg1 & vbLf
				WriteLog "Warning: This system is running from a battery. Needs to be connected to an AC power adapter."
				
			Else
			
				WriteLog "This laptop is running from an AC Power adapter."
				vbool1=True
				
			End If
			
		Next
		
	Else
	' this is a desktop, nothing to do
		WriteLog "This system is a desktop."
		vbool1=True 
	End If
	
	Set colItems = Nothing	
			
	
' display a warning for network connection
 
	 aStatus = Split("Disconnected/Connecting/Connected/Disconnecting/" _
	 & "Hardware not present/Hardware disabled/Hardware malfunction/" _
	 & "Media disconnected/Authenticating/Authentication succeeded/" _
	 & "Authentication failed", "/")
	 
	 iEthernet = 0
	 iWireless = 9
	 vbool2= False
	 
	 WriteLog "Detecting Network Adpater Status..." 
	
	 Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter",,48)
	 
	 For Each objItem In colItems
	 
	 	If (objItem.AdapterTypeId = iEthernet) And (Not (IsNull(objItem.NetConnectionStatus))) Then 
	 	
	 		If Not((InStr(1,LCase(objItem.Name), "vmware",1)> 0) Or _
	 			(InStr(1,LCase(objItem.Name), "bluetooth",1)>0) Or _
	 				(InStr(1,LCase(objItem.Name), "wireless",1)>0)) Then 
	 		 
	 			WriteLog "Network Adpater: " &  objItem.Name & " : " & aStatus(objItem.NetConnectionStatus)
	 			
	 			
	 			If LCase (aStatus(objItem.NetConnectionStatus))=lcase("Connected") Then
	 				
	 				Set objNetworkAdapters = objWMIService.ExecQuery( _
	 					"select * from Win32_NetworkAdapterConfiguration where IPEnabled = 1")
	 					
					For Each objAdapter In objNetworkAdapters
						
						If InStr(1,LCase(objAdapter.Caption),LCase(objItem.Name),1) > 0 Then 
						
						 	If Not IsNull(objAdapter.IPAddress) Then
								For i = LBound(objAdapter.IPAddress) to UBound(objAdapter.IPAddress)
								 
								  If Not Instr(objAdapter.IPAddress(i), ":") > 0 Then
									  strIP = strIP & objAdapter.IPAddress(i) & " "
								  End If
								  
								Next
							 
							 	If Len(Trim(strIP)) > 0 Then
									vbool2= True
									'MsgBox "Found a connected Network Adpater: " &  objAdapter.Caption & " with IP address: " & strIP
									WriteLog "Found a connected Network Adpater: " &  objAdapter.Caption & " with IP address: " & strIP
								End If
								
							End If
							
							If vbool2=True Then Exit for
							
						End If
						
					Next
					
					If vbool2=False Then
						WriteLog "NO IP address found for the Network Adpater: " &  objAdapter.Caption	
					End If
					
				Else
					WriteLog "Network Adpater: " &  objItem.Name & " is not connected."
					vbool2= False
				End If
				
	 		End If
	 		
	 	End If
	 	
	 	If vbool2= True Then Exit For
	 	
	Next
 
	If vbool2= False Then sWarning=sWarning & sMsg2 & vbLf
	
	If vbool1= False or vbool2= False Then
	
			WriteLog "Display a warning for laptop on battery and/or network connection."
			
			If vFromTS=True Then
			
				sWarning= sWarning & vbLf
				sWarning= sWarning & "Click OK to continue or Cancel to stop and reboot the system." & vbLf 
				
				If MsgBox (sWarning, vbOKCancel + vbExclamation, "Warning") <> 1 Then
					
					WriteLog "User pressed the Cancel button. Process is stopped and a reboot will be initiated by Task Sequence."
					
					' Set properties to indicate a reboot is needed and this script should be re-executed
																
					oTaskSequence("SMSTSRebootRequested") = "true"
					oTaskSequence("SMSTSRetryRequested") = "true"
				
					WriteLog "Exiting to initiate a reboot with retry (to pick up where we left off)"
					
					WScript.Quit
				else
					WriteLog "User pressed the OK button."
					
				End If
				
			Else
				
				MsgBox sWarning ,vbExclamation, "Warning"
				
			End if 
	End If
	
'exit script

	WriteLog "script: " & ScriptName & " execution completed."
	WScript.Quit

'***********************************************************************

Function Script_Started_from_TS
	
    Err.Clear
	On Error Resume Next
    Set oTaskSequence = CreateObject("Microsoft.SMS.TSEnvironment")
	If Err.Number  <> 0 Then
		
		On Error Goto 0
		Script_Started_from_TS = False		
		Exit Function
	End If
	On Error Goto 0
	Script_Started_from_TS  = True
	
End Function


Function WriteLog(sLogMsg)

	Dim sTime, sDate, sTempMsg, oLog, oConsole

	On Error Resume Next		
	' Suppress messages containing password
	If not bDebug then
		If Instr(1, sLogMsg, "password", 1) > 0 then
			sLogMsg = "<Message containing password has been suppressed>"
		End if
	End if

	' Populate the variables to log
		sTempMsg = "[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  sLogMsg
		
	' If debug, echo the message
	If bDebug then
		Set oConsole = objFSO.GetStandardStream(1) 
		oConsole.WriteLine sLogMsg
	End if

	' Create the log entry
	Err.Clear
	
	Set oLog = objFSO.OpenTextFile(LogFile, ForAppending, True)
	
	If Err then
		Err.Clear
		Exit Function
	End If
	
	oLog.WriteLine sTempMsg
	
	oLog.Close
	Err.Clear

End Function


















