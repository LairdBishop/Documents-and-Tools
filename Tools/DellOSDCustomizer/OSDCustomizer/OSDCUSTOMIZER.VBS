' //******************************************************************************************************************
' // Author: 	Amar Maouche - Dell IMS
' // Version:		30.06.2016
' // 
' // Updated with the possibility to run the tool from Local drive, USB media and MDT deployment Share.
' // Purpose:   
' //			This script can run from a SCCM /MDT task sequence or as a command line for Dell Image Wizard Images.
' // Usage:     Cscript OSDCustomizer.vbs [/CreateProfile:Filename] or [/ApplyProfile:<profilename.ini> or AssetTag] or [/MDTForms:DeployWiz_LanguageUI.xml,DeployWiz_ComputerName.xml]
' //
' // The arguments are as below and only one can be used at the same time:
' //		The [/CreateProfile:<Profilename.ini>] can be used to prepare the <Profilename.ini>. This file will be applied later if running  a second time with /ApplyProfile argument. 
' //		The [/ApplyProfile:<Profilename.ini>] can be used to apply the specified <Profilename.ini>. Using this argument, it will run OSDCustomizer silently to apply specific settings.
' //		The [/ApplyProfile:AssetTag] can be used to get the name of the OSDProfile.INI to apply. This is usefull for CFI scenario to run OSDCustomizer silently to apply specific settings.
' //		The [/MDTForms:DeployWiz_LanguageUI.xml,DeployWiz_ComputerName.xml,NICSettings_Definition_ENU.xml] can be used only with an MDT Image, to manage the regional settings and /or Computername Join domain and/or NIC settings.

' // Notes: 	 *********
' //	
' //	Below are the different options to run the script:
' //
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs   										---> without argument, this will display the HTA page
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/CreateProfile:<profilename.ini>]   		----> this will displaye HTA to prepare the <profilename.ini> to be used by running a 2nd time the script with /ApplyProfile
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/ApplyProfile:<profilename.ini>]    		----> this will apply silently the settings in <profilename.ini> created previusly with option /CreateProfile
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/ApplyProfile:AssetTag]					----> this will detect the profile name from the content of Asset Tag and run silently the settings in <profilename.ini>
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/MDTForms:DeployWiz_LanguageUI.xml] 		----> Used only with MDT image and will display the MDT form for regional settings and apply the selected settings
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/MDTForms:DeployWiz_ComputerName.xml]	----> Used only with MDT image and will display the MDT Form for Computername, Join domain and OUs and apply the selected settings
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/MDTForms:NICSettings_Definition_ENU.xml]	----> Used only with MDT image and will display the MDT Form for NIC settings
' //		-  Cscript.exe /Wscript.exe OSDCustomizer.vbs [/MDTForms:DeployWiz_LanguageUI.xml,DeployWiz_ComputerName.xml,NICSettings_Definition_ENU.xml]	----> Used only with MDT image and will display the 3 MDT Forms as above.
' //
' //  This script is calling OSD_Jobs.wsf with Job="MDT" for an MDT image or run the OSD_core.vbs script for NON MDT Images.	
' //*******************************************************************************************************************

Const s_ForReading = 1
Const s_ForWriting = 2
Const s_ForAppending = 8

Dim s_objFSO, s_objShell, s_ScriptDir, s_Root_Drv, s_OSDCustomizer_Jobs_File, s_OSD_Core_File, s_MDT_SCRIPTROOT
Dim s_LogFile, s_sLogPath, s_Args, objArgs, s_cmd, s_oTaskSequence, iSMDT, s_WinDir
Dim sPhase, isRunTS

isRunTS="NO"
on error resume next

'declare objects
Set s_objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set s_objShell = WScript.CreateObject("WScript.Shell")

' set the script directory
s_ScriptDir = WScript.ScriptFullName

s_ScriptName= s_objFSO.GetFileName(s_ScriptDir)
s_ScriptDir = Left(s_ScriptDir, InStrRev(s_ScriptDir, "\"))

'set root drive letter
s_Root_Drv = s_objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%")
s_WinDir=s_objShell.ExpandEnvironmentStrings ("%WINDIR%")

'set OSDCustomizer_Jobs filename
s_OSDCustomizer_Jobs_File="OSD_Jobs.wsf"
s_OSD_Core_File="OSD_Core.vbs"
s_SetTop="SetTop.exe"

'set path for Default factory MDT deploy\scripts folder
s_MDT_SCRIPTROOT=s_Root_Drv & "\Deploy\Scripts"

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

'check if an MDT image

If Not (s_objFSO.FileExists(s_MDT_SCRIPTROOT & "\media.tag") And s_objFSO.FileExists(s_MDT_SCRIPTROOT & "\LiteTouch.wsf")) Then
	'WScript.Echo "MDT SCRIPTS folder not found on local drive. checking parent folder of current script directory"
	iSMDT="NO"
	s_MDT_SCRIPTROOT=GetTheParent(s_ScriptDir)	
		
	If s_objFSO.FileExists(s_MDT_SCRIPTROOT & "\LiteTouch.wsf") Then
		
		'WScript.Echo "MDT SCRIPTROOT=" & s_MDT_SCRIPTROOT
		iSMDT="YES"
		
	Else
	
		'WScript.Echo "MDT SCRIPTS folder not found. Checking one level up from the current script directory"
		s_MDT_SCRIPTROOT= GetTheParent(s_MDT_SCRIPTROOT)
		
		If s_objFSO.FileExists(s_MDT_SCRIPTROOT & "\LiteTouch.wsf") Then
			'WScript.Echo "MDT SCRIPTROOT=" & s_MDT_SCRIPTROOT
			iSMDT="YES"
		Else
			'WScript.Echo "No MDT SCRIPTS folder found."
			iSMDT="NO"
		End If
		
	End If
Else
	'WScript.Echo "MDT SCRIPTS folder found on local drive. MDT SCRIPTROOT=" & s_MDT_SCRIPTROOT
	iSMDT="YES"
End If


'check if running from UNC path and copy all OSDcustomizer c:\windows\temp\OSDCustomizer folder
If Left(s_ScriptDir, 2)="\\" Then
	 s_cmd="cmd /c XCOPY /CHERKY " & Chr(34) & s_ScriptDir & "*" & Chr(34) & " " & s_WinDir & "\Temp\OSDCustomizer\" 
	 s_objShell.Run s_cmd,0,True
	 s_ScriptDir=s_WinDir & "\Temp\OSDCustomizer\"
End If      


If Script_Started_from_TS() = True Then
  isRunTS="YES"
  
  If iSMDT="YES" then
      
      'check if MDT running from UNC path and copy all OSDcustomizer c:\windows\temp\OSDCustomizer folder
      If Left(s_MDT_SCRIPTROOT, 2)="\\" Then

	      	'create a MDT.flg file
			Set s_oLog= s_objFSO.CreateTextFile(s_ScriptDir & "\MDT.FLG", True)
			s_oLog.WriteLine(s_MDT_SCRIPTROOT)
			s_oLog.Close
      
      End If      
      
  Else
  	  iSMDT="NO"	
     
  End If
  
Else

  iSMDT="NO"
  isRunTS="NO"
End If 

' set Log file
If Instr(1, UCase(s_Args), "/POST", 1) > 0 Then
	sPhase="POST"
	Post_OSD_LogFile="POST_OSDCustomizer.log"
End if

s_LogFile ="OSDCustomizer.log"
strSafeDate= DatePart("yyyy",Date) & Right ("0" & DatePart("m",Date),2) & Right("0" &DatePart("d", Date),2)
strSafeTime= Right ("0" & Hour(Now),2) & Right ("0" & Minute(Now),2) & Right ("0" & Second(Now),2)
strDateTime=strSafeDate &"-" & strSafeTime
s_LogFileNameBackup ="OSDCustomizer-" & strDateTime & ".log" 

'set log path to point to _SMSTSLogPath if running script from a TS
	If isRunTS="YES" Then
		
		If iSMDT="YES" then
	    		s_sLogPath = s_Root_Drv & "\MININT\SMSOSD\OSDLOGS" & "\"	
		Else
	    		s_sLogPath = s_oTaskSequence("_SMSTSLogPath") & "\"	
		End If
		
		If Not s_objFSO.FolderExists(s_sLogPath) Then
		 	' Logpath not exist then use the Windows\temp folder
			If Not s_objFSO.FolderExists(s_WinDir & "\Temp") Then
				s_objFSO.CreateFolder(s_WinDir & "\Temp")
			End If
	    		s_sLogPath=s_WinDir & "\Temp\"
	    End If
	    
	Else
			s_sLogPath=s_ScriptDir		
	End If

	S_LogFile = s_sLogPath & s_LogFile
	
	If sPhase="POST" Then
    	Post_OSD_LogFile =s_sLogPath & Post_OSD_LogFile
    End If


'backup old log if exist
If s_objFSO.FileExists (s_LogFile) Then

 	If  Not (sPhase="POST") Then
 		'backup old log if exist
		s_objFSO.CopyFile s_LogFile, s_sLogPath & s_LogFileNameBackup
		s_objFSO.DeleteFile s_LogFile, True	
		
		Set s_oLog= s_objFSO.CreateTextFile(s_LogFile, True)
		
		's_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  " OSDCustomizer execution script started.")	
		
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "Script " & s_ScriptName & " execution is started.")
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")

		s_oLog.Close
		
	Else 
	
		WriteLog "====================================================================================="
		WriteLog "Script " & s_ScriptName & " PHASE2 execution is started."
		WriteLog "====================================================================================="

		
	'	WriteLog "[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  " POST OSDCUSTOMIZER Phase started."
	
	End If

Else
	
	'create a new log file
		Set s_oLog= s_objFSO.CreateTextFile(s_LogFile, True)
		's_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  " OSDCustomizer execution script started.")	
		
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "Script " & s_ScriptName & " execution is started.")
		s_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")

		s_oLog.Close

End If


if s_Args <>"" then
  WriteLog "Running script with following arguments:" & s_Args
end If

s_objShell.CurrentDirectory = s_ScriptDir

'execute the right script depending on image type
If isRunTS="YES" and iSMDT="YES" then

	'MDT scripts folder found. this is an MDT image
	'run the OSD_Jobs.WSF script with JobID=MDT
	
	s_cmd="cscript.exe" & " " & Chr(34) & s_ScriptDir & s_OSDCustomizer_Jobs_File & Chr(34) & " " & "//job:" & chr(34) & "MDT" & Chr(34) & " " & s_Args

	 WriteLog "Running OSD_Jobs.WSF script with Job id=MDT: " & s_cmd
	s_objShell.Run s_cmd,0,True
	
Else

	'not MDT image
	'Run SetTop.exe to make the OSDCustomizer HTA top-most window
	s_cmd=s_ScriptDir & s_SetTop
	 WriteLog "Running SetTop.exe to make the OSDCustomizer HTA top-most window. Executed command: " & s_cmd
	s_objShell.Run s_cmd,0,False	
	
	'run the OSD_core.vbs script
	s_cmd="cscript.exe" & " " & Chr(34) & s_ScriptDir & s_OSD_Core_File & Chr(34) & " " & s_Args
	 WriteLog "Running OSD_Core.vbs script with following command: " & s_cmd
	s_objShell.Run s_cmd,0,True	
	
End If

'sMsg="Script execution is completed."

If  sPhase="POST" Then
	sMsg="Script " & s_ScriptName & " PHASE 2 execution is completed."
	WriteLog "====================================================================================="
	WriteLog sMsg
	WriteLog "====================================================================================="

End If

					
WScript.Quit

'***********************************************************************
' --------------------------------------------
' if scripting object "Microsoft.SMS.TSEnvironment" can be created
' then Returnvalue = true and Object is created
' else Returnvalue = false
' --------------------------------------------
Function Script_Started_from_TS
	Script_Started_from_TS  = False
    Err.Clear
	On Error Resume Next
    Set s_oTaskSequence = CreateObject("Microsoft.SMS.TSEnvironment")
	If Err.Number  <> 0 Then
		
		On Error Goto 0
		Script_Started_from_TS = False		
		Exit Function
	End If
	On Error Goto 0
	Script_Started_from_TS  = True
	
End Function

Function WriteLog(sLogMsg)

		Dim sTime, sDate, sTempMsg, s_oLog, oConsole

		On Error Resume Next		
		' Suppress messages containing password
		If not s_bDebug then
			If Instr(1, sLogMsg, "password", 1) > 0 then
				sLogMsg = "<Message containing password has been suppressed>"
			End if
		End if

		' Populate the variables to log
			sTempMsg = "[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  sLogMsg
			
		' If debug, echo the message
		If s_bDebug then
			Set oConsole = s_objFSO.GetStandardStream(1) 
			oConsole.WriteLine sLogMsg
		End if

		' Create the log entry
		Err.Clear
		
		Set s_oLog = s_objFSO.OpenTextFile(s_LogFile, s_ForAppending, True)
		
		If Err then
			Err.Clear
			Exit Function
		End If
		
		s_oLog.WriteLine sTempMsg
		
		s_oLog.Close
		Err.Clear

End Function

Function GetTheParent(DriveSpec)
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
   GetTheParent = objFSO.GetParentFolderName(Drivespec)
End Function
		       	