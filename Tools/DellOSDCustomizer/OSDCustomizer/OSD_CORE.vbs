' //******************************************************************************************************************
' // Author: 	Amar Maouche - Dell IMS
' // Version:		21.03.2017
' //
' // Requirement: This script has to be called by main script OSDCustomizer.vbs 
' //			
' //*******************************************************************************************************************

' initialize variables and objects

	Const OSD_ForReading = 1
	Const OSD_ForWriting = 2
	Const OSD_ForAppending = 8
	const HKEY_LOCAL_MACHINE = &H80000002
	
	Const OSD_Success = 0
	Const OSD_Failure = 1
	
	Dim OSD_objFSO, OSD_oShell, OSD_objWMIService, OSD_RootDrv, OSD_LogFile, OSD_sScriptDir
	Dim OSD_sSettingsIniFile, OSD_dicIPAddresses, OSD_dicDefaultGateway
	Dim OSD_sOSDFrontEndHTAFile, OSD_sKeybHTAFile, OSD_sOSDProfileIniFile
	Dim OSD_sLogPath, OSD_oLog, s_WinDir
	
	Dim OSD_ScriptName, OSD_skipWizard, OSD_Arguments, OSD_NamedArgs, OSD_sArg, OSD_sCmd, OSD_strComputer, OSD_bDebug, OSD_bCreateINI, OSD_bApplyINI
	Dim OSD_bMerge, OSD_bDefaultGateway, OSD_sSite, OSD_iRetVal
	Dim OSD_sArgCreate, OSD_sArgApply, OSD_sArgMDTForms, MDT_sSCRIPTROOT, OSD_sSplitArgMDTForms
	
	Dim OSD_sSystemLocale, OSD_sUserLocale, OSD_sInputLocale, OSD_sUILanguage, OSD_sCountry, OSD_sTimeZone
	Dim OSD_sPrefix, OSD_sSuffix, OSD_sNewComputername, OSD_sCurrentHostname, OSD_sAssetTag, OSD_sServiceTag, OSD_sDesktop, OSD_Server, OSD_sLaptop, OSD_sConstruct
	Dim OSD_sDomain, OSD_sDomainOU, OSD_sDomainUser, OSD_sDomainUserPassword
	Dim OSD_sIPAddress, OSD_sGatewayAddress, OSD_sPrimDNS, OSD_sSecDNS, OSD_sRegion, OSD_sLocationName, OSD_RunCommand
	Dim OSD_sJoin, OSD_sWorkgroup, sJoinSystem
	
	Dim sArchitecture
	Dim sDomainXML_template, sNetworkInterfaceXML_template, sOOBEXml_template, sUnattendXml
		
	Dim OSD_sSkipKeyboardForm
			
	Dim OSD_oTaskSequence
	Dim OSD_oTSProgressUI
	Dim OSD_bRunningFromTS
	Dim OSD_varCFI	
	Dim OSD_sTSInWinPE
	Dim OSD_IsMDT
	OSD_IsMDT="NO"
			
	Set OSD_objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set OSD_oShell = WScript.CreateObject("WScript.Shell")
	Set OSD_objWMIService = Nothing	
	On Error Resume Next
	Set OSD_objWMIService = GetObject("winmgmts:")
	
	OSD_bRunningFromTS= False
	OSD_bDefaultGateway=False
	OSD_skipWizard =False
	
	OSD_bDebug= False
	OSD_bCreateINI= False
	OSD_bApplyINI= False
	OSD_strComputer= "."
	OSD_iRetVal= OSD_Success
	
	sRunPost="NO"
	
' set the script directory
	OSD_sScriptDir = empty
	OSD_sScriptDir = WScript.ScriptFullName
	OSD_sScriptName= OSD_objFSO.GetFileName(OSD_sScriptDir)
	
	OSD_sScriptDir = Left(OSD_sScriptDir, InStrRev(OSD_sScriptDir, "\"))
	OSD_oShell.CurrentDirectory = OSD_sScriptDir
	
' get the root system drive letter
	Set oEnv = OSD_oShell.Environment("PROCESS")
	OSD_RootDrv = OSD_oShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%")
	s_WinDir = OSD_oShell.ExpandEnvironmentStrings ("%WINDIR%")
	
' check if running from TS
	If Script_Started_from_TS() = True Then	
		OSD_bRunningFromTS = True
	Else
		On Error Resume Next
		' turn off setup flag in registry so we can query wmi
		OSD_oShell.RegWrite "HKLM\SYSTEM\Setup\SystemSetupInProgress", 0, "REG_DWORD"
	
		' open a command window if required for troubleshooting
	'	OSD_oShell.Run s_WinDir & "\System32\cmd.exe", 1, False
	End If

'check if MDT.FLG exists then MDT running from UNC path
If OSD_objFSO.FileExists(OSD_sScriptDir & "\MDT.FLG") Then
	' This is an MDT running from UNC path
	'get the UNC path for deplyoment share
	Set OSD_oFile = OSD_objFSO.OpenTextFile(OSD_sScriptDir & "\MDT.FLG", OSD_ForReading, True)
	While OSD_oFile.AtEndOfStream = False
		MDT_sSCRIPTROOT = OSD_oFile.ReadLine
		OSD_IsMDT="YES"
	Wend
	OSD_oFile.Close
	
Else

	'check if MDT Local drive or USB media /share drive letter 

	MDT_sSCRIPTROOT =OSD_RootDrv & "\Deploy\Scripts"

	If Not (OSD_objFSO.FileExists(MDT_sSCRIPTROOT & "\media.tag") And OSD_objFSO.FileExists(MDT_sSCRIPTROOT & "\LiteTouch.wsf")) Then
		'WScript.Echo "MDT SCRIPTS folder not found on local drive. checking parent folder of current script directory"
		OSD_IsMDT="NO"
		MDT_sSCRIPTROOT=GetTheParent(OSD_sScriptDir)	
			
		If OSD_objFSO.FileExists(MDT_sSCRIPTROOT & "\LiteTouch.wsf") Then
			'WScript.Echo "MDT SCRIPTROOT=" & MDT_sSCRIPTROOT
			OSD_IsMDT="YES"
		Else
			'WScript.Echo "MDT SCRIPTS folder not found. Checking one level up from the current script directory"
			MDT_sSCRIPTROOT= GetTheParent(MDT_sSCRIPTROOT)
			If OSD_objFSO.FileExists(MDT_sSCRIPTROOT & "\LiteTouch.wsf") Then
			'	WScript.Echo "MDT SCRIPTROOT=" & MDT_sSCRIPTROOT
				OSD_IsMDT="YES"
			Else
			'	WScript.Echo "No MDT SCRIPTS folder found."
				OSD_IsMDT="NO"
			End If
		End If
	Else
		'WScript.Echo "MDT SCRIPTS folder found on local drive. MDT SCRIPTROOT=" & MDT_sSCRIPTROOT
		OSD_IsMDT="YES"
	End If

End If

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

' set Log file
	If Instr(1, UCase(s_Args), "/POST", 1) > 0 Then
		sPhase="POST"
		Post_OSD_LogFile="POST_OSDCustomizer.log"
	End if

	OSD_LogFile ="OSDCustomizer.log"
	strSafeDate= DatePart("yyyy",Date) & Right ("0" & DatePart("m",Date),2) & Right("0" &DatePart("d", Date),2)
	strSafeTime= Right ("0" & Hour(Now),2) & Right ("0" & Minute(Now),2) & Right ("0" & Second(Now),2)
	strDateTime=strSafeDate &"-" & strSafeTime
	OSD_LogFileNameBackup ="OSDCustomizer-" & strDateTime & ".log" 

			    
'set log path
	If OSD_bRunningFromTS Then
	
	    	If OSD_IsMDT="YES" then
	    		OSD_sLogPath = OSD_RootDrv & "\MININT\SMSOSD\OSDLOGS" & "\"	
			Else
	    		OSD_sLogPath = OSD_oTaskSequence("_SMSTSLogPath") & "\"	
			End If
		
			If Not OSD_objFSO.FolderExists(OSD_sLogPath) Then
				If Not OSD_objFSO.FolderExists(s_WinDir & "\Temp") Then
					OSD_objFSO.CreateFolder(s_WinDir & "\Temp")
				End If		
				OSD_sLogPath=s_WinDir & "\Temp\"
			End If
			
	Else
			OSD_sLogPath=OSD_sScriptDir		
	End If

	OSD_LogFile = OSD_sLogPath & OSD_LogFile
	
	If sPhase="POST" Then
    	Post_OSD_LogFile =s_sLogPath & Post_OSD_LogFile
    End If

'backup old log if exist
	If OSD_objFSO.FileExists (OSD_LogFile) Then
	
 		If  sPhase="POST" Then
		
			WriteLog "====================================================================================="
			WriteLog "Script " & OSD_sScriptName & " PHASE2 execution is started."
			WriteLog "====================================================================================="
			
			
		Else
			WriteLog "====================================================================================="
			WriteLog "Script " & OSD_sScriptName & " execution is started."
			WriteLog "====================================================================================="
			
		End If

	Else

		'create a new log file
		Set OSD_oLog= OSD_objFSO.CreateTextFile(OSD_LogFile, True)
		OSD_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")
		OSD_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "Script " & OSD_sScriptName & " execution is started.")
		OSD_oLog.WriteLine("[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  "=====================================================================================")
			
		OSD_oLog.Close
	End If
 
 'get asset info of target computer
	WScript.Sleep 3000
	Call GetAssetInfo
  
' process the parameters for /CreateProfile /ApplyProfile /MDTForms /POST options	
	If Wscript.Arguments.Count = 0 Then	    
		
		WriteLog "No Arguments specified. OSDCustomizer will Proceed as Normal."
		
	ElseIf Wscript.Arguments.Count > 1 Then	 
		
		WriteLog "Too many Arguments. Only one of /CreateProfile /ApplyProfile or /MDTForms Arguments is allowed."
	    
	    WriteLog "====================================================================================="
		WriteLog "Script " & OSD_sScriptName & " execution is aborted."
		WriteLog "====================================================================================="
					
	    MsgBox "Too many Arguments. Only one of /CreateProfile /ApplyProfile or /MDTForms Arguments is allowed. Process aborted.",vbSystemModal, "Wrong Commandline"
		OSD_iRetVal=OSD_Failure
		
						
		
		WScript.Quit
		  
	Else
	
		Set OSD_NamedArgs = Wscript.Arguments.Named 
		OSD_sArgCreate= Trim(OSD_NamedArgs("CreateProfile"))
		OSD_sArgApply= Trim(OSD_NamedArgs("ApplyProfile"))
		OSD_sArgMDTForms=Trim(OSD_NamedArgs("MDTForms"))
		OSD_sArgPost=Trim(OSD_NamedArgs("POST"))
		
		If Len(OSD_sArgCreate) > 0 Then
		
					WriteLog "'/CreateProfile' parameter was specified."
		       		OSD_bCreateINI = True
		       		'get name of OSDProfileINI to create
		       		If InStr (1,OSD_sArgCreate, "\",1) > 0 Then
		        		OSD_sOSDProfileIniFile = OSD_sArgCreate
		        	Else
		        		OSD_sOSDProfileIniFile = OSD_sScriptDir & OSD_sArgCreate
		        	End If
		        	
		        	WriteLog "OSDProfileINI file name is:" & OSD_sOSDProfileIniFile
		        	
	    ElseIf Len(OSD_sArgApply) > 0 Then
	    
					WriteLog "'/ApplyProfile' parameter was specified."
		       		OSD_bApplyINI = True
		       		
		       		If InStr (1,UCase(OSD_sArgApply), "AssetTag",1) > 0 Then
		       			WriteLog "'AssetTag' parameter was specified.Gettings name of profile from AssetTag."
		       				       		
		       			'get content of Asset Tag from BIOS or From _RootDrv\AssetTag.txt
							
							If OSD_objFSO.FileExists(OSD_RootDrv & "\AssetTag.txt") Then
								sAssetTagFile=OSD_RootDrv & "\AssetTag.txt"
								Set objTextFile = OSD_objFSO.OpenTextFile(sAssetTagFile, OSD_ForReading)
								sAssetProfile = Trim(objtextFile.ReadLine)	
								
							ElseIf OSD_objFSO.FileExists(OSD_sScriptDir & "AssetTag.txt") Then
								sAssetTagFile=OSD_sScriptDir & "AssetTag.txt"
								Set objTextFile = OSD_objFSO.OpenTextFile(sAssetTagFile, OSD_ForReading)
								sAssetProfile = Trim(objtextFile.ReadLine)
										
							End If
				
							If Len(Trim(sAssetProfile))> 0 Then 
								WriteLog "*** Asset Tag detected from file " & sAssetTagFile & ". Asset Tag value = "& sAssetProfile
								
								'set name of OSDProfileINI to apply
					        	OSD_sOSDProfileIniFile = OSD_sScriptDir & sAssetProfile & ".ini"
					        	WriteLog "OSDProfileINI file name to apply is:" & OSD_sOSDProfileIniFile
					        	
							Else
							
							'check if Asset Tag is not empty
								
									sAssetProfile = ucase(trim(sAssetTag))
									If Len(sAssetProfile)> 0 Then
									
										WriteLog "*** Asset Tag detected from BIOS. Asset Tag value = "& sAssetProfile
										
										'set name of OSDProfileINI to apply
							        	OSD_sOSDProfileIniFile = OSD_sScriptDir & sAssetProfile & ".ini"
							        	WriteLog "OSDProfileINI file name to apply is:" & OSD_sOSDProfileIniFile
							        	
									End If
									
							End If
						
		       		Else
			       		
			       		'get name of OSDProfileINI to create
			       		If InStr (1,OSD_sArgApply, "\",1) > 0 Then
			        		OSD_sOSDProfileIniFile = OSD_sArgApply
			        	Else
			        		OSD_sOSDProfileIniFile = OSD_sScriptDir & OSD_sArgApply
			        	End If
			        	
			        	WriteLog "OSDProfileINI file name is:" & OSD_sOSDProfileIniFile
			        
			        End If
			        
		ElseIf Len(OSD_sArgMDTForms) > 0 Then
				
					WriteLog "'/MDTForms' parameter was specified."
		       		If OSD_objFSO.FolderExists(MDT_sSCRIPTROOT) Then
		       					       			
			       		' split the OSD_sArgMDTForms to get the name of forms to display: only 3 forms are supported here
			       		' DeployWiz_LanguageUI.xml, DeployWiz_ComputerName.xml, NICSettings_Definition_ENU.xml
			       		
			       		OSD_sSplitArgMDTForms=Split(OSD_sArgMDTForms , ":")
			       		Dim OSD_sRightArgMDTForms, OSD_arrMDTForms, OSD_sForm, OSD_i
			       		OSD_sRightArgMDTForms= OSD_sSplitArgMDTForms(UBound(OSD_sSplitArgMDTForms))
			       		OSD_arrMDTForms = Split(OSD_sRightArgMDTForms, ",")
			       		
			       			       		
			       		' process selected value and apply settinsg based on displayed forms
			       		If IsArray(OSD_arrMDTForms) Then
		            		For OSD_i = 0 to UBound(OSD_arrMDTForms)          
			    				
			    				OSD_sForm=OSD_arrMDTForms(OSD_i)
			    				If OSD_objFSO.FileExists(MDT_sSCRIPTROOT & "\" & OSD_sForm) Then
			    					'display the form
			    					OSD_IsMDT="YES"
			    					
			    					Select Case UCase(OSD_sForm)
			    					
			    						Case UCase("DeployWiz_LanguageUI.xml")
			    							'reset MDT variables
			    							oEnvironment.Item("SkipLocaleSelection")="NO"
			    							oEnvironment.Item("SkipTimeZoneName")="NO"
			    							
			    							'Display MDT regional settings
			    							OSD_sCmd= "MSHTA.exe " & MDT_sSCRIPTROOT & "\Wizard.hta /definition:" & OSD_sForm
			    							ExecuteCommand OSD_sCmd
			    							
			    							'get properties
											OSD_sSystemLocale=Trim(oEnvironment.Item("SystemLocale"))
											If Len(OSD_sSystemLocale) > 0 Then WriteLog "SystemLocale=" & OSD_sSystemLocale
											
											OSD_sUserLocale=Trim(oEnvironment.Item("UserLocale"))
											If Len(OSD_sUserLocale) > 0  Then WriteLog "UserLocale=" & OSD_sUserLocale
											
											OSD_sInputLocale=Trim(oEnvironment.Item("InputLocale"))
											If Len(OSD_sInputLocale) > 0 Then WriteLog "InputLocale=" & OSD_sInputLocale
											
											OSD_sUILanguage=Trim(oEnvironment.Item("UILanguage"))
											If Len(OSD_sUILanguage) > 0 Then WriteLog "UILanguage=" & OSD_sUILanguage
																					
											OSD_sTimeZone=Trim(oEnvironment.Item("TimeZoneName"))
											If Len(OSD_sTimeZone) > 0 Then WriteLog "TimeZone=" & OSD_sTimeZone
											
			    							'Run SetRegionalSettings function
			    							setRegionalOptions
			    							
			    						Case UCase("DeployWiz_ComputerName.xml")
			    							
			    							'reset MDT variables
			    							oEnvironment.Item("SkipComputerName")="NO"
			    							oEnvironment.Item("SkipDomainMembership")="NO"
			    								    							
			    							'Display MDT Computername, Domain and Organisation Unit form
			    							OSD_sCmd= "MSHTA.exe " & MDT_sSCRIPTROOT & "\Wizard.hta /definition:" & OSD_sForm
			    							ExecuteCommand OSD_sCmd
			    							
			    							'get selected properties
			    							
			    							If oEnvironment.Item("OSDCOMPUTERNAME") <> "" Then
			    								OSD_sCurrentHostname=CreateObject("Wscript.network").ComputerName
			    								WriteLog "Current Computername is: " & OSD_sCurrentHostname
			    								
			    								OSD_sNewComputername=oEnvironment.Item("OSDCOMPUTERNAME")
			    								WriteLog "New Computername is: " & OSD_sNewComputername
												
												' rename hostname with new computername
		                      					Rename_Hostname OSD_sCurrentHostname, OSD_sNewComputername
												
			    							Else
			    								WriteLog "New Computername is empty. New hostname is not set."
			    								
			    							End If
			    							
			    							'check for Join domain properties
			    							If oEnvironment.Item("JoinWorkgroup") <> "" then
												WriteLog "Not attempting to join a domain because JoinWorkgroup = " & oEnvironment.Item("JoinWorkgroup") & "."
												
											ElseIf oEnvironment.Item("JoinDomain") <> "" then
												WriteLog "Join domain is selected with below properties:"
												WriteLog "Domaine name:" & oEnvironment.Item("JoinDomain")
												If oEnvironment.Item("DomainAdminPassword")<> "" Then
													WriteLog "Domain Admin Password: ****************"
												End If
												If oEnvironment.Item("DomainAdmin")<> "" Then
													WriteLog "Domain Admin User:" & oEnvironment.Item("DomainAdmin")
												End If
												If oEnvironment.Item("MachineObjectOU")<> "" Then
													WriteLog "Organisation Unit or MachineObjectOU:" & oEnvironment.Item("MachineObjectOU")
												End If
												
											End If	
			    						
			    						Case UCase("NICSettings_Definition_ENU.xml")
			    							
			    							'Display MDT NIC Setings form
			    							OSD_sCmd= "MSHTA.exe " & MDT_sSCRIPTROOT & "\Wizard.hta /definition:" & OSD_sForm
			    							ExecuteCommand OSD_sCmd
			    						
			    						
			    						Case Else 
			    						
			    							WriteLog "Error: Wrong MDT Form name entered as argument. Only following MDT forms are supported: DeployWiz_LanguageUI.xml, DeployWiz_ComputerName.xml and NICSettings_Definition_ENU.xml." 
			    						
			    							WriteLog "====================================================================================="
				    						WriteLog "Script " & OSD_sScriptName & " execution is aborted."
				    						WriteLog "====================================================================================="
					
			    							
			    							MsgBox "Wrong MDT Form name entered as argument. Only following MDT forms are supported: DeployWiz_LanguageUI.xml, DeployWiz_ComputerName.xml and NICSettings_Definition_ENU.xml. Process aborted.",vbSystemModal, "Wrong Argument"
											OSD_iRetVal=OSD_Failure
											
						
											
											WScript.Quit 
			    					
			    							
			    							
			    					End Select
			    					
			    				Else
			    				
			    					Msgbox "ERROR - File missing:" & OSD_RootDrv & "\Deploy\Scripts\" & OSD_sForm ,vbSystemModal, "File not found"
									WriteLog "ERROR - File missing:" & OSD_RootDrv & "\Deploy\Scripts\" & OSD_sForm
											    			
			    				End If
			    			
			    			Next
			    			
			    		End If
			    		
			    		If OSD_IsMDT="YES" Then 
							'script completed
							Writelog "====================================================================================="
							WriteLog "Script " & OSD_sScriptName & " execution is completed."
							Writelog "====================================================================================="
						
							
							'exit
							WScript.Quit
						End If
						
			    	Else
			    	
			    		
						WriteLog "ERROR - Folder missing:" & MDT_sSCRIPTROOT
						WriteLog "====================================================================================="
						WriteLog "Script " & OSD_sScriptName & " execution is aborted."
				    	WriteLog "====================================================================================="
				    	
						Msgbox "ERROR - Folder missing:" & MDT_sSCRIPTROOT ,vbSystemModal, "Folder not found"
						OSD_iRetVal=OSD_Failure
						
						
						
						WScript.Quit 	
								
			    	End If    	

	    ElseIf Len(OSD_sArgPost) > 0 Then
	    	
	    			WriteLog "/POST:" & OSD_sArgPost & " parameter was specified."
	    			
	    			'Phase 2 of OSDCustomizer to run as POST Sysprep at Logon session
	    			
	    			'get the OSDProfile.ini from /POST argument						    			
					OSD_sOSDProfileIniFile= OSD_sArgPost
					
					'get OSD selected config from OSDProfile.ini file
					getOSDSelectionConfig
					
						    
					'get installed MUI Languages from registry
					WriteLog "Getting installed UI Languages before application installation..."
					If getUILanguage=False Then
						bSetRegionalSettings= "YES"
					End If
					
					' call OSD_Application.vbs									
					OSD_sCmd = Chr(34) &  OSD_sScriptDir & "OSD_Applications.vbs" & Chr(34) & " /POST:" & Chr(34) &  OSD_sOSDProfileIniFile & Chr(34)
					
        			WriteLog "About to run command: " & OSD_sCmd
        			ExecuteCommand OSD_sCmd			

	    			'check if regional settings need to be applied again if a language pack is installed
	    			'If bSetRegionalSettings= "YES" Then
	    				'WriteLog "Getting UI Languages after application installation..."
	    				'If getUILanguage =True Then 
	    					
	    					'set regional settings for the new UI Language
					   			setRegionalOptions
	    				'Else
	    			'		WriteLog "UI Language " & OSD_sUILanguage & " cannot be determined if installed. It will not set in regional settings."
	    					
	    			'	End If

	    			'End If
					
					
					'exit
					WScript.Quit
	    
	    Else
			    	WriteLog" Error: Too many Arguments. Only one of /CreateProfile /ApplyProfile /MDTForms Arguments is allowed."
				    WriteLog "====================================================================================="
				    WriteLog "Script " & OSD_sScriptName & " execution is aborted."
				    WriteLog "====================================================================================="
					
				    MsgBox "Too many Arguments. Only one of /CreateProfile /ApplyProfile or /MDTForms Arguments is allowed. Process aborted.",vbSystemModal, "Wrong Commandline"
					OSD_iRetVal=OSD_Failure
					
						
					
					WScript.Quit 			
		End If		
		
	End If
	
	On Error Goto 0
	Err.Clear

' set the OSDSettings.ini file path
	OSD_sSettingsIniFile = "OSDSettings.ini"
	If OSD_objFSO.FileExists(OSD_sScriptDir & "cfg\" & "OSDSettings.ini") Then
		OSD_sSettingsIniFile = OSD_sScriptDir & "cfg\" & "OSDSettings.ini"
	Else
		WriteLog "Error: File not found: " &  OSD_sSettingsIniFile 
		WriteLog "====================================================================================="
		WriteLog "Script " & OSD_sScriptName & " execution is aborted."
		WriteLog "====================================================================================="
		
		MsgBox "Error: File not found: " &  OSD_sSettingsIniFile & ". Process aborted.",vbSystemModal, "File not found"
		OSD_iRetVal=OSD_Failure
		
						
		
		WScript.Quit   	
	End If

' set the hta files path
	OSD_sOSDFrontEndHTAFile = "OSDFrontEnd.hta"
	If OSD_objFSO.FileExists(OSD_sScriptDir & "hta\" &"OSDFrontEnd.hta") Then
		OSD_sOSDFrontEndHTAFile = OSD_sScriptDir & "hta\" &"OSDFrontEnd.hta"
	Else
		WriteLog "Error: File not found: " &  OSD_sOSDFrontEndHTAFile 
		WriteLog "====================================================================================="
		WriteLog "Script " & OSD_sScriptName & " execution is aborted."
		WriteLog "====================================================================================="
		
		MsgBox "Error: File not found: " &  OSD_sOSDFrontEndHTAFile & ". Process aborted.",vbSystemModal, "File not found"
		OSD_iRetVal=OSD_Failure
		
						
		
		WScript.Quit   	
	End If
	
	OSD_sKeybHTAFile ="KeybLayout.hta"
	If OSD_objFSO.FileExists(OSD_sScriptDir & "hta\" &"KeybLayout.hta") Then
		OSD_sKeybHTAFile =OSD_sScriptDir & "hta\" &"KeybLayout.hta"
	End If

	OSD_sAppsHTAFile ="AppForm.hta"
	If OSD_objFSO.FileExists(OSD_sScriptDir & "hta\" &"AppForm.hta") Then
		OSD_sAppsHTAFile =OSD_sScriptDir & "hta\" &"AppForm.hta"
	End If
	
' set the OSDProfile.ini file if not specified as an argument 
	If Not (OSD_bCreateINI Or OSD_bApplyINI) Then
		OSD_sOSDProfileIniFile = "OSDProfile.ini"
		If OSD_bRunningFromTS Then
			OSD_sOSDProfileIniFile = OSD_sLogPath & "OSDProfile.ini"
		Else
			OSD_sOSDProfileIniFile =  OSD_sScriptDir & "OSDProfile.ini"
		End If
	End If	

'check if running from TS and use the SCCM Log path for Log files and read some TS variables
	If OSD_bRunningFromTS Then
	
		WriteLog "Script running from a Task Sequence"
		
	    'get _SMSTSInWinPE variable value to determine if running script at WinPE or Full OS phase
	    OSD_sTSInWinPE=OSD_oTaskSequence("_SMSTSInWinPE")
		WriteLog "Task Sequence variable _SMSTSInWinPE=" &OSD_sTSInWinPEE
	    
	    'get CFI TS variable value to determine if running a Dell CMS deployment Scenario
	    OSD_varCFI=OSD_oTaskSequence("CFI")
	    WriteLog "Custom Task Sequence Variable CFI=" & OSD_varCFI
	    
	    WriteLog "close the task sequence progress bar (but TS is still running in back)"
	    On Error Resume Next
	    Err.Clear
	    Set OSD_oTSProgressUI = Wscript.CreateObject("Microsoft.SMS.TSProgressUI") 
		OSD_oTSProgressUI.CloseProgressDialog()
		On Error Goto 0
		
	Else
		
		WriteLog "Script not running from MDT or SCCM Task sequence."
		
	End If

'check if /ApplyProfile specified and run OSDCustomizer silently
	If OSD_bApplyINI Then
		
		WriteLog "Argument OSD_bApplyINI = '" & OSD_bApplyINI & "'"  
		WriteLog "OSDCustomizer should run in Zero Touch mode using the settings provided from OSDProfile.ini file."
		WriteLog "Looking for file : " & OSD_sOSDProfileIniFile
		
		If Not OSD_objFSO.FileExists(OSD_sOSDProfileIniFile) Then 
			WriteLog "File missing: " & OSD_sOSDProfileIniFile & ". OSDCustomizer will run in Lite Touch mode"			
		Else 
			WriteLog "Found File: " & OSD_sOSDProfileIniFile & ". OSDCustomizer will run in Zero Touch Mode"
		End If
		
	Else
	
		If OSD_bCreateINI Then
		
			WriteLog "Argument OSD_bCreateINI ='" & OSD_bCreateINI & "'"		
			WriteLog "Starting the FrontEnd.HTA form to prepare the OSD Profile INI file: " & OSD_sOSDProfileIniFile
			
			'run the OSDFrontEnd HTA form			
			If OSD_objFSO.FileExists(OSD_sOSDFrontEndHTAFile) Then
				
				OSD_sCmd="mshta.exe "  & Chr(34) &  OSD_sOSDFrontEndHTAFile	& Chr(34) & " /CreateProfile:" &  Chr(34) & OSD_sOSDProfileIniFile & Chr(34)
				WriteLog "About to run command: " & OSD_sCmd
				 
				OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
				
				If OSD_iRetVal <> OSD_Success then
					Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "OSD FrontEnd HTA"
					WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
					WriteLog "ERROR: the OSDFrontEnd.hta page cannot be displayed."
					OSD_iRetVal=OSD_Failure
				End If
				
			Else
			
				WriteLog "Error: Unable to find (OSDFrontEnd.hta) as specified here: " & OSD_sOSDFrontEndHTAFile
				Msgbox "Error: Unable to find (OSDFrontEnd.hta).",vbSystemModal, "OSD FrontEnd HTA"
				OSD_iRetVal=OSD_Failure
				
			End If

			'run the APPForm HTA if Application installation set in OSDsettings.ini
			sAppsInstall=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AppsInstall")))
			
			If sAppsInstall="YES" Then 

			    OSD_SkipAppsForm=Trim(ReadIni(OSD_sSettingsIniFile, "Applications", "SkipAppsForm"))
			
			    If Trim(UCase(OSD_SkipAppsForm))="NO" Or Len(Trim(OSD_SkipAppsForm))=0 Then	

				If OSD_objFSO.FileExists(OSD_sAppsHTAFile) Then
					OSD_sCmd="mshta.exe "  & Chr(34) &  OSD_sAppsHTAFile	& Chr(34) & " /CreateProfile:" &  Chr(34) & OSD_sOSDProfileIniFile & Chr(34)
					WriteLog "About to run command: " & OSD_sCmd
				 
					OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
				
					If OSD_iRetVal <> OSD_Success then
						
						WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
						WriteLog "ERROR: the APPForm.hta page cannot be displayed."
						Writelog "====================================================================================="
						WriteLog "Script " & OSD_sScriptName & " excution has been aborted."
						Writelog "====================================================================================="
						
						MsgBox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "APPForm HTA"
						
						OSD_iRetVal=OSD_Failure
					End If
				
				Else
					WriteLog "Error: Unable to find (AppForm.hta) as specified here: " & OSD_sAppsHTAFile
					MsgBox "Error: Unable to find (AppForm.hta).",vbSystemModal, "AppForm"
					OSD_iRetVal=OSD_Failure
				
				End If
			  Else
					WriteLog "Skipping AppForm.hta as SkipAppsForm=YES in config file: " & OSD_sSettingsIniFile
					
				
			   End If


			End If		

			If OSD_iRetVal=OSD_Failure Then
				Writelog "====================================================================================="
				WriteLog "Script " & OSD_sScriptName & " excution has been aborted."
				Writelog "====================================================================================="
				'WriteLog "Script execution aborted."	
				
						
				
				WScript.Quit
				
			Else
				
				If OSD_objFSO.FileExists(OSD_sOSDProfileIniFile) Then
					WriteLog "File: " & OSD_sOSDProfileIniFile & " is created."
					Writelog "====================================================================================="
					WriteLog "Script " & OSD_sScriptName & " excution is completed."
					Writelog "====================================================================================="
				Else
					WriteLog "Error: Unable to find the OSD Profile file as specified here: " & OSD_sOSDProfileIniFile
					
					WriteLog "====================================================================================="
					WriteLog "Script " & OSD_sScriptName & " excution has been aborted due to missing required file."
					Writelog "====================================================================================="
				End If
				
						
				WScript.Quit
				
			End If		
		
		End If
		
	End If

'read OSDSettings.ini and display a frontend hta form if DefaultGateway mode is not selected in OSDSettings.Ini file.
	OSD_oShell.CurrentDirectory = OSD_sScriptDir
	
	WriteLog "Getting OSD Configuration settings from: " & OSD_sSettingsIniFile
	OSD_iRetVal= get_Settings_From_INIFile
	If OSD_iRetVal <> OSD_Success Then
		WriteLog "ERROR - One or more settings not found..." & OSD_iRetVal	
	End If

'Check if DefaultGateway mode is selected then create an OSDProfile.ini file based on detected IP DefaultGateway address 
	If OSD_bDefaultGateway=True And OSD_skipWizard=True Then
		Create_OSDProfile_Using_DefaultGateway_Settings
	End If

'check if OSDProfile.ini exists otherwise exit
	If Not OSD_objFSO.FileExists(OSD_sOSDProfileIniFile) Then
		
		WriteLog "Error File Missing: " & OSD_sOSDProfileIniFile & " does not exist. Process aborted."
		Writelog "====================================================================================="
		WriteLog "Script " & OSD_sScriptName & " excution has been aborted due to missing required file."
		Writelog "====================================================================================="
		MsgBox "File " & OSD_sOSDProfileIniFile & " does not exist. Process aborted.",vbSystemModal, "File Error"
		
		OSD_iRetVal=OSD_Failure
		
						
		WScript.Quit
		
	ElseIf OSD_skipWizard=False Then
	
		'get OSD selected config from OSDProfile.ini file
		getOSDSelectionConfig
		
	End if

'apply settings based on OSDProfile.ini file
	If OSD_bRunningFromTS Then 		'Running from TS
		
					WriteLog "copy OSDProfile.ini to _SMSTSLogPath for troubleshooting if needed"
				    OSD_objFSO.CopyFile OSD_sOSDProfileIniFile, OSD_sLogPath & "OSDProfile.ini", true
					
					'Set TS variables and apply settings depending if TS in WinPE or Full OS
					If UCase(Trim(OSD_sTSInWinPE))= "TRUE" Then 	
					
							'TS running at WinPE phase.
							
							WriteLog "Task Sequence running at WinPE phase"
							WriteLog "Only OSD Variables are set according to OSDProfile.INI file"
							
							'assign OSD TS variables
							setOSDVariables
				  
					Else 
						
							'TS running at Full OS phase.
							
							WriteLog "TS running at Full OS phase"
								
							If Ucase(Trim(OSD_varCFI))="TRUE" Then 'This is considered as a Dell CFI deployment scenario
					
								WriteLog "Task Sequence variable CFI=True. This is considered as Dell CS SCCM/MDT Factory scenario."
								
								'assign OSD TS variables
									setOSDVariables
								
								'set regional settings
								   setRegionalOptions
								
								'Custom Command specific
								If Len(OSD_RunCommand) > 0 Then ExecuteCommand OSD_RunCommand
							
								' rename hostname if a new computername is specified
									If Len(Trim(OSD_sNewComputername)) > 0 Then
										OSD_sCurrentHostname=CreateObject("Wscript.network").ComputerName
										WriteLog "Current hostname is: " & OSD_sCurrentHostname
										Rename_Hostname OSD_sCurrentHostname, OSD_sNewComputername
										
									End If
									
							 Else 		
							 		'not CFI scenario
									WriteLog "Task Sequence variable CFI= " & OSD_varCFI
									WriteLog "This is not considered as Dell CS SCCM/MDT factory scenario"
									WriteLog "Only OSD Variables are set according to OSDProfile.INI file"
									
									'assign OSD TS variables
									setOSDVariables
									
									'set regional settings using online xml file
								   setRegionalOptions
								   
									
							 End If
							 	
					End If
		
	Else		'NOT running from TS
		
								WriteLog "Script not running from a Task Sequence. Check and getting the Sysprep Image State..."
								
								sOS_SysPrepStatekey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State\ImageState"
								
								'read the value of sOS_SysPrepState
								sOS_SysPrepState = UCase(Trim(OSD_oShell.RegRead(sOS_SysPrepStatekey)))
								WriteLog "Sysprep Image State= " & sOS_SysPrepState
								
								'check if applications are required for installation
									sAppsInstall=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AppsInstall")))
								
								If sOS_SysPrepState="IMAGE_STATE_COMPLETE" Then
										WriteLog "Sysprep is already completed. Not using Unattend.xml to apply the settings."
										
											'set regional settings using online xml file
											   setRegionalOptions
											 
											 ' rename hostname if a new computername is specified
												If Len(Trim(OSD_sNewComputername)) > 0 Then
													OSD_sCurrentHostname=CreateObject("Wscript.network").ComputerName
													WriteLog "Current hostname is: " & OSD_sCurrentHostname
													Rename_Hostname OSD_sCurrentHostname, OSD_sNewComputername
													
												End If
											
											'Custom Command specific
											If Len(OSD_RunCommand) > 0 Then ExecuteCommand OSD_RunCommand
											
											'add a shortcut in startup menu for OSDCUSTOMIZER.VBS
											SetStartOSDCustomizer
											 
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
				
											WriteLog "Sysprep is not completed. we will be updating entries in the Unattend.xml to apply the settings."
												
											'Set the templates xml paths based on the OS Architecture
											
											sOS_architect=sArchitecture
											
											WriteLog "PROCESSOR ARCHITECTURE= " & sOS_architect
											If Not UCase(sOS_architect)="X86" Then sOS_architect="X64"
											
											sDomainXML_template ="Domain_" & sOS_architect & ".xml"
											sNetworkInterfaceXML_template ="NetworkInterface_" & sOS_architect & ".xml"
												
											If sAppsInstall="YES" Then
												
													sOOBEXml_template ="OobeSystem_Autologon_" & sOS_architect & ".xml"
													sAutologonSetInUnattendXML=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonSetInUnattendXML")))
													
													If Not sAutologonSetInUnattendXML="YES" then
														
																sAutologonUser=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUser"))
																sAutologonUserPassword=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUserPassword"))						
																
																If Len(sAutologonUser)> 0 And Len(sAutologonUserPassword) > o Then
																	'add a shortcut in startup menu for OSDCUSTOMIZER.VBS
																	SetStartOSDCustomizer
																
																Else
																	
																	If Len(sAutologonUser)= 0 Then
																		sMsg= "Error: File OSDSettings.ini Section [Applications] Property 'AutologonUser' is empty. Process aborted."
																	End If
																	
																	If Len(sAutologonUserPassword)= 0 Then
																		sMsg= "Error: File OSDSettings.ini Section [Applications] Property 'AutologonUser' is empty. Process aborted."
																	End If
																	
																	WriteLog sMsg
																	MsgBox sMsg, vbSystemModal, "Wrong Setting" 
																	OSD_iRetVal=OSD_Failure 
																End If
																
													Else
																'add a shortcut in startup menu for OSDCUSTOMIZER.VBS
																SetStartOSDCustomizer
																
													End If		
											Else
													sOOBEXml_template ="OobeSystem_" & sOS_architect & ".xml"
											End If			
									
											If OSD_objFSO.FileExists(OSD_sScriptDir  & sDomainXML_template) Then
												sDomainXML_template =OSD_sScriptDir & sDomainXML_template
											ElseIf OSD_objFSO.FileExists(OSD_sScriptDir & "Templates\" & sDomainXML_template) Then
												sDomainXML_template =OSD_sScriptDir & "Templates\" & sDomainXML_template
											End If
											
											If OSD_objFSO.FileExists(OSD_sScriptDir  & sOOBEXml_template) Then
												sOOBEXml_template =OSD_sScriptDir & sOOBEXml_template
											ElseIf OSD_objFSO.FileExists(OSD_sScriptDir & "Templates\" &sOOBEXml_template) Then
												sOOBEXml_template =OSD_sScriptDir & "Templates\" & sOOBEXml_template
											End If
											
											If OSD_objFSO.FileExists(OSD_sScriptDir  & sNetworkInterfaceXML_template) Then
												sNetworkInterfaceXML_template =OSD_sScriptDir & sNetworkInterfaceXML_template
											ElseIf OSD_objFSO.FileExists(OSD_sScriptDir & "Templates\" &sNetworkInterfaceXML_template) Then
												sNetworkInterfaceXML_template =OSD_sScriptDir & "Templates\" & sNetworkInterfaceXML_template
											End If
											
											WriteLog "Used domain answer file template as: " & sDomainXML_template
											WriteLog "Used oobe answer file template as: " & sOOBEXml_template
											
											'update the unattend.xml and GeoID
											WriteLog "Searching for an existing " & OSD_RootDrv & "\Windows\Panther\Unattend.xml"
											sUnattendXml=OSD_RootDrv & "\Windows\Panther\Unattend.xml"
									
											If OSD_objFSO.FileExists(sUnattendXml) Then
													'WriteLog "Unattend.xml file found as:" & sUnattendXml
													'update the unattend.xml
													UpdateUnattendxml
													'Update GeoID if required
													If Len(OSD_sCountry) > 0 Then
														sGeoIDXml=OSD_sScriptDir & "GeoID.xml"							
														CreateGeoIDXML OSD_sCountry
														'run the command to update the GeoID for current user and default user
														If OSD_objFSO.FileExists(sGeoIDXml) then
															OSD_sCmd="cmd /c control intl.cpl,, /f:" & Chr(34) & sGeoIDXml & Chr(34)
															On Error Resume Next
															WriteLog "About to run command: " & OSD_sCmd
								        					ExecuteCommand OSD_sCmd
														Else	
															WriteLog "ERROR - File not found:" & sGeoIDXml
														End If
													End If
													'turn setup flag back on
														OSD_oShell.RegWrite "HKLM\SYSTEM\Setup\SystemSetupInProgress", 1, "REG_DWORD"			
													'run Windeploy to get Sysprep to continue the Windows Setup
														OSD_sCmd=OSD_RootDrv & "\Windows\System32\Oobe\Windeploy.exe"
														WriteLog "About to run command: " & OSD_sCmd
														OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
														If OSD_iRetVal <> OSD_Success then
															Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "OSDCustomizer"
															WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal				
															OSD_iRetVal=OSD_Failure
														End If
											Else
													WriteLog "Error: " & OSD_RootDrv & "\Windows\Panther\Unattend.xml not found. Selected settings are not applied."
													Msgbox "Answer file: " & OSD_RootDrv & "\Windows\Panther\Unattend.xml not found" ,vbSystemModal, "Error"
													'turn setup flag back on
														OSD_oShell.RegWrite "HKLM\SYSTEM\Setup\SystemSetupInProgress", 1, "REG_DWORD"
														'run Windeploy to get Sysprep to continue the Windows Setup
														OSD_sCmd=OSD_RootDrv & "\Windows\System32\Oobe\Windeploy.exe"
														WriteLog "About to run command: " & OSD_sCmd
														OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
														If OSD_iRetVal <> OSD_Success then
															Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "OSDCustomizer"
															WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
															OSD_iRetVal=OSD_Failure
														End If
											End If
									End If
	End If

	If OSD_iRetVal=OSD_Failure Then
			Writelog "====================================================================================="
			WriteLog "Script " & OSD_sScriptName & " excution has been aborted."
			Writelog "====================================================================================="
			
						
			
			WScript.Quit
	End If

'script completed
	If sRunPost="NO" Then 
	
		'WriteLog "OSDCustomizer script execution completed."
 		Writelog "====================================================================================="
		WriteLog "Script " & OSD_sScriptName & " excution is completed."
		Writelog "====================================================================================="
 	Else
 	
 		WriteLog "Script " & OSD_sScriptName & " PHASE 1 PRE Sysprep excution is completed."
 		WriteLog "Phase 2 should continue as POST sysprep at first logon."
 		WriteLog "=============================================================================="
 		
 	End If
	
						
	
'exit
	WScript.Quit


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'    SUB & FUNCTIONS
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Function get_Settings_From_INIFile

	'check if OSD_bApplyINI exists
	If OSD_bApplyINI Then
		get_Settings_From_INIFile = OSD_Success
		Exit Function
	End If

	'get the default ComputerNaming section from OSDSettings.ini
	OSD_sDesktop=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Desktop"))
	OSD_sLaptop=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Laptop"))
	OSD_sConstruct=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Construct"))
	If Len(OSD_sConstruct) > 0 Then 
		WriteLog "Computernaming convention is set as: " & OSD_sConstruct
	End If

	'get APPS item
	sAppsInstall=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AppsInstall")))
	If sAppsInstall="YES" Then
	
		sAutologonSetInUnattendXML=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonSetInUnattendXML")))
		If Not sAutologonSetInUnattendXML="YES" then
			sAutologonUser=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUser"))
			sAutologonPwd=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUserPassword"))
			
			If Len(sAutologonUser)= 0 Then
				WriteLog "Error: Section [Applications] Property AutologonUser is not set correctly."
				get_Settings_From_INIFile=OSD_Failure
				Exit Function
				
			End If
		End If
		sAPPSPerLocation=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AppsPerLocation")))
		sAPPSPerProfile=Trim(UCase(ReadIni(OSD_sSettingsIniFile, "Applications","AppsPerProfile")))
	End If
	
	' determine if OSDSettings.ini is configured for DefaultGateway mode deployment
	If OSD_objFSO.FileExists(OSD_sSettingsIniFile) Then	
		If Left(UCase(ReadIni(OSD_sSettingsIniFile,"General","DefaultGateway")),1)="Y" Then
			OSD_bDefaultGateway =True
			WriteLog "DefaultGateway mode is set in [General] section of OSDSettings.ini: DefaultGateway=" & UCase(ReadIni(OSD_sSettingsIniFile,"General","DefaultGateway"))
		End If
	Else
		WriteLog "Error: Unable to find (" & OSD_sSettingsIniFile & ")" 
		get_Settings_From_INIFile=OSD_Failure
		Exit Function
	End If

	' Get selected settings using IP DefaultGateway section details
	If OSD_bDefaultGateway=True Then
		WriteLog "Getting network details (IP addresses and DefaultGateway) using WMI on current system..." 
		OSD_bDefaultGateway=False
		Set OSD_dicIPAddresses = CreateObject("Scripting.Dictionary")
		Set OSD_dicDefaultGateway = CreateObject("Scripting.Dictionary")
		Set OSD_objWMIService = GetObject("winmgmts:\\" & OSD_strComputer & "\root\cimv2")
		
		If OSD_objWMIService is Nothing then
			WriteLog "Unable to obtain network details (IP addresses and DefaultGateway) since WMI is unavailable."
			OSD_iRetVal=OSD_Failure
			OSD_skipWizard = False
		Else
			'Get network details
			OSD_iRetVal = GetNetworkDetails(OSD_dicIPAddresses, OSD_dicDefaultGateway)
			If OSD_iRetVal = OSD_Success Then
					OSD_bDefaultGateway=True
					WriteLog "------ Processing the [DefaultGateway] section in OSDSettings.ini------"
					
					' Check each default gateway value to see if a match can be found and at first match then other DefaultGateway are ignored
							For each sElement in OSD_dicDefaultGateway
								
								'get the associated site name
								OSD_sSite=Trim(ReadIni(OSD_sSettingsIniFile, "DefaultGateway", sElement))
								
								if Len(OSD_sSite) = 0 Then
									WriteLog sElement & " DefaultGateway not defined in the section [DefaultGateway]"
									OSD_bDefaultGateway=False
								Else
									WriteLog " DefaultGateway " & sElement &"=" & OSD_sSite
									OSD_bDefaultGateway=True
									Exit For
								End If
							Next
								
					'get the associated value from OSDSettings.ini for the selected Site	
					
					If OSD_bDefaultGateway=True Then
						
							WriteLog "------ Processing the [" & OSD_sSite & "] section in OSDSettings.ini------"
							OSD_sSystemLocale=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "SystemLocale"))
							If Len(OSD_sSystemLocale) > 0 Then WriteLog "SystemLocale=" & OSD_sSystemLocale
							
							OSD_sUserLocale=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "UserLocale"))
							If Len(OSD_sUserLocale) > 0 Then WriteLog "UserLocale=" & OSD_sUserLocale
							
							OSD_sInputLocale=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "InputLocale"))
							If Len(OSD_sInputLocale) > 0 Then WriteLog "InputLocale=" & OSD_sInputLocale
							
							OSD_sUILanguage=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "UILanguage"))
							If Len(OSD_sUILanguage) > 0 Then WriteLog "UILanguage=" & OSD_sUILanguage
							
							OSD_sCountry=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Country"))
							If Len(OSD_sCountry) > 0 Then WriteLog "Country=" & OSD_sCountry
							
							'get type
							GetType
							
							'set ComputerName
							
							OSD_sPrefix=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Prefix"))
							If Len(OSD_sPrefix)>0 Then	
								WriteLog "Prefix=" & OSD_sPrefix
							End If
							
							OSD_sSuffix=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Suffix"))
							If Len(OSD_sSuffix) > 0 Then	
								WriteLog "Suffix=" & OSD_sSuffix
							End If
										
							OSD_sNewComputername= OSD_sPrefix & OSD_sSuffix
									
							If InStr (1,UCase(OSD_sNewComputername), "SERVICE_TAG",1) > 0 Then 	
								OSD_sNewComputername=Replace(UCase(OSD_sNewComputername),"SERVICE_TAG",OSD_sServiceTag)
							End If
										
							If InStr (1,UCase(OSD_sNewComputername), "ASSET_TAG",1) > 0 Then	
								OSD_sNewComputername=Replace(UCase(OSD_sNewComputername),"ASSET_TAG",OSD_sAssetTag) 
							End If
																							
							If Len(OSD_sConstruct) > 0 Then
								OSD_sNewComputername=GetComputerName
							End If
							
							OSD_sTimeZone=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "TimeZone"))
							If Len(OSD_sTimeZone) > 0 Then WriteLog "TimeZone=" & OSD_sTimeZone
							
							'IP address, Gateway, PrimDNs, Sec DNS
							OSD_sIPAddress=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "IPAddress"))
							If Len(OSD_sIPAddress) > 0 Then WriteLog "IPAddress=" & OSD_sIPAddress
							
							OSD_sGatewayAddress=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "GatewayAddress"))
							If Len(OSD_sGatewayAddress) > 0 Then WriteLog "GatewayAddress=" & OSD_sGatewayAddress
							
							OSD_sPrimDNS=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "PrimaryDNS"))
							If Len(OSD_sPrimDNS) > 0 Then WriteLog "PrimaryDNS=" & OSD_sPrimDNS
							
							OSD_sSecDNS=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "SecondaryDNS"))
							If Len(OSD_sSecDNS) > 0 Then WriteLog "SecondaryDNS=" & OSD_sSecDNS
							
							
							'Workgroup /Domain details
							OSD_sJoin=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Join"))
							If Len(OSD_sJoin) > 0 Then WriteLog "Join=" & OSD_sJoin
							
							If UCase(OSD_sJoin)=UCase("Workgroup") Then
									OSD_sWorkgroup=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Workgroup"))
									If Len(OSD_sWorkgroup) > 0 Then WriteLog "Workgroup=" & OSD_sWorkgroup
							ElseIf UCase(OSD_sJoin)=UCase("Domain") Then
									OSD_sDomain=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "Domain"))
									If Len(OSD_sDomain) > 0 Then WriteLog "Domain=" & OSD_sDomain
									
									OSD_sDomainOU=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "DomainOU"))
									If Len(OSD_sDomainOU) > 0 Then WriteLog "DomainOU=" & OSD_sDomainOU
									
									OSD_sDomainUser=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "DomainUser"))
									If Len(OSD_sDomainUser) > 0 Then WriteLog "DomainUser=" & OSD_sDomainUser
									
									OSD_sDomainUserPassword=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, "DomainUserPassword"))
									If Len(OSD_sDomainUserPassword) > 0 Then WriteLog "DomainUserPassword=" & OSD_sDomainUserPassword
							End If
							
							'add Applicationxxx from OSDSettings.ini if default gateway found
							If sAppsInstall="YES" Then
									i=0
									bStop=False
									Do While bStop=False
										 	If i=0 Then
												ReDim sArrayApps(i)
											Else
												ReDim Preserve sArrayApps(i)
											End If
											If Len(i)=1 Then
												sAppProperty1="Application00" & i+1
											End If
											If Len(i)=2 Then
												sAppProperty2="Application0" & i+1
											End If
											If Len(i)=3 Then
												sAppProperty3="Application" & i+1
											End If
											sArrayApps(i)=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty1))
											If Len(sArrayApps(i))=0 Then
												sArrayApps(i)= Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty2))
												If Len (sArrayApps(i))=0 Then
													sArrayApps(i)= Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty3))
													If Len (sArrayApps(i))=0 Then
														bStop=True
														'numCmd=i
													Else
														WriteLog "Found " & sAppProperty3 & "=" & sArrayApps(i)
														i=i+1
													End If
												Else
													WriteLog "Found " & sAppProperty2 & "=" & sArrayApps(i)
													i=i+1
												End If	
											Else
												WriteLog "Found " & sAppProperty1 & "=" & sArrayApps(i)
												i=i+1
											End If
										Loop
								End If
							
							OSD_skipWizard =True
							get_Settings_From_INIFile = OSD_Success
							Exit Function
					End If
			End If
				
		End If

	End If
	
  ' start the Wizard if not using DefaultGateway settings

  If Not OSD_skipWizard Then
		
	'display or not KeybLayout.hta
	
		OSD_sSkipKeyboardForm=Trim(ReadIni(OSD_sSettingsIniFile, "General", "SkipKeyboardForm"))
		
		If Trim(UCase(OSD_sSkipKeyboardForm))="NO" Or Len(Trim(OSD_sSkipKeyboardForm))=0 Then
			WriteLog "Looking for file KeybLayout.hta page ..."
			
			If OSD_objFSO.FileExists(OSD_sKeybHTAFile) Then
				WriteLog "Displaying the KeybLayout.hta page for keyboard layout selection..."
				OSD_sCmd="cmd /c mshta.exe " & Chr(34) & OSD_sKeybHTAFile & Chr(34)
				OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
 WScript.Sleep 3000
				
				If OSD_iRetVal <> OSD_Success then
					'Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "KeybLayout.hta Form"
					WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
					WriteLog "ERROR: the KeybLayout.hta Form cannot be displayed."
					get_Settings_From_INIFile=OSD_Failure
					
				End If
				
			Else
			
				WriteLog "Warning: Unable to find file: " & OSD_sKeybHTAFile
				WriteLog "Skipping the display of KeybLayout.hta Form"
				OSD_iRetVal=OSD_Failure
				get_Settings_From_INIFile=OSD_Failure	
				
			End If
			
		Else
		
			WriteLog "Property 'SkipKeyboardForm=YES'. Skipping the display of KeybLayout.hta Form"
			'get the KeyboardLayout setting
			sKeyboardLayout=Trim(ReadIni(OSD_sSettingsIniFile, "General", "KeyboardLayout"))
			If sKeyboardLayout <> "" Then
				setCurrentKeyboard sKeyboardLayout
			End If
			
		End If
		
		 WScript.Sleep 3000
	'run the OSDFrontEnd HTA form
		OSD_oShell.CurrentDirectory = OSD_sScriptDir
		If RunOSDFrontEnd = "YES" Then
		
					WriteLog "Executing the OSDFrontEnd.hta page... from file: '" & OSD_sOSDFrontEndHTAFile &"'"
							
					If OSD_objFSO.FileExists(OSD_sOSDFrontEndHTAFile) Then
						OSD_sCmd="cmd /c mshta.exe "  & Chr(34) &  OSD_sOSDFrontEndHTAFile	& Chr(34)
												
						OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
						
						If OSD_iRetVal <> OSD_Success then
							Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "OSD FrontEnd HTA"
							WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
							WriteLog "ERROR: the OSDFrontEnd.hta page cannot be displayed."
							get_Settings_From_INIFile=OSD_Failure
						
						Else
							get_Settings_From_INIFile=OSD_Success	
						End If
						
					Else
					
						WriteLog "Error: Unable to find (OSDFrontEnd.hta)."
						Msgbox "Error: Unable to find (OSDFrontEnd.hta).",vbSystemModal, "OSD FrontEnd HTA"
						OSD_iRetVal=OSD_Failure
						get_Settings_From_INIFile=OSD_Failure
				
					End If	
		End If
			
		'run the Appform HTA
		
		If sAppsInstall="YES" Then
		
				OSD_SkipAppsForm=Trim(ReadIni(OSD_sSettingsIniFile, "Applications", "SkipAppsForm"))
			
				If Trim(UCase(OSD_SkipAppsForm))="NO" Or Len(Trim(OSD_SkipAppsForm))=0 Then
							
						WriteLog "Executing the AppForm.hta page... from file: '" & OSD_sAppsHTAFile &"'"
								
						If OSD_objFSO.FileExists(OSD_sAppsHTAFile) Then
							OSD_sCmd="cmd /c mshta.exe "  & Chr(34) &  OSD_sAppsHTAFile	& Chr(34)
							
							OSD_iRetVal = OSD_oShell.Run(OSD_sCmd, 0, true)
							
							If OSD_iRetVal <> OSD_Success then
								Msgbox "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal,vbSystemModal, "APPForm HTA"
								WriteLog "ERROR - Execution command:" & OSD_sCmd & " returned a non-zero return code, rc = " & OSD_iRetVal
								WriteLog "ERROR: the AppForm.hta page cannot be displayed."
								get_Settings_From_INIFile=OSD_Failure
							
							Else
								get_Settings_From_INIFile=OSD_Success	
							End If
							
						Else
						
							WriteLog "Error: Unable to find (AppForm.hta)."
							Msgbox "Error: Unable to find (AppForm.hta).",vbSystemModal, "AppForm"
							OSD_iRetVal=OSD_Failure
							get_Settings_From_INIFile=OSD_Failure
					
						End If
				Else
						WriteLog "Property 'SkipAppsForm=YES'. Skipping the display of Applications Apps.hta Form"
				End If 	
				
		End If
 	
 	End If
	
	OSD_oShell.CurrentDirectory = OSD_sScriptDir	
		
End Function

Function RunOSDFrontEnd

	RunOSDFrontEnd="YES"

	'get configuration settings from OSDSettings.ini
	WriteLog "Getting configuration settings from file: " & OSD_sSettingsIniFile
	
	'get SkipLocationName
	SkipLocationName=UCase(Trim(ReadIni(OSD_sSettingsIniFile, "General","SkipLocationName")))
	WriteLog "Section [General] Property SkipLocationName= " & SkipLocationName
	
	
	'get SkipRegionalSettings
	SkipRegionalSettings=UCase(Trim(ReadIni(OSD_sSettingsIniFile, "General", "SkipRegionalSettings")))
	
	
	If SkipRegionalSettings="YES" Then
		SkipUserLocale="YES"
		SkipSystemLocale="YES"
		SkipInputLocale="YES"
		SkipUILanguage="YES"
		SkipCountry="YES"
		SkipTimeZone="YES"
		
		WriteLog "Skipping diskplay of RegionalSettings fields due to seetings in OSDSettings.ini file."
		
	Else
		
		If SkipUserLocale="YES"And SkipSystemLocale="YES" And _
		 SkipInputLocale="YES" And SkipUILanguage="YES" And _
		 SkipCountry="YES" And SkipTimeZone="YES" Then
		 
			SkipRegionalSettings="YES"
			WriteLog "Skipping diskplay of RegionalSettings fields due to seetings in OSDSettings.ini file."
		
		End If
		
	End If
	
	
	'get SkipComputerName
	SkipComputerName=UCase(Trim(ReadIni(OSD_sSettingsIniFile, "General", "SkipComputerName")))
	If SkipComputerName="YES" Then
		WriteLog "Skipping diskplay of Computername field due to seetings in OSDSettings.ini file."
			
	End If
		
	'get SkipDomainJoin
	SkipDomainJoin=UCase(Trim(ReadIni(OSD_sSettingsIniFile, "General", "SkipDomainJoin")))
	If SkipDomainJoin ="YES" Then
		SkipDomainName="YES"
		SkipDomainOU="YES"
		SkipDomainUser="YES"
		SkipDomainUserPassword="YES"
		
		WriteLog "Skipping diskplay of Domain Join fields due to seetings in OSDSettings.ini file."
		
	Else
		If SkipDomainName="YES"And SkipDomainOU="YES" And _
		 SkipDomainUser="YES" And SkipDomainUserPassword="YES" Then
			SkipDomainJoin="YES"
			WriteLog "Skipping diskplay of Domain Join fields due to seetings in OSDSettings.ini file."
			
		End If
		
	End If

	If SkipRegionalSettings="YES" And SkipComputerName="YES" And SkipDomainJoin="YES" And SkipLocationName= "YES" Then
			RunOSDFrontEnd="NO"
			WriteLog "Skipping diskplay of OSDFrontEnd Form due to seetings in OSDSettings.ini file."
			
	End If
	
	'check if AppsPerlocation or sAPPSPerProfile is set to YES then set RunOSDFrontEnd="YES"
	sAppsPerLocation= UCase(Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AppsPerLocation")))
	sAppsPerProfile= UCase(Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AppsPerProfile")))
	
	WriteLog "Section [Applications] Property AppsPerLocation= " & sAppsPerLocation
	WriteLog "Section [Applications] Property AppsPerProfile= " & sAppsPerProfile
	
	If (sAppsPerLocation="YES" Or sAppsPerProfile="YES" ) And Not(RunOSDFrontEnd)="YES" Then
			RunOSDFrontEnd="YES"
			WriteLog "Display of OSDFrontEnd Form to select applications based on location / Profile name."
			
	End If

End Function


' --------------------------------------------
' if scripting object "Microsoft.SMS.TSEnvironment" can be created
' then Returnvalue = true and Object is created
' else Returnvalue = false
' --------------------------------------------
Function Script_Started_from_TS
	
	Script_Started_from_TS = False	
	'Exit Function 
    Err.Clear
	On Error Resume Next
    Set OSD_oTaskSequence = CreateObject("Microsoft.SMS.TSEnvironment")
	If Err.Number  <> 0 Then
		WriteLog "Not running from Task sequence. Script_Started_from_TS ='" & Script_Started_from_TS & "'"
		On Error Goto 0
		Script_Started_from_TS = False	
		Exit Function
	End If
	On Error Goto 0
	Script_Started_from_TS  = True
	WriteLog "Running from Task sequence. Script_Started_from_TS ='" & Script_Started_from_TS & "'"
		
End Function


Sub getOSDSelectionConfig
	Dim sSelSection
	Dim OSD_sCmd, i
  
	if OSD_objFSO.FileExists(OSD_sOSDProfileIniFile) = True Then
	
		WriteLog "Getting selected settings from OSDProfile.ini file..."
		
		'read  the main selected section
			sSelSection=Trim(ReadIni(OSD_sOSDProfileIniFile,"Main","Selected"))
			WriteLog "Selected=" & sSelSection
			
		'get Properties
			OSD_sRegion=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"REGION"))
			If Len(OSD_sRegion) > 0 Then WriteLog "REGION=" & OSD_sRegion
			
			OSD_sLocationName=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Location"))
			If Len(OSD_sLocationName) > 0 Then WriteLog "Location=" & OSD_sLocationName
			
			OSD_RunCommand=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"RunCommand"))
			If Len(OSD_RunCommand) > 0 Then WriteLog "RunCommand=" & OSD_RunCommand
			
			OSD_sSystemLocale=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"SystemLocale"))
			If Len(OSD_sSystemLocale) > 0 Then WriteLog "SystemLocale=" & OSD_sSystemLocale
			
			OSD_sUserLocale=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"UserLocale"))
			If Len(OSD_sUserLocale) > 0  Then WriteLog "UserLocale=" & OSD_sUserLocale
			
			OSD_sInputLocale=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"InputLocale"))
			If Len(OSD_sInputLocale) > 0 Then WriteLog "InputLocale=" & OSD_sInputLocale
			
			OSD_sUILanguage=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"UILanguage"))
			If Len(OSD_sUILanguage) > 0 Then WriteLog "UILanguage=" & OSD_sUILanguage
			
			OSD_sCountry=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Country"))
			If Len(OSD_sCountry) > 0 Then WriteLog "Country=" & OSD_sCountry
			
			OSD_sTimeZone=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"TimeZone"))
			If Len(OSD_sTimeZone) > 0 Then WriteLog "TimeZone=" & OSD_sTimeZone
			
			OSD_sIPAddress=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"IPAddress"))
			If Len(OSD_sIPAddress) > 0 Then WriteLog "IPAddress=" & OSD_sIPAddress
			
			OSD_sGatewayAddress=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"GatewayAddress"))
			If Len(OSD_sGatewayAddress) > 0 Then WriteLog "GatewayAddress=" & OSD_sGatewayAddress
			
			OSD_sPrimDNS=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"PrimaryDNS"))
			If Len(OSD_sPrimDNS) > 0 Then WriteLog "PrimaryDNS=" & OSD_sPrimDNS
			
			OSD_sSecDNS=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"SecondaryDNS"))
			If Len(OSD_sSecDNS) > 0 Then WriteLog "SecondaryDNS=" & OSD_sSecDNS
			
			OSD_sJoin=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Join"))
			If Len(OSD_sJoin) > 0 Then WriteLog "Join=" & OSD_sJoin
			
			If UCase(OSD_sJoin)=UCase("Workgroup") Then
					OSD_sWorkgroup=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Workgroup"))
					If Len(OSD_sWorkgroup) > 0 Then WriteLog "Workgroup=" & OSD_sWorkgroup
			Elseif UCase(OSD_sJoin)=UCase("Domain") Then
					OSD_sDomain=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Domain"))
					If Len(OSD_sDomain) > 0 Then WriteLog "Domain=" & OSD_sDomain
					
					OSD_sDomainOU=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"DomainOU"))
					If Len(OSD_sDomainOU) > 0 Then WriteLog "DomainOU=" & OSD_sDomainOU
						
					OSD_sDomainUser=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"DomainUser"))
					If Len(OSD_sDomainUser) > 0 Then WriteLog "DomainUser=" & OSD_sDomainUser
						
					OSD_sDomainUserPassword=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"DomainUserPassword"))
					If Len(OSD_sDomainUserPassword) > 0 Then 
						WriteLog "DomainUserPassword=*********************"
					End If
			End If
			
			OSD_sNewComputername=Trim(ReadIni(OSD_sOSDProfileIniFile,sSelSection,"Computername"))
		
			If InStr (1,UCase(OSD_sNewComputername), "SERVICE_TAG",1) > 0 Then 	
				OSD_sNewComputername=Replace(UCase(OSD_sNewComputername),"SERVICE_TAG",OSD_sServiceTag) 	
			End If
				
			If InStr (1,UCase(OSD_sNewComputername), "ASSET_TAG",1) > 0 Then		
				OSD_sNewComputername=Replace(UCase(OSD_sNewComputername),"ASSET_TAG",OSD_sAssetTag) 
			End If
			
			If Len(OSD_sNewComputername) > 0 Then WriteLog "Computername=" & OSD_sNewComputername
			
			If sAppsInstall="YES" Then
				i=0
				bStop=False
				
				Do While bStop=False
				
					 	If i=0 Then
							ReDim sArrayApps(i)
						Else
							ReDim Preserve sArrayApps(i)
						End If
						If Len(i)=1 Then
							sAppProperty1="Application00" & i+1
						End If
						If Len(i)=2 Then
							sAppProperty2="Application0" & i+1
						End If
						If Len(i)=3 Then
							sAppProperty3="Application" & i+1
						End If
						
						sArrayApps(i)=Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty1))
						If Len(sArrayApps(i))=0 Then
							sArrayApps(i)= Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty2))
							If Len (sArrayApps(i))=0 Then
								sArrayApps(i)= Trim(ReadIni(OSD_sSettingsIniFile, OSD_sSite, sAppProperty3))
								If Len (sArrayApps(i))=0 Then
									bStop=True
								Else
									WriteLog "Found " & sAppProperty3 & "=" & sArrayApps(i)
									i=i+1
								End If
							Else
								WriteLog "Found " & sAppProperty2 & "=" & sArrayApps(i)
								i=i+1
							End If	
						Else
							WriteLog "Found " & sAppProperty1 & "=" & sArrayApps(i)
							i=i+1
						End If				
				Loop
				
			End If
	Else
	
		WriteLog "Error: OSDProfile.ini file not found."
		
	End If
			
End Sub

Sub setOSDVariables
	
	WriteLog "*** setOSDVariables() - started..."
	
	If Script_Started_from_TS() = True Then 

				'set Custom variable REGION 
				
					If Len(Trim(OSD_sRegion)) > 0 Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("REGION")=OSD_sRegion
						Else
							OSD_oTaskSequence ("REGION") = OSD_sRegion
						End If
						WriteLog "TS variable REGION = " & OSD_sRegion
					End If
				
				'set OSD variables for language settings
				
					If Len(Trim(OSD_sSystemLocale)) > 0 Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("SystemLocale")=OSD_sSystemLocale
							WriteLog "TS variable SystemLocale=" & OSD_sSystemLocale
						Else
							OSD_oTaskSequence ("OSDSystemLocale") = OSD_sSystemLocale
							WriteLog "TS variable OSDSystemLocale = " & OSD_sSystemLocale
						End If
					End If
					
					If Len(Trim(OSD_sUserLocale)) > 0 Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("UserLocale")=OSD_sUserLocale
							WriteLog "TS variable UserLocale=" & OSD_sUserLocale
						Else
							OSD_oTaskSequence ("OSDUserLocale") = OSD_sUserLocale
							WriteLog "TS variable OSDUserLocale = " & OSD_sUserLocale
						End if
					End If
					
					If Len(Trim(OSD_sInputLocale)) > 0  Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("InputLocale")=OSD_sInputLocale
							WriteLog "TS variable InputLocale=" & OSD_sInputLocale
						Else
							OSD_oTaskSequence ("OSDInputLocale") = OSD_sInputLocale
							WriteLog "TS variable OSDInputLocale = " & OSD_sInputLocale
						End if
					End If
					
					If Len(Trim(OSD_sUILanguage)) > 0  Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("UILanguage")=OSD_sUILanguage
							WriteLog "TS variable UILanguage=" & OSD_sUILanguage
						Else
							OSD_oTaskSequence ("OSDUILanguage") = OSD_sUILanguage
							WriteLog "TS variable OSDUILanguage = " & OSD_sUILanguage
						End if
					End If
					
				'set OSDTimeZone
					If Len(Trim(OSD_sTimeZone)) > 0 Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("TimeZoneName")=OSD_sTimeZone
							WriteLog "TS variable TimeZoneName=" & OSD_sTimeZone
						Else
							OSD_oTaskSequence ("OSDTimeZone") = OSD_sTimeZone
							WriteLog "TS variable OSDTimeZone = " & OSD_sTimeZone
						End if
					End If
				
				'set OSDComputername TS variable
					If Len(Trim(OSD_sNewComputername)) > 0  Then
						If OSD_IsMDT="YES" Then
							oEnvironment.Item("SkipComputerName")="NO"
							oEnvironment.Item("OSDComputerName")=OSD_sNewComputername
							WriteLog "TS variable OSDComputerName=" & OSD_sNewComputername
						Else
							OSD_oTaskSequence ("OSDComputerName") = OSD_sNewComputername
							WriteLog "TS variable OSDComputerName = " & OSD_sNewComputername
						End if
					End If
				
				'set CFIPostJoin custom variable
				If Len(Trim(OSD_sWorkgroup)) > 0 Then
					
					If OSD_IsMDT="YES" Then
					
						oEnvironment.Item("CFIPostJoin")="Workgroup"
						WriteLog "TS variable CFIPostJoin=" & oEnvironment.Item("CFIPostJoin")
						
						oEnvironment.Item("JoinWorkgroup")=OSD_sWorkgroup
						WriteLog "TS variable JoinWorkgroup=" & oEnvironment.Item("JoinWorkgroup")
						
					Else
					
						OSD_oTaskSequence("CFIPostJoin")="Workgroup"
						WriteLog "TS variable CFIPostJoin=" & OSD_oTaskSequence("CFIPostJoin")
						
						'OSD_oTaskSequence ("OSDWorkGroupName") = OSD_sWorkgroup
						'WriteLog "TS variable OSDWorkGroupName = " &  OSD_sWorkgroup
						
						OSD_oTaskSequence ("OSDJoinWorkgroupName") = OSD_sWorkgroup
						WriteLog "TS variable OSDJoinWorkgroupName = " &  OSD_sWorkgroup
						
						'OSD_oTaskSequence ("OSDNetworkJoinType") = "1"
						'WriteLog "TS variable OSDNetworkJoinType = 1"
						
						OSD_oTaskSequence ("OSDJoinType") = "1"
						WriteLog "TS variable OSDJoinType = 1"
					End If
				
				ElseIf UCase(OSD_sJoin)=UCase("Domain") Then
						If OSD_IsMDT="YES" Then
							
			    			oEnvironment.Item("SkipDomainMembership")="NO"
			    			oEnvironment.Item("JoinWorkgroup")=""	
							oEnvironment.Item("CFIPostJoin")="Domain"
							WriteLog "TS variable CFIPostJoin=" & oEnvironment.Item("CFIPostJoin")	
						Else
							OSD_oTaskSequence("CFIPostJoin")="Domain"
							WriteLog "TS variable CFIPostJoin=" & OSD_oTaskSequence("CFIPostJoin")
						End If
						
						'set OSDDomainName (SCCM), OSDJoinDomainName (SCCM), JoinDomain for system (MDT), DomainAdminDomain for user account (MDT) TS variable
							If Len(Trim(OSD_sDomain)) > 0 Then
								'check if MDt and domainuser contain \ then extract domain name for the account
								If OSD_IsMDT="YES" Then
								
									If Len(Trim(OSD_sDomainUser)) > 0  Then	
										If InStr (1, OSD_sDomainUser, "\", 1) > 0 Then
											sSplitOSD_sDomainUser=Split(OSD_sDomainUser , "\")
											oEnvironment.Item("DomainAdminDomain")=sSplitOSD_sDomainUser(0)	
										Else
											oEnvironment.Item("DomainAdminDomain")=OSD_sDomain
										End If
									End If
									
									WriteLog "TS variable DomainAdminDomain=" & oEnvironment.Item("DomainAdminDomain")
									oEnvironment.Item("JoinDomain")=OSD_sDomain
									WriteLog "TS variable JoinDomain=" & oEnvironment.Item("JoinDomain")
									
								Else
									OSD_oTaskSequence ("OSDDomainName") = OSD_sDomain
									WriteLog "TS variable OSDDomainName = " &  OSD_sDomain
									
									OSD_oTaskSequence ("OSDJoinDomainName") = OSD_sDomain
									WriteLog "TS variable OSDJoinDomainName = " &  OSD_sDomain
									
									OSD_oTaskSequence ("OSDJoinType") = "0"
									WriteLog "TS variable OSDJoinType = 0"
								End If
							End If
							
						' for SCCM set OSDDomainOUName in format like LDAP://OU=MyOu,DC=MyDom,DC=MyCompany,DC=com
						' for MDT set MachineObjectOU in format like OU=MyOu,DC=MyDom,DC=MyCompany,DC=com
							If Len(Trim(OSD_sDomainOU)) > 0 Then
								If OSD_IsMDT="YES" Then
									oEnvironment.Item("MachineObjectOU")= OSD_sDomainOU
									WriteLog "TS variable MachineObjectOU=" & OSD_sDomainOU
								Else
									OSD_oTaskSequence ("OSDDomainOUName") = "LDAP://" & OSD_sDomainOU
									WriteLog "TS variable OSDDomainOUName = " & "LDAP://" & OSD_sDomainOU
									OSD_oTaskSequence ("OSDJoinDomainOUName") = "LDAP://" & OSD_sDomainOU
									WriteLog "TS variable OSDJoinDomainOUName = " & "LDAP://" & OSD_sDomainOU
								End If
							End If
							
						'set OSDJoinAccount (SCCM), DomainAdmin (MDT)
							If Len(Trim(OSD_sDomainUser)) > 0  Then
							
									If InStr (1, OSD_sDomainUser, "\", 1) > 0 Then
									
										If OSD_IsMDT="YES" Then
											sSplitOSD_sDomainUser=Split(OSD_sDomainUser , "\")
											oEnvironment.Item("DomainAdmin")=sSplitOSD_sDomainUser(1)
											WriteLog "TS variable DomainAdmin=" & oEnvironment.Item("DomainAdmin")
										Else
											OSD_oTaskSequence ("OSDJoinAccount") = OSD_sDomainUser
											WriteLog "TS variable OSDJoinAccount = " & OSD_sDomainUser
										End If
										
									Else
									
										If OSD_IsMDT="YES" Then
										
											oEnvironment.Item("DomainAdmin") =OSD_sDomainUser
											
										Else
											If Len(Trim(OSD_sDomain)) > 0 Then 
												OSD_oTaskSequence ("OSDJoinAccount") = OSD_sDomain & "\" & OSD_sDomainUser
											ElseIf Len(OSD_oTaskSequence ("OSDJoinDomainName")) > 0 Then
												If InStr (1, OSD_oTaskSequence ("OSDJoinDomainName"), ".", 1) > 0 Then
													OSD_oTaskSequence ("OSDJoinAccount") = OSD_sDomainUser & "@" & OSD_oTaskSequence ("OSDJoinDomainName")
												Else
													OSD_oTaskSequence ("OSDJoinAccount") = OSD_oTaskSequence ("OSDJoinDomainName") & "\" & OSD_sDomainUser
												End If
											Else
												OSD_oTaskSequence ("OSDJoinAccount") = OSD_sDomainUser
											End If
										End If
										
									End If
								
							End If
						
						'set OSDJoinPassword
							If Len(Trim(OSD_sDomainUserPassword)) > 0  Then
								If OSD_IsMDT="YES" Then
									oEnvironment.Item("DomainAdminPassword")=OSD_sDomainUserPassword
									WriteLog "TS variable Domain Admin Password: ****************"
								Else
									OSD_oTaskSequence ("OSDJoinPassword") = OSD_sDomainUserPassword
									WriteLog "TS variable OSDJoinPassword = " & OSD_sDomainUserPassword
								End if
							End If
				End If
	Else
			WriteLog "Cannot define OSD variables due to not running from Task Sequence"
	
	End If
	
	WriteLog "*** setOSDVariables() - finished"

End Sub


Sub Rename_Hostname(sCurrentName, sNewName)

	WriteLog "*** Rename_Hostname() - started..."
	WriteLog "Computername will be renamed to new hostname:" & sNewName
	
	Dim objComputer, Return, OSD_objWMIService 
	Return = 0
	
	Set OSD_objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

	For Each objComputer In OSD_objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem where Name ='" & sCurrentName & "'")
		
		Return = objComputer.rename(sNewName)
        	
        If Return <> 0 Then
           	WriteLog "Rename_Hostname() - rename failed. Error = " & Err.Number &  vbTab & Err.Description
        Else
           	WriteLog "Rename_Hostname() - rename succeeded. a reboot will be needed to apply the new hostname"
        End If
	Next
	
	WriteLog "*** Rename_Hostname() - finished"

End Sub

Sub setRegionalOptions
	
	WriteLog "*** Sub setRegionalOptions() - started..."

	Dim sFile, oFile, oFSO, i

    
    If Len(Trim(OSD_sCountry)) > 0 Or Len(Trim(OSD_sSystemLocale)) > 0 Or Len(Trim(OSD_sInputLocale)) > 0 Or Len(Trim(OSD_sUserLocale)) > 0 Or Len(Trim(OSD_sUILanguage)) > 0 Then
    
	    	
	    	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
		
			' Create XML output File
		    
		    sFile = "RegionalOptions.xml"
		    
		    WriteLog "Creating output Regional Settings file " & sFile
		    
		    sFile = OSD_sLogPath & "RegionalOptions.xml"
		    Set oFile = oFSO.CreateTextFile(sFile, True)
	
        	WriteLog "Creating Regional Settings XML file"
	
	        oFile.WriteLine "<gs:GlobalizationServices xmlns:gs=""urn:longhornGlobalizationUnattend"">"
	        oFile.WriteLine ""
	        oFile.WriteLine "    <!-- user list -->"
	        oFile.WriteLine "    <gs:UserList>"
	        oFile.WriteLine "        <gs:User UserID=""Current"" CopySettingsToDefaultUserAcct=""true"" CopySettingsToSystemAcct=""true""/>"
	        oFile.WriteLine "    </gs:UserList>"
	        oFile.WriteLine ""
			
			If Len(Trim(OSD_sCountry)) > 0 Then
				
				oFile.WriteLine "    <!-- location -->"
	            oFile.WriteLine "    <gs:LocationPreferences>"
	        	oFile.WriteLine "        <gs:GeoID Value="& chr(34) & OSD_sCountry & chr(34) & "/>"
	        	oFile.WriteLine "    </gs:LocationPreferences>"
	        	
			End If
			
	        If Len(Trim(OSD_sSystemLocale)) > 0 Then
	          	
	            oFile.WriteLine "    <!-- system locale -->"
	            oFile.WriteLine "    <gs:SystemLocale Name=" & chr(34) & OSD_sSystemLocale & chr(34) & "/>"
	            oFile.WriteLine ""
	          
	        	
	        End If
			
	        If Len(Trim(OSD_sInputLocale)) > 0 Then
            	    isKeybUS=False
            	    isKeybGB=False
            	    isKeybFR=False
	            oFile.WriteLine "    <!-- input preferences -->"
	            oFile.WriteLine "    <gs:InputPreferences>"
	            
	            arrInputLocales = Split(OSD_sInputLocale , ";")
	            strLocaleLines = ""
	            If IsArray(arrInputLocales) Then
	                For i = 0 to UBound(arrInputLocales)    	
					If arrInputLocales(i)="0409:00000409" Then
		   				isKeybUS=True
					End If
					
					If arrInputLocales(i)="0809:00000809" Then
		   				isKeybGB=True
					End If
						If UCase(arrInputLocales(i))=UCase("040c:0000040c") Then
			   				isKeybFR=True
						End If
	                    If i = 0 Then
	                        oFile.WriteLine "        <gs:InputLanguageID Action=" & chr(34) & "add" & chr(34) & " ID=" & chr(34) & arrInputLocales(OSD_i) & chr(34) & " Default=" & chr(34) & "true" & chr(34) & "/>"
	                    Else
	                        oFile.WriteLine "        <gs:InputLanguageID Action=" & chr(34) & "add" & chr(34) & " ID=" & chr(34) & Trim(arrInputLocales(OSD_i)) & chr(34) & "/>"
	                    End If
		             Next
	            Else
	                oFile.WriteLine "        <gs:InputLanguageID Action=" & chr(34) & "add" & chr(34) & " ID=" & chr(34) & OSD_sInputLocale & chr(34) & " Default=" & chr(34) & "true" & chr(34) & "/>"            
	            End If
	            
	            oFile.WriteLine "    </gs:InputPreferences>"
	            oFile.WriteLine ""
	        End If
	
	        If Len(Trim(OSD_sUserLocale)) > 0 Then
	        
					oFile.WriteLine "	<!-- user locale -->"
	            	oFile.WriteLine "    <gs:UserLocale>"
	            	oFile.WriteLine "        <gs:Locale Name=" & chr(34) & OSD_sUserLocale & chr(34) & " SetAsCurrent=""true"" ResetAllSettings=""true""/>"
	            	oFile.WriteLine "    </gs:UserLocale>"
			End If
			
			If Len(Trim(OSD_sUILanguage)) > 0 Then
			
					oFile.WriteLine "	<!-- Display Language -->"
	            	oFile.WriteLine "    <gs:MUILanguagePreferences>"
	            	oFile.WriteLine "        <gs:MUILanguage Value=" & chr(34) & OSD_sUILanguage & chr(34) & "/>"
	            	oFile.WriteLine "    </gs:MUILanguagePreferences>"
			End If
			
			oFile.WriteLine "</gs:GlobalizationServices>"
	        oFile.WriteLine ""
	        oFile.Close
	        
	        
		'apply RegionalOptions.xml
		WriteLog "Applying regional settings..."
        OSD_sCmd = "cmd /c control.exe intl.cpl,,/f:" & chr(34) & sFile & Chr(34)
        
        WriteLog "About to run command: " & OSD_sCmd
        ExecuteCommand OSD_sCmd
        WScript.Sleep 3000
        
        'check if en-US or en-GB or fr-FR is not selected then remove it from the OS
        If isKeybFR=False Then
        	sFile=OSD_sLogPath & "RemoveFRKeyb.xml"
			iret=PrepXML(sFile,"040c:0000040c")
			OSD_sCmd = "cmd /c control.exe intl.cpl,,/f:" & chr(34) & sFile & Chr(34)
    		WriteLog "About to run command: " & OSD_sCmd
    		ExecuteCommand OSD_sCmd
    		WScript.Sleep 1000
		End If
        
        If isKeybUS=False Then
        	sFile=OSD_sLogPath & "RemoveUSKeyb.xml"
			iret=PrepXML(sFile,"0409:00000409")
			OSD_sCmd = "cmd /c control.exe intl.cpl,,/f:" & chr(34) & sFile & Chr(34)
    		WriteLog "About to run command: " & OSD_sCmd
    		ExecuteCommand OSD_sCmd
    		WScript.Sleep 1000
		End If
		
		If isKeybGB=False Then
        	sFile=OSD_sLogPath & "RemoveGBKeyb.xml"
			iret=PrepXML(sFile,"0809:00000809")
			OSD_sCmd = "cmd /c control.exe intl.cpl,,/f:" & chr(34) & sFile & Chr(34)
    		WriteLog "About to run command: " & OSD_sCmd
    		ExecuteCommand OSD_sCmd
    		WScript.Sleep 1000
		End If
	        
   	Else
		WriteLog "No Regional settings to apply."
   	End If
		
	If 	Len(Trim(OSD_sTimeZone)) > 0 Then
	

	    'set the Timezone using builtin Windows 7 tool tzutil.exe
	    WriteLog "Applying Time Zone setting..."
	    OSD_sCmd = "Tzutil.exe /s" & " " & chr(34) & OSD_sTimeZone & chr(34)
	    WriteLog "About to run command: " & OSD_sCmd
	    
	    ExecuteCommand OSD_sCmd
	    
	Else
		WriteLog "No TimeZone setting to apply."
		
	End If
	
		
	WriteLog "*** Sub setRegionalOptions() - finished."

End Sub

Function PrepXML (sXML, sKeyboard )
	
	Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
	' Create XML output File
	WriteLog "Creating output Regional Settings file " & sXML
	Set oFile = oFSO.CreateTextFile(sXML, True)
    oFile.WriteLine "<gs:GlobalizationServices xmlns:gs=""urn:longhornGlobalizationUnattend"">"
    oFile.WriteLine ""
    oFile.WriteLine "    <!-- user list -->"
    oFile.WriteLine "    <gs:UserList>"
    oFile.WriteLine "        <gs:User UserID=""Current""/>"
    oFile.WriteLine "    </gs:UserList>"
	oFile.WriteLine "    <gs:InputPreferences>"
    oFile.WriteLine "        <gs:InputLanguageID Action=" & chr(34) & "remove" & chr(34) & " ID=" & chr(34) & sKeyboard & chr(34) & "/>"            
    oFile.WriteLine "    </gs:InputPreferences>"
	oFile.WriteLine "</gs:GlobalizationServices>"
    oFile.Close
    
End function 


Sub ExecuteCommand (scommand)
	WriteLog "*** Executing command: " & sCommand
	OSD_oShell.Run scommand,1,True
	'WScript.Sleep 1000
End Sub

Sub setCurrentKeyboard(sKeybID)

	On Error Resume Next
	
	sLocale=hex(GetLocale)
	while len(sLocale) < 4
		sLocale = "0" & sLocale
	wend

	while len(sKeybID) < 8
		sKeybID = "0" & sKeybID
	wend

	sID=sLocale & ":" & sKeybID
	
	If UCase(OSD_RootDrv)="X:" Then
   		'In WinPE session
		OSD_sCmd = "wpeutil.exe SetKeyboardLayout " & sID
		oSh.run OSD_sCmd, 0 , True
	Else 
		'In customer Windows OS session
		sKeybXML=OSD_oShell.ExpandEnvironmentStrings("%TEMP%") &"\" & "keyb.xml"
		If OSD_objFSO.FileExists(sKeybXML) Then
		 	OSD_objFSO.DeleteFile sKeybXML, True
		End If
		CreateKeyb_XML sKeybXML, sID	
		OSD_sCmd="control intl.cpl,, /f:" & Chr(34) & sKeybXML & Chr(34)
		oSh.run OSD_sCmd, 0 ,true
		WScript.Sleep 3000
		
	End If
	
end Sub

Sub CreateKeyb_XML(sXML, sKeyboard)

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
	  
	Set objRoot = xmlDoc.createElement("gs:GlobalizationServices")  
	xmlDoc.appendChild objRoot
	objRoot.setAttribute "xmlns:gs","urn:longhornGlobalizationUnattend"
	objRoot.appendChild xmlDoc.createComment("User List ")
	
	Set objRecord = xmlDoc.createElement("gs:UserList") 
	objRoot.appendChild objRecord 
	  
	Set objName = xmlDoc.createElement("gs:User")
	objRecord.appendChild objName
	objName.setAttribute "UserID","Current"

	objRoot.appendChild xmlDoc.createComment("input preferences ")
	
	Set objRecord = xmlDoc.createElement("gs:InputPreferences") 
	objRoot.appendChild objRecord 
	  
	Set objName = xmlDoc.createElement("gs:InputLanguageID")  
	objRecord.appendChild objName
	objName.setAttribute "Action","add"
	objName.setAttribute "ID",sKeyboard
	objName.setAttribute "Default","true"
	
	xmlDoc.Save sXML

End Sub

Sub Create_OSDProfile_Using_DefaultGateway_Settings

' update OSDProfile.ini from values specified in OSDSettings.ini file when DefaultGateway Mode is selected
		
	If OSD_bDefaultGateway=True And OSD_skipWizard =True Then
		
	' Create the OSDProfile.ini
		WriteLog "------ Creating OSDProfile.ini file using the following settings ------"
		
		Set oFile = OSD_objFSO.OpenTextFile(OSD_sOSDProfileIniFile, OSD_ForAppending, True)
		
		oFile.WriteLine "[Main]"
		oFile.WriteLine "Selected=" & OSD_sSite
		WriteLog "Selected Site=" & OSD_sSite
		oFile.WriteLine ""
		oFile.WriteLine "[" & OSD_sSite & "]"
		oFile.WriteLine "SystemLocale=" & OSD_sSystemLocale
		WriteLog "SystemLocale=" & OSD_sSystemLocale
		
		oFile.WriteLine "UserLocale=" & OSD_sUserLocale
		WriteLog "UserLocale=" & OSD_sUserLocale
		
		oFile.WriteLine "InputLocale=" & OSD_sInputLocale
		WriteLog "InputLocale=" & OSD_sInputLocale
		
		oFile.WriteLine "UILanguage=" & OSD_sUILanguage
		WriteLog "UserLocale=" & OSD_sUserLocale
		
		oFile.WriteLine "Country=" & OSD_sCountry
		WriteLog "Country=" & OSD_sCountry
		
		oFile.WriteLine "Computername=" & OSD_sNewComputername
		WriteLog "Computername=" & OSD_sNewComputername
		
		oFile.WriteLine "TimeZone=" & OSD_sTimeZone
		WriteLog "TimeZone=" & OSD_sTimeZone
		
		oFile.WriteLine "Join=" & OSD_sJoin
		WriteLog "Join=" & OSD_sJoin
		If UCase(OSD_sJoin)=UCase("Workgroup") Then
			oFile.WriteLine "Workgroup=" & OSD_sWorkgroup
			WriteLog "Workgroup=" & OSD_sWorkgroup
			
		ElseIf UCase(OSD_sJoin)=UCase("Domain") Then
		
			oFile.WriteLine "Domain=" & OSD_sDomain
			WriteLog "Domain=" & OSD_sDomain
					
			oFile.WriteLine "DomainOU=" & OSD_sDomainOU
			WriteLog "DomainOU=" & OSD_sDomainOU
			
			oFile.WriteLine "DomainUser=" & OSD_sDomainUser
			WriteLog "DomainUser=" & OSD_sDomainUser
			
			oFile.WriteLine "DomainUserPassword=" & OSD_sDomainUserPassword
			WriteLog "DomainUserPassword=" & OSD_sDomainUserPassword
		End If
		
			'add command line if exist to OSDProfile.ini
			If UBound(sArrayApps)> 0 Then
					
					For i=0 To UBound(sArrayApps)
						 objTextFile.WriteLine "Application00" & i+1 &"=" & sArrayApps(i)
						 
						 WriteLog "Application00" & i+1 &"="  & sArrayApps(i)
					Next
			End If
					
		oFile.Close
	End If
End Sub


Function GetComputerName
	Dim  sDeskLapSrv, OSD_strComputer, i, temparray, sHostname
	OSD_strComputer = "."
	
	'get the default ComputerNaming section from OSDSettings.ini
	OSD_sDesktop=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Desktop"))
	OSD_sLaptop=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Laptop"))
	OSD_Server=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Server"))
	OSD_sConstruct=Trim(ReadIni(OSD_sSettingsIniFile, "ComputerNaming","Construct"))
	
	'get service tag, asset tag
	Set OSD_objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & OSD_strComputer & "\root\cimv2")
	Set colSMBIOS = OSD_objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
	For Each objSMBIOS in colSMBIOS
		If Len(Trim(objSMBIOS.SerialNumber)) > 0 Then
    		OSD_sServiceTag=objSMBIOS.SerialNumber
    	End If
    	
    	If Len(Trim(objSMBIOS.SMBIOSAssetTag)) > 0 Then 
			OSD_sAssetTag=objSMBIOS.SMBIOSAssetTag
		End If
	Next	
	
	'set the computer type
	Select Case UCase(GetType)
		 Case ucase("Desktop")
			 sDeskLapSrv=OSD_sDesktop
		 Case Ucase("Laptop")
		 	 sDeskLapSrv=OSD_sLaptop
		 Case Ucase("Server")
		 	sDeskLapSrv=OSD_Server
		 Case Else
		 	 sDeskLapSrv=""
	End Select
	
	'set Computername like "Construct" field in OSDSettings.ini
	sHostname=""
	temparray=Split(OSD_sConstruct,"+")
	
	For i=0 To UBound(temparray)

		If UCase(temparray(i)) = UCase("<Prefix>") Then
			sHostname=sHostname + OSD_sPrefix
		End If
		
		If UCase(temparray(i)) = UCase("<Desktop_Laptop_Server>") Then
			sHostname= sHostname + sDeskLapSrv
		End If
		
		If UCase(temparray(i)) = UCase("<Suffix>") Then
			sHostname=sHostname + OSD_sSuffix 
		End If
		
		If UCase(temparray(i)) = UCase("<Service_Tag>") Then
			sHostname=sHostname + UCase(OSD_sServiceTag)
		End If
		
		If UCase(temparray(i)) = UCase("<Asset_Tag>") Then
			sHostname=sHostname + UCase(OSD_sAssetTag)
		End If
	Next
	
	If InStr (1,UCase(sHostname), "SERVICE_TAG",1) > 0 Then 	
		sHostname=Replace(UCase(sHostname),"SERVICE_TAG",OSD_sServiceTag)
		
	End If
				
	If InStr (1,UCase(sHostname), "ASSET_TAG",1) > 0 Then	
		sHostname=Replace(UCase(sHostname),"ASSET_TAG",OSD_sAssetTag) 
	End If
	
	'set the computername
	GetComputerName=sHostname

End Function 

'//---------------------------------------------------------------------------
	'//  Function:	GetAssetInfo()
	'//  Purpose:	Get asset information using WMI
	'//---------------------------------------------------------------------------
Function GetAssetInfo

	 	Dim bFoundBattery, bFoundAC
		Dim objResults, objInstance
		Dim i, scmd
		Dim bisX64

		writelog "Getting asset info"

		' Get the SMBIOS asset tag from the Win32_SystemEnclosure class
		OSD_strComputer ="."
		Set OSD_objWMIService = Nothing
		Set OSD_objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & OSD_strComputer & "\root\cimv2")

		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
		bIsLaptop = false
		bIsDesktop = false
		bIsServer = false
		For each objInstance in objResults

			If objInstance.ChassisTypes(0) = 12 or objInstance.ChassisTypes(0) = 21 then
				' Ignore docking stations
			Else

				If Len(Trim(objInstance.SMBIOSAssetTag)) > 0 Then 
					sAssetTag = Trim(objInstance.SMBIOSAssetTag)
				End If
				
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

		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_BIOS")
		For each objInstance in objResults
	         ' Get the serial number

			If Len(Trim(objInstance.SerialNumber)) > 0 then
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
		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_Processor")
		For each objInstance in objResults

			' Get the processor speed

			If not IsNull(objInstance.MaxClockSpeed) then
				sProcessorSpeed = Trim(objInstance.MaxClockSpeed)
			End if


			' Determine if the machine supports SLAT (only supported with Windows 8)

			On error resume next
			bSupportsSLAT = objInstance.SecondLevelAddressTranslationExtensions
			On Error Goto 0
		'msgbox "bSupportsSLAT =" & bSupportsSLAT 	

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

		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
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

		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct")
		For each objInstance in objResults

			If not IsNull(objInstance.UUID) then
				sUUID = Trim(objInstance.UUID)
			End if

		Next
		If sUUID = "" then
			writelog "Unable to determine UUID via WMI."
		End if


		' Get the product from the Win32_BaseBoard class

		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_BaseBoard")
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
			sReturn = OSD_oShell.Run(scmd , 0, True)
			On Error Goto 0
			If sReturn = 0 Then 
					'read the value of reg HKLM\System\CurrentControlSet\Control\PEFirmwareType
					
					sregPEFirmwareType="HKLM\System\CurrentControlSet\Control\PEFirmwareType"
					
					writelog "determine UEFI or BIOS mode by reading registry key HKLM\System\CurrentControlSet\Control\PEFirmwareType."
					
					sFirmware = OSD_oShell.RegRead(sregPEFirmwareType)
					writelog "PEFirmwareType value is:" & sFirmware
					
					If sFirmware="0x1" Then bIsUEFI = False
					If sFirmware="0x2" Then bIsUEFI = True

			Else
					writelog "Unable to determine if running UEFI via registry." & ". Error description :" & Err.Description
				
			End If
			
		Else
		
			'On error resume next
			scmd="cmd /c BCDEDIT.exe /ENUM >" & Chr(34) & oEnv("tmp") & "\BcdeditEnum.txt" & Chr(34)
			sReturn = OSD_oShell.Run(scmd , 0, True)
			On Error Goto 0
			If sReturn = 0 Then 
				If OSD_objFSO.FileExists (oEnv("tmp") & "\BcdeditEnum.txt") Then
					
					Set ini = OSD_objFSO.OpenTextFile( oEnv("tmp") & "\BcdeditEnum.txt", 1, False)
					Do While (not ini.AtEndOfStream)
						line = ini.ReadLine
						line = Trim(line)
						
						If InStr(1, UCase (line), ucase("Path"),1) > o And InStr(1, UCase (line), ucase("\EFI\Microsoft\Boot\bootmgfw.efi"),1) > o Then 
							bIsUEFI = True
							
							Exit Do 
						End If
					Loop
					ini.Close
				'	writelog "deleting temp file " &  oEnv("tmp") & "\BcdeditEnum.txt"
					OSD_objFSO.DeleteFile oEnv("tmp") & "\BcdeditEnum.txt", True
					
				Else
					writelog "NOT found " &  oEnv("tmp") & "\BcdeditEnum.txt"
					writelog "Unable to determine if running UEFI via command BCDEDIT.exe /ENUM."
				End If
			Else
				writelog "Unable to determine if running UEFI via command BCDEDIT.exe /ENUM." & ". Error description :" & Err.Description
				
			End If 
		End If
			
		' See if we are running on battery

		If oEnv("SystemDrive") = "X:" and OSD_objFSO.FileExists("X:\Windows\Inf\Battery.inf") then
			
			' Load the battery driver

			OSD_oShell.Run "drvload X:\Windows\Inf\Battery.inf", 0, true
			
		End if

		bFoundAC = False
		bFoundBattery = False
		Set objResults = OSD_objWMIService.ExecQuery("Select * from Win32_Battery")
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
		
'		If bSupportsSLAT or sSupportsSLAT = "" Then
'			sSupportsSLAT = ConvertBooleanToString(bSupportsSLAT)
'			WriteLog "sSupportsSLAT = " &  ConvertBooleanToString(bSupportsSLAT)
'		Else
'			writelog "Property SupportsSLAT = " & sSupportsSLAT
'		End if

		writelog "Finished getting asset info"

		GetAssetInfo = OSD_Success

End Function

	
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


Function GetNetworkDetails(OSD_dicIPAddresses, OSD_dicDefaultGateway)

	Dim OSD_iRetVal, objNetworkAdapters, objAdapter, sElement, sTmp1, arrTmp

	OSD_iRetVal = OSD_Failure

	' Get a list of IP-enabled adapters

	Set objNetworkAdapters = OSD_objWMIService.ExecQuery("select * from Win32_NetworkAdapterConfiguration where IPEnabled = 1")

	For Each objAdapter In objNetworkAdapters

		WriteLog "Checking network adapter: " & objAdapter.Caption
		
		' Get the IP addresses
		If not (IsNull(objAdapter.IPAddress)) then
			for each sElement in objAdapter.IPAddress
				if sElement = "0.0.0.0" or Left(sElement, 7) = "169.254" or sElement = "" then
					WriteLog "Ignoring IP Address " & sElement
				else
					If not OSD_dicIPAddresses.Exists(sElement) then
						OSD_dicIPAddresses.Add sElement, ""
					End if
					WriteLog "IP Address = " & sElement
				end if
			next
		End if

		' Get the default gateway values
		If not (IsNull(objAdapter.DefaultIPGateway)) then
			for each sElement in objAdapter.DefaultIPGateway
				if sElement <> "" then
					If not OSD_dicDefaultGateway.Exists(sElement) then
						OSD_dicDefaultGateway.Add sElement, ""
					End if
					WriteLog "Default Gateway = " & sElement
					OSD_iRetVal=OSD_Success
				end if
			next
		End If
		
	next  

	GetNetworkDetails = OSD_iRetVal
	WriteLog "Finished retrieving network info via WMI"

End Function


'		Function IsLaptop( myComputer )
		' This Function checks if a computer has a battery pack.
		' One can assume that a computer with a battery pack is a laptop.
		'
		' Argument:
		' myComputer   [string] name of the computer to check,
		'                       or "." for the local computer
		' Return value:
		' True if a battery is detected, otherwise False
		'
'		    On Error Resume Next
'		    Set objWMIService = GetObject( "winmgmts://" & myComputer & "/root/cimv2" )
'		    Set colItems = objWMIService.ExecQuery( "Select * from Win32_Battery", , 48 )
'		    IsLaptop = False
'		    For Each objItem in colItems
'		        IsLaptop = True
'		    Next
'		    If Err Then Err.Clear
'		    On Error Goto 0
'		End Function 


Function GetType
	GetType=""
	
	
	If sIsLaptop= True Then
		GetType="LAPTOP"
	ElseIf sIsDesktop= True Then
		GetType="DESKTOP"
	ElseIf sIsServer= True Then
		GetType="SERVER"
	End If
	
	'get Asset and service tag
	
	   OSD_sServiceTag=sSerialNumber
	   OSD_sAssetTag=sAssetTag
	
	
End Function

Function WriteLog(sLogMsg)

		Dim sTime, sDate, sTempMsg, OSD_oLog, oConsole

		On Error Resume Next		
		' Suppress messages containing password
		If not OSD_bDebug then
			If Instr(1, sLogMsg, "password", 1) > 0 then
				sLogMsg = "<Message containing password has been suppressed>"
			End if
		End if

		' Populate the variables to log
			sTempMsg = "[LOG]:" & " Date: " & Date & " " & Time  & " Message: " &  sLogMsg
			
		' If debug, echo the message
		If OSD_bDebug then
			Set oConsole = OSD_objFSO.GetStandardStream(1) 
			oConsole.WriteLine sLogMsg
		End if

		' Create the log entry
		Err.Clear
		Set OSD_oLog = OSD_objFSO.OpenTextFile(OSD_LogFile, OSD_ForAppending, True)
		
		If Err then
			Err.Clear
			Exit Function
		End if
		OSD_oLog.WriteLine sTempMsg
		OSD_oLog.Close
		Err.Clear

End Function


Function ReadIni(file, section, item)

		Dim line, equalpos, leftstring, ini

		ReadIni = ""
		file = Trim(file)
		item = Trim(item)

		On Error Resume Next
		Set ini = OSD_objFSO.OpenTextFile( file, 1, False)
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

	temp_ini = OSD_sScriptDir & OSD_objFSO.GetTempName

	Set read_ini = OSD_objFSO.OpenTextFile( file, 1, True, TristateFalse )
	Set write_ini = OSD_objFSO.CreateTextFile( temp_ini, False)

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
	If OSD_objFSO.FileExists(file) then
		OSD_objFSO.DeleteFile file, True
	End if
	OSD_objFSO.CopyFile temp_ini, file, true
	OSD_objFSO.DeleteFile temp_ini, True

End Sub
	
Sub UpdateUnattendxml

	Dim brRemoveCredentials
	On Error Resume Next
		
	brRemoveCredentials=False


		'Update Unattend.xml		
		
		WriteLog "unattendXml file was specified as: " & sUnattendXml 
		OSD_iRetVal=OSD_Success
		
		' Load the XML file
		
		If sUnattendXml <> "" And OSD_objFSO.FileExists(sUnattendXml) Then
			WriteLog "Copying existing unattend.xml for backup."
			'OSD_objFSO.GetFile(sUnattendXml).Attributes = 0
			OSD_objFSO.CopyFile sUnattendXml, OSD_sScriptDir & "unattend_BEFORE_Merge.xml", True
			WriteLog "Copied " & sUnattendXml & " to " & OSD_sScriptDir & "unattend_BEFORE_Merge.xml"
			
			Set oUnattendXml = CreateObject("Microsoft.XMLDOM")
			oUnattendXml.async = False
			oUnattendXml.load sUnattendXml
			
			WriteLog "Loaded " & sUnattendXml
			OSD_bMerge=True
			
		else
			WriteLog "File " & sUnattendXml & " does not exist. Unattend.xml is not updated."
			sUnattendXml=""
			OSD_iRetVal=OSD_Failure
			OSD_bMerge=False
			
		End if
		
		' Add new entries if not exist already in unattend.xml
		
		If OSD_bMerge Then
				WriteLog "Updating " & sUnattendXml &  " with settings found on " & OSD_sOSDProfileIniFile
				
				'oobeSystem pass
				Set oOOBE_template = CreateObject("Microsoft.XMLDOM")
				oOOBE_template.async = False
				oOOBE_template.load sOOBEXml_template
			
				doMergeXML oUnattendXml, oOOBE_template
				oUnattendXml.Save sUnattendXml
				wscript.sleep 2000

				If Len(Trim(OSD_sInputLocale)) >0 Then
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-International-Core']/InputLocale"
					UpdateXML sNodePath,OSD_sInputLocale,oUnattendXml
					bChanged = True
				End If
				
				If Len(Trim(OSD_sSystemLocale)) >0 Then
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-International-Core']/SystemLocale"
					UpdateXML sNodePath,OSD_sSystemLocale,oUnattendXml
					bChanged = True
				End If
				
				If Len(Trim(OSD_sUILanguage)) >0 Then
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-International-Core']/UILanguage"
					UpdateXML sNodePath,OSD_sUILanguage,oUnattendXml
					bChanged = True
				Else
					OSD_sUILanguage="EN-US"
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-International-Core']/UILanguage"
					UpdateXML sNodePath,OSD_sUILanguage,oUnattendXml
					bChanged = True
				End If
				
				If Len(Trim(OSD_sUserLocale)) >0 Then
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-International-Core']/UserLocale"
					UpdateXML sNodePath,OSD_sUserLocale,oUnattendXml
					bChanged = True
				End If
				
				'TimeZone-oobeSystem pass
				If Len(Trim(OSD_sTimeZone)) >0 Then
					sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/TimeZone"
					UpdateXML sNodePath,OSD_sTimeZone,oUnattendXml
					bChanged = True
				End If
				
				'If customer is asking to let Windows to manage display settings then just uncomment the Display section below:				
				'Remove Display node if exists - Specialize pass			
				'Section begin
					Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-Shell-Setup']/Display")
						
						'oNode.ParentNode.removeChild oNode
						WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-Shell-Setup']/Display*** is removed"
						bChanged = True
						
					If not (oNode is Nothing) Then
						oNode.ParentNode.removeChild oNode
						WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-Shell-Setup']/Display*** is removed"
						bChanged = True
					Else
							WriteLog "The Display section entry not found"
					End If
						
				'Section end			
		
				'ComputerName- Specialize pass
				If Len(Trim(OSD_sNewComputername)) > 0 Then
					If Not boJoinDomtemplateLoaded Then
						
						Set oJoinDom_template = CreateObject("Microsoft.XMLDOM")
						oJoinDom_template.async = False 
						oJoinDom_template.load sDomainXML_template
						
						doMergeXML oUnattendXml, oJoinDom_template
						oUnattendXml.Save sUnattendXml
						boJoinDomtemplateLoaded=True
						
					End If
					sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-Shell-Setup']/ComputerName"
					UpdateXML sNodePath,OSD_sNewComputername,oUnattendXml
					bChanged = True
				End If
				
				'TimeZone- Specialize pass
				If Len(Trim(OSD_sTimeZone)) >0 Then
					sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-Shell-Setup']/TimeZone"
					UpdateXML sNodePath,OSD_sTimeZone,oUnattendXml
					bChanged = True
				End If
				
				'Merge DomainXML_template with Unattend.xml
			
				If Len(Trim(OSD_sDomain)) > 0 Or Len(Trim(OSD_sDomainOU)) > 0 Or Len(Trim(OSD_sDomainUser)) > 0 Or Len(Trim(OSD_sDomainUserPassword)) > 0 Then
					sJoinSystem="DOMAIN"
				ElseIf Len(OSD_sWorkgroup) > 0 Then
					sJoinSystem="JoinWorkgroup"
				End If
				
				If Len(Trim(OSD_sIPAddress)) >0 Then
					If Not bNetworkTemplateLoaded Then
						Set oNetworkInterface_template = CreateObject("Microsoft.XMLDOM")
						oNetworkInterface_template.async = False 
						oNetworkInterface_template.load sNetworkInterfaceXML_template
						
						doMergeXML oUnattendXml, oNetworkInterface_template
						oUnattendXml.Save sUnattendXml
						
						bNetworkTemplateLoaded=true
					End If
					sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-TCPIP']/Interfaces/Interface/UnicastIpAddresses/IpAddress"	
					UpdateXML sNodePath,OSD_sIPAddress,oUnattendXml
					bChanged = True
				End If
				
				If Len(Trim(OSD_sGatewayAddress)) >0 Then
					If Not bNetworkTemplateLoaded Then
						Set oNetworkInterface_template = CreateObject("Microsoft.XMLDOM")
						oNetworkInterface_template.async = False 
						oNetworkInterface_template.load sNetworkInterfaceXML_template
						
						doMergeXML oUnattendXml, oNetworkInterface_template
						oUnattendXml.Save sUnattendXml
						
						bNetworkTemplateLoaded=true
					End If
					sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-TCPIP']/Interfaces/Interface/Routes/Route/NextHopAddress"	
					UpdateXML sNodePath,OSD_sGatewayAddress,oUnattendXml
					bChanged = True
				End If
				
				If Len(Trim(OSD_sPrimDNS)) > 0 Then
					If Not bNetworkTemplateLoaded Then
						Set oNetworkInterface_template = CreateObject("Microsoft.XMLDOM")
						oNetworkInterface_template.async = False 
						oNetworkInterface_template.load sNetworkInterfaceXML_template
						
						doMergeXML oUnattendXml, oNetworkInterface_template
						oUnattendXml.Save sUnattendXml
						
						bNetworkTemplateLoaded=true
					End If
					sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-DNS-Client']/Interfaces/Interface/DNSServerSearchOrder/IpAddress"	
					UpdateXML sNodePath,OSD_sPrimDNS,oUnattendXml
					bChanged = True
				End If
				
				If sJoinSystem="DOMAIN" Then
				
										If Not boJoinDomtemplateLoaded Then
											Set oJoinDom_template = CreateObject("Microsoft.XMLDOM")
											oJoinDom_template.async = False 
											oJoinDom_template.load sDomainXML_template
											
											doMergeXML oUnattendXml, oJoinDom_template
											oUnattendXml.Save sUnattendXml
											boJoinDomtemplateLoaded=True
										End If
										
										'remove JoinWorkgroup node if exists
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinWorkgroup")
										
										If not (oNode is Nothing) Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinWorkgroup*** is removed"
											bChanged = True
										End If
									
										'Credentials section
										If Len(Trim(OSD_sDomain))>0 Then
											sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Domain"	
											UpdateXML sNodePath,OSD_sDomain,oUnattendXml
										Else
											'Remove Domain Credentials if empty
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Domain")
											If not (oNode is Nothing) Then
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Domain*** is removed"
												bChanged = True
											End If
											
										End If
										
										If Len(Trim(OSD_sDomainUser))>0 Then
												sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Username"	
												UpdateXML sNodePath,OSD_sDomainUser,oUnattendXml
										Else
											'Remove Username Credentials if empty
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Username")
											If oNode is Nothing Then
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Username*** is removed"
												bChanged = True
												brRemoveCredentials=True
											End If
											
										End If
										
										If Len(Trim(OSD_sDomainUserPassword))>0 Then
												sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Password"	
												UpdateXML sNodePath,OSD_sDomainUserPassword,oUnattendXml
										Else
											'Remove Password Credentials if empty
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Password")
											If oNode is Nothing Then
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Password*** is removed"
												bChanged = True
												brRemoveCredentials=True
											End If
											
										End If
										
										If brRemoveCredentials=True Then
											'Remove Credentials section if empty
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials")
										
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials*** is removed"
												bChanged = True
										End if
																
										'JoinDomain
										If Len(Trim(OSD_sDomain))>0 Then
											sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinDomain"	
											UpdateXML sNodePath,OSD_sDomain,oUnattendXml
										End If
										
										'MachineObjectOU
										If Len(Trim(OSD_sDomainOU)) >0 Then
											sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/MachineObjectOU"
											UpdateXML sNodePath,OSD_sDomainOU,oUnattendXml
										Else
											'Remove MachineObjectOU if exists
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/MachineObjectOU")
											If not (oNode is Nothing) Then
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/MachineObjectOU*** is removed"
												bChanged = True
											End If
										End If
							
				ElseIf UCase(sJoinSystem) = UCase("JoinWorkgroup") Then   ' join workgroup
								
									'update JoinWorkgroup node
										If Len(Trim(OSD_sWorkgroup))= 0 Then
											OSD_sWorkgroup="Workgroup"
										End If
										sNodePath="//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinWorkgroup"	
										UpdateXML sNodePath,OSD_sWorkgroup,oUnattendXml
										bChanged = True
									
									'removeJoinDomain
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinDomain")
										If not (oNode is Nothing) Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/JoinDomain*** is removed"
											bChanged = True
										End If
									
									'Remove MachineObjectOU if exists
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/MachineObjectOU")
										If not (oNode is Nothing) Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/MachineObjectOU*** is removed"
											bChanged = True
										End If
								
									'Remove Domain Credentials 
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Domain")
										If not (oNode is Nothing) Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Domain*** is removed"
											bChanged = True
										End If
					
									'Remove Username Credentials
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Username")
										If oNode is Nothing Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Username*** is removed"
											bChanged = True
											brRemoveCredentials=True
										End If
										
									'Remove Password Credentials
										Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Password")
										If oNode is Nothing Then
											oNode.ParentNode.removeChild oNode
											WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials/Password*** is removed"
											bChanged = True
											brRemoveCredentials=True
										End If
										
									'Remove Credentials section
										If brRemoveCredentials=True Then 
											Set oNode = oUnattendXml.selectSingleNode("//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials")
										
												oNode.ParentNode.removeChild oNode
												WriteLog "The node:***//settings[@pass='specialize']/component[@name='Microsoft-Windows-UnattendedJoin']/Identification/Credentials*** is removed"
												bChanged = True
										End If
				End If 
				
							
					'update unattend for autologon to process the applications installation phase
					If sAppsInstall="YES" Then
					
						sRunPost="YES"
					 	
					 		If Not sAutologonSetInUnattendXML="YES" then
									WriteLog "Updating unattend.xml file with Autologon settings for Phase 2 of OSDCustomizer and applications installation..."
					 	
									sAutologonUser=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUser"))
									sAutologonPwd=Trim(ReadIni(OSD_sSettingsIniFile, "Applications","AutologonUserPassword"))
									sEncrypAutologonPwd =Trim(ReadIni(OSD_sSettingsIniFile, "Applications","EncryptedAutologonUserPassword"))
										
										If Len(Trim(sAutologonUser)) > 0 Then
											sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/AutoLogon/Username"
											UpdateXML sNodePath,sAutologonUser,oUnattendXml
											bChanged = True
										End If
										
										If Len(Trim(sAutologonPwd)) > 0 Then
											sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/AutoLogon/Password/Value"
											UpdateXML sNodePath,sAutologonPwd,oUnattendXml
											bChanged = True
											
										ElseIf Len(Trim(sEncrypAutologonPwd)) > 0 Then
											sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/AutoLogon/Password/Value"
											UpdateXML sNodePath,sEncrypAutologonPwd,oUnattendXml
											sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/AutoLogon/Password/PlainText"
											UpdateXML sNodePath,"False",oUnattendXml
											bChanged = True
										End If
									
							End If
					 		
					 		' set LogonCount to 99
					 		sNodePath="//settings[@pass='oobeSystem']/component[@name='Microsoft-Windows-Shell-Setup']/AutoLogon/LogonCount"
							UpdateXML sNodePath,"99",oUnattendXml
							WriteLog "Set to 99 the value of the node:***//" & sNodePath & "***"
							bChanged = True			
					End If

				
				OSD_iRetVal=OSD_Success
				
		End If  ' end of If OSD_bMerge Then
			
		' Rewrite the Unattend.xml if it has been changed
		
		If bChanged then
			oUnattendXml.Save sUnattendXml
			WriteLog "Rewrote " & sUnattendXml & " with changes"
			WriteLog "Unattend.xml update is completed."
			OSD_iRetVal=OSD_Success
		End if
			
End Sub

Function UpdateXML(oNode,value,oXmlDoc)

	WriteLog "Searching " & oNode
	
	Set oFound = oXmlDoc.selectSingleNode(oNode)
			
	If oFound is nothing then
		
		WriteLog "Child " & oNode & " not found. Unattend.xml is not updated"
				
	Else
		' Found, process the children
	
		WriteLog "Child " & oNode & " already exists, updating its value"
		oXmlDoc.selectSingleNode(oNode).text=value
		bChanged=True 
	End if

End Function

Function doMergeXML(oDestination, oSource)
On Error Resume Next


	Dim oChild
	Dim sPath
	Dim i
	Dim oFound
	Dim o
	Dim doAdd

	For each oChild in oSource.childNodes
	
		If oChild.nodeTypeString = "element" Then
		
			' Build a query to find the node in the destination
			sPath = oChild.nodeName
			
			If not (oChild.Attributes is nothing) then
				If oChild.Attributes.length > 0 then
					
					sPath = sPath & "["
					
					For i = 0 to oChild.Attributes.length - 1
						If UCase(Left(oChild.Attributes.item(i).name, 6)) = "XMLNS:" Then
						
							' Ignore the namespaces when searching, assuming that the namespaces are defined at a higher level
						Else
							sPath = sPath & "@" & oChild.Attributes.item(i).name & "='" & oChild.Attributes.item(i).value & "' and "
						End if
					Next
					sPath = Left(sPath, Len(sPath) - 5) & "]"				
				End if
			End If
			
			WriteLog "Searching for " & sPath
 
			Set oFound = oDestination.selectSingleNode(sPath)
						
			If oFound is nothing then
				WriteLog "Adding new child " & sPath
				oDestination.appendChild oChild
			
			Else
				' Found, process the children
				WriteLog "Child " & sPath & " already exists, checking its children"
				doMergeXML oFound, oChild
			End if
		End if
	Next

End Function



Sub CreateGeoIDXML(GeoID)

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
	  
	Set objRoot = xmlDoc.createElement("gs:GlobalizationServices")  
	xmlDoc.appendChild objRoot
	objRoot.setAttribute "xmlns:gs","urn:longhornGlobalizationUnattend"
	objRoot.appendChild xmlDoc.createComment("User List ")
		
	Set objRecord = xmlDoc.createElement("gs:UserList") 
	objRoot.appendChild objRecord 
	  
	Set objName = xmlDoc.createElement("gs:User")  
	objRecord.appendChild objName
	objName.setAttribute "UserID","Current"
	objName.setAttribute "CopySettingsToDefaultUserAcct","true"
	
	objRoot.appendChild xmlDoc.createComment("location ")
	
	Set objRecord = xmlDoc.createElement("gs:LocationPreferences") 
	objRoot.appendChild objRecord 
	  
	Set objName = xmlDoc.createElement("gs:GeoID")  
	objRecord.appendChild objName
	objName.setAttribute "Value",GeoID
	
	xmlDoc.Save sGeoIDXml
	WriteLog "GeoID.xml file created using the value for Location GeoID=" & GeoID

End Sub

Function getUILanguage
	Dim vFound

	vFound=False

	strComputer = "."
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\"&_ 
	    strComputer & "\root\default:StdRegProv")
	strKeyPath = "SYSTEM\CurrentControlSet\Control\MUI\UILanguages"
	objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
	
	WriteLog "Subkeys under HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\MUI\UILanguages:"
	
	For Each subkey In arrSubKeys
	    If IsNull( arrSubKeys ) = False Then
	    	
	    	'WriteLog subkey
	    	If UCase(Trim(OSD_sUILanguage))= UCase(subkey) Then
	    		
	    		vFound=True
	    		Exit For
	    	Else
	    			
	    		vFound=False
	    	End If
	    	
	    End if
	Next
	
	If vFound=True Then
		 getUILanguage= True
		 WriteLog "Selected UI language " & OSD_sUILanguage & " is already installed. no need to re-apply regional settings."
	    		
	Else
		getUILanguage= False
		WriteLog "Selected UI language " & OSD_sUILanguage & " is not yet installed. Regional settings will be re-applied after the language pack installation."
	    	
	End If
		
End Function

Sub SetStartOSDCustomizer

		Dim oLink

		' Set up to automatically run me, using the appropriate method

		If OSD_objFSO.FileExists(oEnv("SystemRoot") & "\Explorer.exe") then

			' If shortcut for OSDCUSTOMIZER.VBS  doesn't exist then create a new shortcut.

			If not OSD_objFSO.FileExists(OSD_oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk") then

			    ' Not Server Core, create a shortcut
			    writelog "Creating startup folder item to run OSDCustomizer once the shell is loaded."

			    Set oLink = OSD_oShell.CreateShortcut(OSD_oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk")
			    oLink.TargetPath = "wscript.exe"
			    
			    oLink.Arguments = Chr(34) & OSD_sScriptDir & "OSDCUSTOMIZER.VBS" & Chr(34) & " /POST:" & Chr(34) & OSD_sOSDProfileIniFile & Chr(34)
			    
			    oLink.Save

			    writelog "Shortcut """ & OSD_oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" created."

			Else
			     writelog "Shortcut """ & OSD_oShell.SpecialFolders("AllUsersStartup") & "\Dell_OSDCustomizer.lnk"" already exists."
			End If

		Else

			' Server core or "hidden shell", register a "Run" item

			writelog "Creating Run registry key to run the OSDCustomizer for the next reboot."

			On Error Resume Next
			OSD_oShell.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\DELL_OSDCustomizer", "wscript.exe " & Chr(34) & OSD_sScriptDir & "OSDCUSTOMIZER.VBS" & Chr(34) & " /POST:" & Chr(34) & OSD_sOSDProfileIniFile & Chr(34), "REG_SZ"
			Writelog "Wrote Run registry key"
			
			On Error Goto 0

			' Allow execution to continue (assuming new Run item won't actually be run yet)

		End if
	
End Sub

Function GetTheParent(DriveSpec)
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
   GetTheParent = objFSO.GetParentFolderName(Drivespec)
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
		Set oExec = OSD_oShell.Exec(sCmd)
		Do While oExec.Status = 0

			' Sleep
			WScript.Sleep 500
			
			' See if it is time for a heartbeat
			If iHeartbeat > 0 and DateDiff("n", lastHeartbeat, Now) > iHeartbeat then
				iMinutes = DateDiff("n", lastStart, Now)
								
				writeLog "Heartbeat: command has been running for " & iMinutes & " minutes (process ID " & oExec.ProcessID & ")"
				
				If iMinutes > 60 Then
				
					writeLog "ERROR: command has been running for more than 60 minutes... So assuming that there is a problem, the installation is aborted." 
					
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
