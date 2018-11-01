'
'Sample VBScript: Set a single BIOS Setting on the local computer.  Use this script if you have registered a supervisor password.
'
'   command line: cscript.exe ImportBiosConfig.vbs /file:ConfigFile.ini /Password:scrambled SupervisorPassword


'declare application name
Dim objWMIService, objItem, colItems, strComputer, strInParamValue, strReturn, strItem, strStatus, strFileName, strSupervisorPassword, strParameter, colParamItems
Dim oArgCollection, oConfigFile, oPassword, oline, ofile, oFso, oParce, oSetting, oValue, oState, oTestpath, oShell, oParceTwo, colCurrentSetItems, oItem, ooItem 

Set oShell = CreateObject("WScript.Shell")
Set oFSO=CreateObject("Scripting.FileSystemObject")

'create Object to open the procedure file
strFileName = "procedures.vbs"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFileName, 1)	'1 - for reading
If objFile is Nothing Then
    WScript.Echo "You can not run this application without file " & strFileName 
    WScript.Quit
End If
Execute objFile.ReadAll()

'check input parameters
Set oArgCollection = WScript.Arguments.Named
If oArgCollection.Exists("file") Then
	oConfigFile = oArgCollection.item("file")
	If oConfigFile <> "" Then
		If instr(oConfigFile,":")=0 Then oConfigFile = oShell.CurrentDirectory & "\" & oConfigFile 
		If not oFSO.fileExists (oConfigFile) Then 
			Wscript.echo "unable to find BIOS in the specified path " & oConfigFile & ", Aborting !!!"	
			WScript.Quit		
		End If	
	Else
		Wscript.echo "unable to find BIOS config File, Aborting !!!"	
		WScript.Quit
	End If	
Else
	Wscript.echo "BIOS config File is Missing, Aborting !!!"	
	WScript.Quit
End If	

If oArgCollection.Exists("password") Then strSupervisorPassword = oArgCollection.item("password")

'connect to WMI
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

If Err.Number <> 0 Then
	WScript.Echo "Unable to connect to WMI service: " & Hex(Err.Number) & "."
	WScript.Quit
End If
On Error Goto 0


'Check if the supervisor password is registered
If strSupervisorPassword <> "" Then
	strReturn = IsSupervisorPasswordRegistered(objWMIService)
	If strReturn <> 0 Then
		WScript.Echo "You can not run this application if the supervisor password is not registered."
		WScript.Quit
	End If

	'Authenticate with Supervisor privilege
	strParameter = "Start," + strSupervisorPassword + ";"
	strReturn = SetConfigurationMode(objWMIService, strParameter)
	If strReturn <> 0 Then
		WScript.Echo "Supervisor password authentication failed. Error:" & GetErrMsg(Hex(strReturn))
		WScript.Quit
	Else
		WScript.Echo "Supervisor password successfull authenticated."
	End If
End If

'executes a WQL query
Set colItems = objWMIService.ExecQuery("Select * from BiosSetting where InstanceName='ACPI\\PNP0C14\\0_0'")
Set colCurrentSetItems = objWMIService.ExecQuery("Select * from QueryBiosSettings")

'Reading BIOS File
Set oFile = oFSO.OpenTextFile(oConfigFile)

Do Until oFile.AtEndOfStream
    oLine = oFile.ReadLine
	oParce = split ( oline, "=")
	oSetting = Trim(oParce(0))
	oParceTwo = split(oParce(1),",")
	oValue = Trim(oParceTwo(1))
	oState = Trim (oParceTwo(0))
	'WScript.Echo "[DEBUG] Reading from file " & oSetting
	
	If oState = "RW" or oState = "WO" Then
		'Check if setting needs to be changed
		For Each objItem in colCurrentSetItems 
			If Len(objItem.CurrentSetting) > 0 Then
			'	return value contains two elements, each seperated by comma. e.g: "WakeUpOnLAN,Enable"
				strSetting = ObjItem.CurrentSetting
				'Find the corresponding setting
				If Instr(strSetting,oSetting)>0 Then
					'Wscript.Echo "[DEBUG] checking " & sTrSetting & " against " & oSetting 
					strValue = Mid(strSetting, InStr(strSetting, ",") + 1, 256)
					strValue = Trim(split(strValue,",")(1))
					'Check if requested setting is the same that inplace setting
					If strValue = oValue Then
						Wscript.Echo "Setting " & oSetting & " is already configured with value " & oValue & ",skipping !!!"
						WScript.Echo "====================================="
						Exit For
					Else
						'Wscript.Echo "[DEBUG] Setting " & oSetting & " must be changed from value " & strValue & " to " & oValue
						'If they are different, setting will be updated
						strInParamValue = oSetting + "," + oValue + ";"
						Wscript.echo "Setting up: " & strInParamValue
						'set single Bios setting
						For Each oItem in colItems
							'execute the method and obtain the return status
							oItem.SetBiosSetting strInParamValue, strReturn
						Next	
						WScript.Echo "status: " & GetErrMsg(Hex(strReturn))
						WScript.Echo "====================================="
					End If
				End If
			End If
		Next
	End If
Loop
objFile.Close


''deauthenticate from supervisor mode
If strSupervisorPassword <> "" Then
	strParameter = "End," + strSupervisorPassword + ";"
	strReturn = SetConfigurationMode(objWMIService, strParameter)
	If strReturn <> 0 Then
		WScript.Echo "Supervisor password deauthentication failed. Error:" & GetErrMsg(Hex(strReturn))
		WScript.Quit
	Else
		WScript.Echo "Supervisor password successfully deauthenticated."
	End If
End IF

WScript.Quit


''convert an error code to a string
Function GetErrMsg(err)
    Dim strMsg
    Select Case err
        Case "0"
           strMsg = "The operation was successful."
        Case "8004100C"
           strMsg = "Feature or operation is not supported."
        Case "80041008"
           strMsg = "One of the parameters to the call is not correct."
        Case "80041003"
           strMsg = "Write Protect error"
        Case "80041062"
           strMsg = "Operation failed because the client did not have the necessary security privilege."
        Case "80045001"
           strMsg = "Authentication failure."
        Case "80045002"
           strMsg = "Password not registered."
        Case Else
           strMsg = "error code " + err
    End Select

    GetErrMsg = strMsg
End Function

