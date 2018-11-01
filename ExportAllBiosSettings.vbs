'
'Sample VBScript: List all BIOS settings on the local computer
'
'   command line: cscript.exe ListAllBiosSettings.vbs


Dim objWMIService, objItem, colItems, strSetting, strItem, strValue
Dim oComputerModel,oShell, oCurrentDirectory, oFSO,oFile

'Get Current directory
Set oShell = CreateObject("WScript.Shell")
oCurrentDirectory = oShell.CurrentDirectory

'Get Computer Model
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem in colItems
    oComputerModel = objItem.Model
next

'Create Text file
Set oFSO=CreateObject("Scripting.FileSystemObject")
Set oFile = oFSO.CreateTextFile(oCurrentDirectory & "\" & oComputerModel & ".ini",True)

'Connect to WMI
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\.\root\WMI")

If Err.Number <> 0 Then
	WScript.Echo "Unable to connect to WMI service: " & Hex(Err.Number) & "."
	WScript.Quit
End If
On error Goto 0

'executes a WQL query
Set colItems = objWMIService.ExecQuery("Select * from QueryBiosSettings")

For Each objItem in colItems 
	If Len(objItem.CurrentSetting) > 0 Then
		'return value contains two elements, each seperated by comma. e.g: "WakeUpOnLAN,Enable"
		strSetting = ObjItem.CurrentSetting
		strItem = Left(strSetting, InStr(strSetting, ",") - 1)
		strValue = Mid(strSetting, InStr(strSetting, ",") + 1, 256)
		WScript.Echo strItem + " = " + strValue
		oFile.Write strItem + " = " + strValue & vbCrLf
	End If
Next

'Close file
oFile.Close

WScript.Quit
