On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_JobObjectStatus",,48)

Dim objItem 'as Win32_JobObjectStatus
For Each objItem in colItems
	WScript.Echo "AdditionalDescription: " & objItem.AdditionalDescription
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "Operation: " & objItem.Operation
	WScript.Echo "ParameterInfo: " & objItem.ParameterInfo
	WScript.Echo "ProviderName: " & objItem.ProviderName
	WScript.Echo "StatusCode: " & objItem.StatusCode
	WScript.Echo "Win32ErrorCode: " & objItem.Win32ErrorCode
	WScript.Echo ""
Next
