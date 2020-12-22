On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Service",,48)

Dim objItem 'as Win32_Service
For Each objItem in colItems
	WScript.Echo "AcceptPause: " & objItem.AcceptPause
	WScript.Echo "AcceptStop: " & objItem.AcceptStop
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "CheckPoint: " & objItem.CheckPoint
	WScript.Echo "CreationClassName: " & objItem.CreationClassName
	WScript.Echo "DelayedAutoStart: " & objItem.DelayedAutoStart
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DesktopInteract: " & objItem.DesktopInteract
	WScript.Echo "DisplayName: " & objItem.DisplayName
	WScript.Echo "ErrorControl: " & objItem.ErrorControl
	WScript.Echo "ExitCode: " & objItem.ExitCode
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "PathName: " & objItem.PathName
	WScript.Echo "ProcessId: " & objItem.ProcessId
	WScript.Echo "ServiceSpecificExitCode: " & objItem.ServiceSpecificExitCode
	WScript.Echo "ServiceType: " & objItem.ServiceType
	WScript.Echo "Started: " & objItem.Started
	WScript.Echo "StartMode: " & objItem.StartMode
	WScript.Echo "StartName: " & objItem.StartName
	WScript.Echo "State: " & objItem.State
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
	WScript.Echo "SystemName: " & objItem.SystemName
	WScript.Echo "TagId: " & objItem.TagId
	WScript.Echo "WaitHint: " & objItem.WaitHint
	WScript.Echo ""
Next
