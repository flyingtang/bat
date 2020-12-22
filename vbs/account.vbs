
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Account",,48)


For Each objItem in colItems
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "Domain: " & objItem.Domain
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "LocalAccount: " & objItem.LocalAccount
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "SID: " & objItem.SID
	WScript.Echo "SIDType: " & objItem.SIDType
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo ""
Next
WScript.Echo "========================================="

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_AccountSID",,48)


For Each objItem in colItems
	WScript.Echo "Element: " & objItem.Element
	WScript.Echo "Setting: " & objItem.Setting
	WScript.Echo ""
Next

WScript.Echo "========================================="

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Group",,48)


For Each objItem in colItems
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "Domain: " & objItem.Domain
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "LocalAccount: " & objItem.LocalAccount
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "SID: " & objItem.SID
	WScript.Echo "SIDType: " & objItem.SIDType
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo ""
Next

WScript.Echo "========================================="

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_GroupInDomain",,48)


For Each objItem in colItems
	WScript.Echo "GroupComponent: " & objItem.GroupComponent
	WScript.Echo "PartComponent: " & objItem.PartComponent
	WScript.Echo ""
Next


WScript.Echo "========================================="
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_GroupUser",,48)


For Each objItem in colItems
	WScript.Echo "GroupComponent: " & objItem.GroupComponent
	WScript.Echo "PartComponent: " & objItem.PartComponent
	WScript.Echo ""
Next
