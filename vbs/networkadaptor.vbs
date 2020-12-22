On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)

Dim objItem 'as Win32_NetworkAdapter
For Each objItem in colItems
	WScript.Echo "AdapterType: " & objItem.AdapterType
	WScript.Echo "AdapterTypeId: " & objItem.AdapterTypeId
	WScript.Echo "AutoSense: " & objItem.AutoSense
	WScript.Echo "Availability: " & objItem.Availability
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
	WScript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
	WScript.Echo "CreationClassName: " & objItem.CreationClassName
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DeviceID: " & objItem.DeviceID
	WScript.Echo "ErrorCleared: " & objItem.ErrorCleared
	WScript.Echo "ErrorDescription: " & objItem.ErrorDescription
	WScript.Echo "GUID: " & objItem.GUID
	WScript.Echo "Index: " & objItem.Index
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "Installed: " & objItem.Installed
	WScript.Echo "InterfaceIndex: " & objItem.InterfaceIndex
	WScript.Echo "LastErrorCode: " & objItem.LastErrorCode
	WScript.Echo "MACAddress: " & objItem.MACAddress
	WScript.Echo "Manufacturer: " & objItem.Manufacturer
	WScript.Echo "MaxNumberControlled: " & objItem.MaxNumberControlled
	WScript.Echo "MaxSpeed: " & objItem.MaxSpeed
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "NetConnectionID: " & objItem.NetConnectionID
	WScript.Echo "NetConnectionStatus: " & objItem.NetConnectionStatus
	WScript.Echo "NetEnabled: " & objItem.NetEnabled
	WScript.Echo "NetworkAddresses: " & objItem.NetworkAddresses
	WScript.Echo "PermanentAddress: " & objItem.PermanentAddress
	WScript.Echo "PhysicalAdapter: " & objItem.PhysicalAdapter
	WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
	WScript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
	WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
	WScript.Echo "ProductName: " & objItem.ProductName
	WScript.Echo "ServiceName: " & objItem.ServiceName
	WScript.Echo "Speed: " & objItem.Speed
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo "StatusInfo: " & objItem.StatusInfo
	WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
	WScript.Echo "SystemName: " & objItem.SystemName
	WScript.Echo "TimeOfLastReset: " & objItem.TimeOfLastReset
	WScript.Echo ""
Next
