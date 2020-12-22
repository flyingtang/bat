On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskPartition",,48)

Dim objItem 'as Win32_DiskPartition
For Each objItem in colItems
	WScript.Echo "Access: " & objItem.Access
	WScript.Echo "Availability: " & objItem.Availability
	WScript.Echo "BlockSize: " & objItem.BlockSize
	WScript.Echo "Bootable: " & objItem.Bootable
	WScript.Echo "BootPartition: " & objItem.BootPartition
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
	WScript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
	WScript.Echo "CreationClassName: " & objItem.CreationClassName
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DeviceID: " & objItem.DeviceID
	WScript.Echo "DiskIndex: " & objItem.DiskIndex
	WScript.Echo "ErrorCleared: " & objItem.ErrorCleared
	WScript.Echo "ErrorDescription: " & objItem.ErrorDescription
	WScript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
	WScript.Echo "HiddenSectors: " & objItem.HiddenSectors
	WScript.Echo "Index: " & objItem.Index
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "LastErrorCode: " & objItem.LastErrorCode
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
	WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
	WScript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
	WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
	WScript.Echo "PrimaryPartition: " & objItem.PrimaryPartition
	WScript.Echo "Purpose: " & objItem.Purpose
	WScript.Echo "RewritePartition: " & objItem.RewritePartition
	WScript.Echo "Size: " & objItem.Size
	WScript.Echo "StartingOffset: " & objItem.StartingOffset
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo "StatusInfo: " & objItem.StatusInfo
	WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
	WScript.Echo "SystemName: " & objItem.SystemName
	WScript.Echo "Type: " & objItem.Type
	WScript.Echo ""
Next
