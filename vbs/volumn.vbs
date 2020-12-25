On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume",,48)

Dim objItem 'as Win32_Volume
For Each objItem in colItems
	WScript.Echo "Access: " & objItem.Access
	WScript.Echo "Automount: " & objItem.Automount
	WScript.Echo "Availability: " & objItem.Availability
	WScript.Echo "BlockSize: " & objItem.BlockSize
	WScript.Echo "BootVolume: " & objItem.BootVolume
	WScript.Echo "Capacity: " & objItem.Capacity
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "Compressed: " & objItem.Compressed
	WScript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
	WScript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
	WScript.Echo "CreationClassName: " & objItem.CreationClassName
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DeviceID: " & objItem.DeviceID
	WScript.Echo "DirtyBitSet: " & objItem.DirtyBitSet
	WScript.Echo "DriveLetter: " & objItem.DriveLetter
	WScript.Echo "DriveType: " & objItem.DriveType
	WScript.Echo "ErrorCleared: " & objItem.ErrorCleared
	WScript.Echo "ErrorDescription: " & objItem.ErrorDescription
	WScript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
	WScript.Echo "FileSystem: " & objItem.FileSystem
	WScript.Echo "FreeSpace: " & objItem.FreeSpace
	WScript.Echo "IndexingEnabled: " & objItem.IndexingEnabled
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "Label: " & objItem.Label
	WScript.Echo "LastErrorCode: " & objItem.LastErrorCode
	WScript.Echo "MaximumFileNameLength: " & objItem.MaximumFileNameLength
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "NumberOfBlocks: " & objItem.NumberOfBlocks
	WScript.Echo "PageFilePresent: " & objItem.PageFilePresent
	WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
	WScript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
	WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
	WScript.Echo "Purpose: " & objItem.Purpose
	WScript.Echo "QuotasEnabled: " & objItem.QuotasEnabled
	WScript.Echo "QuotasIncomplete: " & objItem.QuotasIncomplete
	WScript.Echo "QuotasRebuilding: " & objItem.QuotasRebuilding
	WScript.Echo "SerialNumber: " & objItem.SerialNumber
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo "StatusInfo: " & objItem.StatusInfo
	WScript.Echo "SupportsDiskQuotas: " & objItem.SupportsDiskQuotas
	WScript.Echo "SupportsFileBasedCompression: " & objItem.SupportsFileBasedCompression
	WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
	WScript.Echo "SystemName: " & objItem.SystemName
	WScript.Echo "SystemVolume: " & objItem.SystemVolume
	WScript.Echo ""
Next
