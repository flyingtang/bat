On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive",,48)

Dim objItem 'as Win32_DiskDrive
For Each objItem in colItems
	WScript.Echo "Availability: " & objItem.Availability
	WScript.Echo "BytesPerSector: " & objItem.BytesPerSector
	WScript.Echo "Capabilities: " & objItem.Capabilities
	WScript.Echo "CapabilityDescriptions: " & objItem.CapabilityDescriptions
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "CompressionMethod: " & objItem.CompressionMethod
	WScript.Echo "ConfigManagerErrorCode: " & objItem.ConfigManagerErrorCode
	WScript.Echo "ConfigManagerUserConfig: " & objItem.ConfigManagerUserConfig
	WScript.Echo "CreationClassName: " & objItem.CreationClassName
	WScript.Echo "DefaultBlockSize: " & objItem.DefaultBlockSize
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DeviceID: " & objItem.DeviceID
	WScript.Echo "ErrorCleared: " & objItem.ErrorCleared
	WScript.Echo "ErrorDescription: " & objItem.ErrorDescription
	WScript.Echo "ErrorMethodology: " & objItem.ErrorMethodology
	WScript.Echo "FirmwareRevision: " & objItem.FirmwareRevision
	WScript.Echo "Index: " & objItem.Index
	WScript.Echo "InstallDate: " & objItem.InstallDate
	WScript.Echo "InterfaceType: " & objItem.InterfaceType
	WScript.Echo "LastErrorCode: " & objItem.LastErrorCode
	WScript.Echo "Manufacturer: " & objItem.Manufacturer
	WScript.Echo "MaxBlockSize: " & objItem.MaxBlockSize
	WScript.Echo "MaxMediaSize: " & objItem.MaxMediaSize
	WScript.Echo "MediaLoaded: " & objItem.MediaLoaded
	WScript.Echo "MediaType: " & objItem.MediaType
	WScript.Echo "MinBlockSize: " & objItem.MinBlockSize
	WScript.Echo "Model: " & objItem.Model
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "NeedsCleaning: " & objItem.NeedsCleaning
	WScript.Echo "NumberOfMediaSupported: " & objItem.NumberOfMediaSupported
	WScript.Echo "Partitions: " & objItem.Partitions
	WScript.Echo "PNPDeviceID: " & objItem.PNPDeviceID
	WScript.Echo "PowerManagementCapabilities: " & objItem.PowerManagementCapabilities
	WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
	WScript.Echo "SCSIBus: " & objItem.SCSIBus
	WScript.Echo "SCSILogicalUnit: " & objItem.SCSILogicalUnit
	WScript.Echo "SCSIPort: " & objItem.SCSIPort
	WScript.Echo "SCSITargetId: " & objItem.SCSITargetId
	WScript.Echo "SectorsPerTrack: " & objItem.SectorsPerTrack
	WScript.Echo "SerialNumber: " & objItem.SerialNumber
	WScript.Echo "Signature: " & objItem.Signature
	WScript.Echo "Size: " & objItem.Size
	WScript.Echo "Status: " & objItem.Status
	WScript.Echo "StatusInfo: " & objItem.StatusInfo
	WScript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
	WScript.Echo "SystemName: " & objItem.SystemName
	WScript.Echo "TotalCylinders: " & objItem.TotalCylinders
	WScript.Echo "TotalHeads: " & objItem.TotalHeads
	WScript.Echo "TotalSectors: " & objItem.TotalSectors
	WScript.Echo "TotalTracks: " & objItem.TotalTracks
	WScript.Echo "TracksPerCylinder: " & objItem.TracksPerCylinder
	WScript.Echo ""
Next
