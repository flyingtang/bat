On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)

Dim objItem 'as Win32_OperatingSystem

For Each objItem in colItems
	' WScript.Echo "BootDevice: " & objItem.BootDevice
	' WScript.Echo "BuildNumber: " & objItem.BuildNumber
	' WScript.Echo "BuildType: " & objItem.BuildType
	WScript.Echo "Caption=" & objItem.Caption
	WScript.Echo "CSName=" & objItem.CSName
	' WScript.Echo "CodeSet: " & objItem.CodeSet
	' WScript.Echo "CountryCode: " & objItem.CountryCode
	' WScript.Echo "CreationClassName: " & objItem.CreationClassName
	' WScript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
	' WScript.Echo "CSDVersion: " & objItem.CSDVersion
	' WScript.Echo "CurrentTimeZone: " & objItem.CurrentTimeZone
	' WScript.Echo "DataExecutionPrevention_32BitApplications: " & objItem.DataExecutionPrevention_32BitApplications
	' WScript.Echo "DataExecutionPrevention_Available: " & objItem.DataExecutionPrevention_Available
	' WScript.Echo "DataExecutionPrevention_Drivers: " & objItem.DataExecutionPrevention_Drivers
	' WScript.Echo "DataExecutionPrevention_SupportPolicy: " & objItem.DataExecutionPrevention_SupportPolicy
	' WScript.Echo "Debug: " & objItem.Debug
	' WScript.Echo "Description: " & objItem.Description
	' WScript.Echo "Distributed: " & objItem.Distributed
	' WScript.Echo "EncryptionLevel: " & objItem.EncryptionLevel
	' WScript.Echo "ForegroundApplicationBoost: " & objItem.ForegroundApplicationBoost
	' WScript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
	' WScript.Echo "FreeSpaceInPagingFiles: " & objItem.FreeSpaceInPagingFiles
	' WScript.Echo "FreeVirtualMemory: " & objItem.FreeVirtualMemory
	' WScript.Echo "InstallDate: " & objItem.InstallDate
	' WScript.Echo "LargeSystemCache: " & objItem.LargeSystemCache
	' WScript.Echo "LastBootUpTime: " & objItem.LastBootUpTime
	' WScript.Echo "LocalDateTime: " & objItem.LocalDateTime
	' WScript.Echo "Locale: " & objItem.Locale
	' WScript.Echo "Manufacturer: " & objItem.Manufacturer
	' WScript.Echo "MaxNumberOfProcesses: " & objItem.MaxNumberOfProcesses
	' WScript.Echo "MaxProcessMemorySize: " & objItem.MaxProcessMemorySize
	' WScript.Echo "MUILanguages: " & objItem.MUILanguages
	' WScript.Echo "Name: " & objItem.Name
	' WScript.Echo "NumberOfLicensedUsers: " & objItem.NumberOfLicensedUsers
	' WScript.Echo "NumberOfProcesses: " & objItem.NumberOfProcesses
	' WScript.Echo "NumberOfUsers: " & objItem.NumberOfUsers
	' WScript.Echo "OperatingSystemSKU: " & objItem.OperatingSystemSKU
	' WScript.Echo "Organization: " & objItem.Organization
	' WScript.Echo "OSArchitecture: " & objItem.OSArchitecture
	' WScript.Echo "OSLanguage: " & objItem.OSLanguage
	' WScript.Echo "OSProductSuite: " & objItem.OSProductSuite
	' WScript.Echo "OSType: " & objItem.OSType
	' WScript.Echo "OtherTypeDescription: " & objItem.OtherTypeDescription
	' WScript.Echo "PAEEnabled: " & objItem.PAEEnabled
	' WScript.Echo "PlusProductID: " & objItem.PlusProductID
	' WScript.Echo "PlusVersionNumber: " & objItem.PlusVersionNumber
	' WScript.Echo "PortableOperatingSystem: " & objItem.PortableOperatingSystem
	' WScript.Echo "Primary: " & objItem.Primary
	' WScript.Echo "ProductType: " & objItem.ProductType
	' WScript.Echo "RegisteredUser: " & objItem.RegisteredUser
	' WScript.Echo "SerialNumber: " & objItem.SerialNumber
	' WScript.Echo "ServicePackMajorVersion: " & objItem.ServicePackMajorVersion
	' WScript.Echo "ServicePackMinorVersion: " & objItem.ServicePackMinorVersion
	' WScript.Echo "SizeStoredInPagingFiles: " & objItem.SizeStoredInPagingFiles
	' WScript.Echo "Status: " & objItem.Status
	' WScript.Echo "SuiteMask: " & objItem.SuiteMask
	' WScript.Echo "SystemDevice: " & objItem.SystemDevice
	' WScript.Echo "SystemDirectory: " & objItem.SystemDirectory
	' WScript.Echo "SystemDrive: " & objItem.SystemDrive
	' WScript.Echo "TotalSwapSpaceSize: " & objItem.TotalSwapSpaceSize
	' WScript.Echo "TotalVirtualMemorySize: " & objItem.TotalVirtualMemorySize
	' WScript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
	' WScript.Echo "Version: " & objItem.Version
	' WScript.Echo "WindowsDirectory: " & objItem.WindowsDirectory
	' WScript.Echo ""
 	' 20201221122955.492346+480
	 WScript.Echo "runtime=" &  GetRuntimeSecond(objItem.LastBootUpTime)
Next


function GetRuntimeSecond(strLastBootUpTime) 
	strOldDate=formatWindowsDate(strLastBootUpTime)
	datOld = CDate(strOldDate)
	datNow = Date()&" "&Time() 
	intDiffSecond=DateDiff("s", datOld, datNow)
	GetRuntimeSecond = intDiffSecond
end function

function formatWindowsDate (strLastBootUpTime)
    tmp=strLastBootUpTime
    count = 4
    y=Left(tmp, count)
    tmp=Mid(tmp, count+1)

    count = 2
    m=Left(tmp, count)
    tmp=Mid(tmp, count+1)

    d=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    h=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    mi=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    s=Left(tmp, 2)
	tmp=Mid(tmp, count+1)
	
	formatWindowsDate= m&"/"&d&"/"&y&" "&h&":"&mi&":"&s
end function
