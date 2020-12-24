
' GBK编码
' 注意先后顺序,系统信息必须先获取
' 2003 函数调用会有问题
' On Error Resume Next

' 第一个返回值：物理内存总量 B
arrSystemInfo = GetSystemInfo()

'包含总内存虚拟内存
GetCPUInfo()

' 进程信息
GetProcessInfo()

' 磁盘情况，剩余空间
GetDiskInfo()

' 网络收发速率
GetNetworkAdaptorInfo() 

' 服务 需要统计
GetServiceInfo()

' SCHTASKS 获取调度信息
GetSchTasksInfo()

function GetSchTasksInfo()
	WScript.Echo "=====TODO GetSchTasksInfo=========="
end function

function GetServiceInfo()
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

end function


function GetNetworkAdaptorInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_Tcpip_NetworkAdapter",,48)

	Dim objItem 'as Win32_PerfFormattedData_Tcpip_NetworkAdapter
	For Each objItem in colItems
		WScript.Echo "BytesReceivedPersec: " & objItem.BytesReceivedPersec
		WScript.Echo "BytesSentPersec: " & objItem.BytesSentPersec
		WScript.Echo "BytesTotalPersec: " & objItem.BytesTotalPersec
		WScript.Echo "Caption: " & objItem.Caption
		WScript.Echo "CurrentBandwidth: " & objItem.CurrentBandwidth
		WScript.Echo "Description: " & objItem.Description
		WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
		WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
		WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
		WScript.Echo "Name: " & objItem.Name
		WScript.Echo "OffloadedConnections: " & objItem.OffloadedConnections
		WScript.Echo "OutputQueueLength: " & objItem.OutputQueueLength
		WScript.Echo "PacketsOutboundDiscarded: " & objItem.PacketsOutboundDiscarded
		WScript.Echo "PacketsOutboundErrors: " & objItem.PacketsOutboundErrors
		WScript.Echo "PacketsPersec: " & objItem.PacketsPersec
		WScript.Echo "PacketsReceivedDiscarded: " & objItem.PacketsReceivedDiscarded
		WScript.Echo "PacketsReceivedErrors: " & objItem.PacketsReceivedErrors
		WScript.Echo "PacketsReceivedNonUnicastPersec: " & objItem.PacketsReceivedNonUnicastPersec
		WScript.Echo "PacketsReceivedPersec: " & objItem.PacketsReceivedPersec
		WScript.Echo "PacketsReceivedUnicastPersec: " & objItem.PacketsReceivedUnicastPersec
		WScript.Echo "PacketsReceivedUnknown: " & objItem.PacketsReceivedUnknown
		WScript.Echo "PacketsSentNonUnicastPersec: " & objItem.PacketsSentNonUnicastPersec
		WScript.Echo "PacketsSentPersec: " & objItem.PacketsSentPersec
		WScript.Echo "PacketsSentUnicastPersec: " & objItem.PacketsSentUnicastPersec
		WScript.Echo "TCPActiveRSCConnections: " & objItem.TCPActiveRSCConnections
		WScript.Echo "TCPRSCAveragePacketSize: " & objItem.TCPRSCAveragePacketSize
		WScript.Echo "TCPRSCCoalescedPacketsPersec: " & objItem.TCPRSCCoalescedPacketsPersec
		WScript.Echo "TCPRSCExceptionsPersec: " & objItem.TCPRSCExceptionsPersec
		WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
		WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
		WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
		WScript.Echo ""
	Next

end function

' 还缺磁盘总大小
function GetDiskInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfDisk_LogicalDisk",,48)

	Dim objItem 'as Win32_PerfFormattedData_PerfDisk_LogicalDisk
	For Each objItem in colItems
		if objItem.Name <> "_Total" Then
			WScript.Echo "Name: " & objItem.Name
			WScript.Echo "AvgDiskBytesPerRead: " & objItem.AvgDiskBytesPerRead
			WScript.Echo "AvgDiskBytesPerTransfer: " & objItem.AvgDiskBytesPerTransfer
			WScript.Echo "AvgDiskBytesPerWrite: " & objItem.AvgDiskBytesPerWrite
			WScript.Echo "AvgDisksecPerRead: " & objItem.AvgDisksecPerRead
			WScript.Echo "AvgDisksecPerTransfer: " & objItem.AvgDisksecPerTransfer
			WScript.Echo "AvgDisksecPerWrite: " & objItem.AvgDisksecPerWrite
			WScript.Echo "DiskBytesPersec: " & objItem.DiskBytesPersec
			WScript.Echo "DiskReadBytesPersec: " & objItem.DiskReadBytesPersec
			WScript.Echo "DiskReadsPersec: " & objItem.DiskReadsPersec
			WScript.Echo "DiskTransfersPersec: " & objItem.DiskTransfersPersec
			WScript.Echo "DiskWriteBytesPersec: " & objItem.DiskWriteBytesPersec
			WScript.Echo "DiskWritesPersec: " & objItem.DiskWritesPersec
			WScript.Echo "FreeMegabytes: " & objItem.FreeMegabytes
			WScript.Echo "PercentDiskReadTime: " & objItem.PercentDiskReadTime
			WScript.Echo "PercentDiskTime: " & objItem.PercentDiskTime
			WScript.Echo "PercentDiskWriteTime: " & objItem.PercentDiskWriteTime
			WScript.Echo "PercentFreeSpace: " & objItem.PercentFreeSpace
			WScript.Echo "PercentIdleTime: " & objItem.PercentIdleTime
		end if
	Next
end function

function GetProcessInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)

	Dim objItem 'as Win32_Process
	TotalVisibleMemorySize =  arrSystemInfo(0)

	For Each objItem in colItems
		WScript.Echo "ProcessProcessId: " & objItem.ProcessId
		WScript.Echo "ProcessName: " & objItem.Name
		WScript.Echo "ProcessParentProcessId: " & objItem.ParentProcessId
		WScript.Echo "ProcessThreadCount: " & objItem.ThreadCount
		WScript.Echo "ProcessHandleCount: " & objItem.HandleCount
		WScript.Echo "ProcessCpuTime:"  &  (objItem.KernelModeTime + objItem.UserModeTime) / 10000000
		' 单位 字节
        WScript.Echo "ProcessWorkingSetSize: " & objItem.WorkingSetSize
	  if TotalVisibleMemorySize > 0 Then
		WScript.Echo "ProcessProMemPercent: " & (objItem.WorkingSetSize\1024) \ TotalVisibleMemorySize
      else
	    WScript.Echo "ProcessProMemPercent: "
	  end if
	  	WScript.Echo "=================="
	Next
	
end function



function GetCPUInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)

	Dim objItem 'as Win32_Processor
	Dim intCpuCount
	intCpuCount = 0
	For Each objItem in colItems
		WScript.Echo "CpuAddressWidth: " & objItem.AddressWidth
		WScript.Echo "CpuCaption: " & objItem.Caption
		WScript.Echo "CpuCpuStatus: " & objItem.CpuStatus
		WScript.Echo "CpuCreationClassName: " & objItem.CreationClassName
		WScript.Echo "CpuCurrentClockSpeed: " & objItem.CurrentClockSpeed
		WScript.Echo "CpuCurrentVoltage: " & objItem.CurrentVoltage
		WScript.Echo "CpuDataWidth: " & objItem.DataWidth
		WScript.Echo "CpuDescription: " & objItem.Description
		WScript.Echo "CpuLoadPercentage: " & objItem.LoadPercentage
		WScript.Echo "CpuManufacturer: " & objItem.Manufacturer
		WScript.Echo "CpuMaxClockSpeed: " & objItem.MaxClockSpeed
		WScript.Echo "CpuName: " & objItem.Name
		' WScript.Echo "CpuNumberOfCores: " & objItem.NumberOfCores
		' WScript.Echo "CpuNumberOfEnabledCore: " & objItem.NumberOfEnabledCore
		' WScript.Echo "CpuNumberOfLogicalProcessors: " & objItem.NumberOfLogicalProcessors
		WScript.Echo "CpuProcessorId: " & objItem.ProcessorId
		WScript.Echo "CpuProcessorType: " & objItem.ProcessorType
		WScript.Echo "CpuRevision: " & objItem.Revision
		WScript.Echo "CpuRole: " & objItem.Role
		' WScript.Echo "CpuSecondLevelAddressTranslationExtensions: " & objItem.SecondLevelAddressTranslationExtensions
		' WScript.Echo "CpuSerialNumber: " & objItem.SerialNumber
		WScript.Echo "CpuSocketDesignation: " & objItem.SocketDesignation
		WScript.Echo "CpuSystemName: " & objItem.SystemName
		' WScript.Echo "CpuThreadCount: " & objItem.ThreadCount
		' WScript.Echo "CpuVirtualizationFirmwareEnabled: " & objItem.VirtualizationFirmwareEnabled
		' WScript.Echo "CpuVMMonitorModeExtensions: " & objItem.VMMonitorModeExtensions
		intCpuCount = intCpuCount + 1
	Next
	WScript.Echo "CpuCount:" & intCpuCount
end function


function GetSystemInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
	Dim objItem 'as Win32_OperatingSystem
	Dim TotalVisibleMemorySize
	For Each objItem in colItems
		WScript.Echo "SysCaption=" & objItem.Caption
		WScript.Echo "SysCSName=" & objItem.CSName
		WScript.Echo "SysDescription: " & objItem.Description
		WScript.Echo "SysFreePhysicalMemory: " & objItem.FreePhysicalMemory
		WScript.Echo "SysFreeSpaceInPagingFiles: " & objItem.FreeSpaceInPagingFiles
		WScript.Echo "SysFreeVirtualMemory: " & objItem.FreeVirtualMemory
		WScript.Echo "SysLargeSystemCache: " & objItem.LargeSystemCache
		WScript.Echo "SysManufacturer: " & objItem.Manufacturer
		WScript.Echo "SysMaxNumberOfProcesses: " & objItem.MaxNumberOfProcesses
		WScript.Echo "SysMaxProcessMemorySize: " & objItem.MaxProcessMemorySize
		WScript.Echo "SysName: " & Split(objItem.Name, "|")(0)
		WScript.Echo "SysNumberOfLicensedUsers: " & objItem.NumberOfLicensedUsers
		WScript.Echo "SysNumberOfProcesses: " & objItem.NumberOfProcesses
		WScript.Echo "SysNumberOfUsers: " & objItem.NumberOfUsers
		' WScript.Echo "SysOSArchitecture: " & objItem.OSArchitecture
		WScript.Echo "SysSerialNumber: " & objItem.SerialNumber
		WScript.Echo "SysTotalSwapSpaceSize: " & objItem.TotalSwapSpaceSize
		WScript.Echo "SysTotalVirtualMemorySize: " & objItem.TotalVirtualMemorySize
		WScript.Echo "SysTotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
		WScript.Echo "SysVersion: " & objItem.Version
		WScript.Echo "Sysruntime=" &  GetRuntimeStr(GetRuntimeSecond(objItem.LastBootUpTime))
		TotalVisibleMemorySize = objItem.TotalVisibleMemorySize
		
	Next
	GetSystemInfo=array(TotalVisibleMemorySize)
end function


function GetRuntimeStr(second)
	const intDaySecond = 86400 'day
	const intHourSecond = 3600 'day
	const intMinuteSecond = 60 'day

	intDay = second \ intDaySecond
	second = second mod intDaySecond    

	intHour = second \ intHourSecond
	second = second mod intHourSecond    

	intMinute = second \ intMinuteSecond
	second = second mod intMinuteSecond    

	if intDay > 0 Then
		strRuntime = strRuntime & intDay & "天"
	end if

	if intHour > 0 Then
		strRuntime = strRuntime  & intHour &"小时"
	end if

	if intMinute > 0 Then
		strRuntime = strRuntime  & intMinute & "分钟"
	end if
	strRuntime = strRuntime  & second & "秒"
	GetRuntimeStr=strRuntime
end function


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

