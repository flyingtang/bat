On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Process",,48)

Dim objItem 'as Win32_PerfFormattedData_PerfProc_Process
For Each objItem in colItems
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "CreatingProcessID: " & objItem.CreatingProcessID
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "ElapsedTime: " & objItem.ElapsedTime
	WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
	WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
	WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
	WScript.Echo "HandleCount: " & objItem.HandleCount
	WScript.Echo "IDProcess: " & objItem.IDProcess
	WScript.Echo "IODataBytesPersec: " & objItem.IODataBytesPersec
	WScript.Echo "IODataOperationsPersec: " & objItem.IODataOperationsPersec
	WScript.Echo "IOOtherBytesPersec: " & objItem.IOOtherBytesPersec
	WScript.Echo "IOOtherOperationsPersec: " & objItem.IOOtherOperationsPersec
	WScript.Echo "IOReadBytesPersec: " & objItem.IOReadBytesPersec
	WScript.Echo "IOReadOperationsPersec: " & objItem.IOReadOperationsPersec
	WScript.Echo "IOWriteBytesPersec: " & objItem.IOWriteBytesPersec
	WScript.Echo "IOWriteOperationsPersec: " & objItem.IOWriteOperationsPersec
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "PageFaultsPersec: " & objItem.PageFaultsPersec
	WScript.Echo "PageFileBytes: " & objItem.PageFileBytes
	WScript.Echo "PageFileBytesPeak: " & objItem.PageFileBytesPeak
	WScript.Echo "PercentPrivilegedTime: " & objItem.PercentPrivilegedTime
	WScript.Echo "PercentProcessorTime: " & objItem.PercentProcessorTime
	WScript.Echo "PercentUserTime: " & objItem.PercentUserTime
	WScript.Echo "PoolNonpagedBytes: " & objItem.PoolNonpagedBytes
	WScript.Echo "PoolPagedBytes: " & objItem.PoolPagedBytes
	WScript.Echo "PriorityBase: " & objItem.PriorityBase
	WScript.Echo "PrivateBytes: " & objItem.PrivateBytes
	WScript.Echo "ThreadCount: " & objItem.ThreadCount
	WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
	WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
	WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
	WScript.Echo "VirtualBytes: " & objItem.VirtualBytes
	WScript.Echo "VirtualBytesPeak: " & objItem.VirtualBytesPeak
	WScript.Echo "WorkingSet: " & objItem.WorkingSet
	WScript.Echo "WorkingSetPeak: " & objItem.WorkingSetPeak
	WScript.Echo "WorkingSetPrivate: " & objItem.WorkingSetPrivate
	WScript.Echo ""
Next
