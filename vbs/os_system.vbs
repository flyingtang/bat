On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfOS_System",,48)

Dim objItem 'as Win32_PerfRawData_PerfOS_System
For Each objItem in colItems
	WScript.Echo "AlignmentFixupsPersec: " & objItem.AlignmentFixupsPersec
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "ContextSwitchesPersec: " & objItem.ContextSwitchesPersec
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "ExceptionDispatchesPersec: " & objItem.ExceptionDispatchesPersec
	WScript.Echo "FileControlBytesPersec: " & objItem.FileControlBytesPersec
	WScript.Echo "FileControlOperationsPersec: " & objItem.FileControlOperationsPersec
	WScript.Echo "FileDataOperationsPersec: " & objItem.FileDataOperationsPersec
	WScript.Echo "FileReadBytesPersec: " & objItem.FileReadBytesPersec
	WScript.Echo "FileReadOperationsPersec: " & objItem.FileReadOperationsPersec
	WScript.Echo "FileWriteBytesPersec: " & objItem.FileWriteBytesPersec
	WScript.Echo "FileWriteOperationsPersec: " & objItem.FileWriteOperationsPersec
	WScript.Echo "FloatingEmulationsPersec: " & objItem.FloatingEmulationsPersec
	WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
	WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
	WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "PercentRegistryQuotaInUse: " & objItem.PercentRegistryQuotaInUse
	WScript.Echo "PercentRegistryQuotaInUse_Base: " & objItem.PercentRegistryQuotaInUse_Base
	WScript.Echo "Processes: " & objItem.Processes
	WScript.Echo "ProcessorQueueLength: " & objItem.ProcessorQueueLength
	WScript.Echo "SystemCallsPersec: " & objItem.SystemCallsPersec
	WScript.Echo "SystemUpTime: " & objItem.SystemUpTime
	WScript.Echo "Threads: " & objItem.Threads
	WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
	WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
	WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
	WScript.Echo ""
Next
