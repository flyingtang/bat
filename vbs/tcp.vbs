On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_TCPv4",,48)

Dim objItem 'as Win32_PerfRawData_Tcpip_TCPv4
For Each objItem in colItems
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "ConnectionFailures: " & objItem.ConnectionFailures
	WScript.Echo "ConnectionsActive: " & objItem.ConnectionsActive
	WScript.Echo "ConnectionsEstablished: " & objItem.ConnectionsEstablished
	WScript.Echo "ConnectionsPassive: " & objItem.ConnectionsPassive
	WScript.Echo "ConnectionsReset: " & objItem.ConnectionsReset
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "Frequency_Object: " & objItem.Frequency_Object
	WScript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
	WScript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
	WScript.Echo "Name: " & objItem.Name
	WScript.Echo "SegmentsPersec: " & objItem.SegmentsPersec
	WScript.Echo "SegmentsReceivedPersec: " & objItem.SegmentsReceivedPersec
	WScript.Echo "SegmentsRetransmittedPersec: " & objItem.SegmentsRetransmittedPersec
	WScript.Echo "SegmentsSentPersec: " & objItem.SegmentsSentPersec
	WScript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
	WScript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
	WScript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
	WScript.Echo ""
Next
