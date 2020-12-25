
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_NetworkInterface",,48)

Dim objItem 'as Win32_PerfRawData_Tcpip_NetworkInterface
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
