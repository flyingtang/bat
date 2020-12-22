On Error Resume Next
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
