class ObjDiskInfo
    Public ObjDisks
    Public mapNameDiskObj
    Public Capacity
    Public VolunmFreeSpace
    Public VolunmUsedSpace
     private sub class_Initialize
         ' Called automatically when class is created
         ObjDisks = Array()
         Capacity = 0
         VolunmFreeSpace = 0
         Set mapNameDiskObj = CreateObject("Scripting.Dictionary")
     end sub
 
     private sub class_Terminate
         ' Called automatically when all references to class instance are removed
     end sub

     sub CollectVolumn()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume",,48)
        
        Dim objItem 'as Win32_Volume
        For Each objItem in colItems
            set tmp = New ObjVolumn
            tmp.VolunmLabel = objItem.Label
            tmp.VolunmName = Replace(Replace(objItem.Name, "\", ""), " ", "")
            tmp.VolunmSerialNumber = objItem.SerialNumber
            tmp.VolunmFreeSpace = objItem.FreeSpace
            tmp.VolunmCapacity = objItem.Capacity
            tmp.VolunmUsedSpace = objItem.Capacity - objItem.FreeSpace
            tmp.VolunmPercentUsed = tmp.VolunmUsedSpace / objItem.Capacity * 100 
            tmp.VolunmFileSystem = objItem.FileSystem

            ' tmp.VolunmDriveType = objItem.DriveType
            select case objItem.DriveType
                case  "0"
                    tmp.VolunmDriveType ="Unknown"
                case  "1"
                    tmp.VolunmDriveType ="No Root Directory"
                case  "2"
                    tmp.VolunmDriveType ="Removable Disk"
                case  "3"
                    tmp.VolunmDriveType ="Local Disk"
                case  "4"
                    tmp.VolunmDriveType ="Network Drive"
                case  "5"
                    tmp.VolunmDriveType ="Compact Disc"
                case  "6"
                    tmp.VolunmDriveType ="RAM Disk"
                case else
            end select
       
            Capacity = Capacity + objItem.Capacity
            VolunmFreeSpace = VolunmFreeSpace + objItem.FreeSpace
            mapNameDiskObj.Add tmp.VolunmName, tmp
         Next
         VolunmUsedSpace = Capacity - VolunmFreeSpace
     end sub
 
     sub Collect()
        Call CollectVolumn
       ' io 信息
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfDisk_LogicalDisk",,48)
        Dim objItem 'as Win32_PerfFormattedData_PerfDisk_LogicalDisk
        For Each objItem in colItems
            if objItem.Name <> "_Total" Then
                size = UBound(ObjDisks)
                if size < 0 Then
                    size = 0
                end if

                WScript.Echo Size
                Redim Preserve ObjDisks(size + 1)

                set tmp = New ObjDisk
                tmp.DiskName = objItem.Name
                tmp.DiskAvgDiskBytesPerRead = objItem.AvgDiskBytesPerRead
                tmp.DiskAvgDiskBytesPerTransfer = objItem.AvgDiskBytesPerTransfer
                tmp.DiskAvgDiskBytesPerWrite = objItem.AvgDiskBytesPerWrite
                tmp.DiskAvgDisksecPerRead = objItem.AvgDisksecPerRead
                tmp.DiskAvgDisksecPerTransfer = objItem.AvgDisksecPerTransfer
                tmp.DiskAvgDisksecPerWrite = objItem.AvgDisksecPerWrite
                tmp.DiskBytesPersec = objItem.DiskBytesPersec
                tmp.DiskReadBytesPersec = objItem.DiskReadBytesPersec
                tmp.DiskReadsPersec = objItem.DiskReadsPersec
                tmp.DiskTransfersPersec = objItem.DiskTransfersPersec
                tmp.DiskWriteBytesPersec = objItem.DiskWriteBytesPersec
                tmp.DiskWritesPersec = objItem.DiskWritesPersec
                tmp.DiskFreeMegabytes = objItem.FreeMegabytes
                tmp.DiskPercentDiskReadTime = objItem.PercentDiskReadTime
                tmp.DiskPercentDiskTime = objItem.PercentDiskTime
                tmp.DiskPercentDiskWriteTime = objItem.PercentDiskWriteTime
                tmp.DiskPercentFreeSpace = objItem.PercentFreeSpace
                tmp.DiskPercentIdleTime = objItem.PercentIdleTime

                if mapNameDiskObj.Exists(tmp.DiskName) Then
                    set obj = mapNameDiskObj.Item(tmp.DiskName)
                    tmp.VolunmLabel = obj.VolunmLabel
                    tmp.VolunmName = obj.VolunmName
                    tmp.VolunmSerialNumber = obj.VolunmSerialNumber
                    tmp.VolunmFreeSpace = obj.VolunmFreeSpace
                    tmp.VolunmCapacity = obj.VolunmCapacity
                    tmp.VolunmUsedSpace = obj.VolunmUsedSpace
                    tmp.VolunmPercentUsed = obj.VolunmPercentUsed
                    tmp.VolunmDriveType = obj.VolunmDriveType
                    tmp.VolunmFileSystem = obj.VolunmFileSystem
       
                end if 
                set ObjDisks(size) = tmp
            end if
        Next
    end sub
     
     sub Print()

        For Each tmp In ObjDisks
            if not IsEmpty(tmp) Then
                call tmp.Print
            end if
        Next 
        WScript.Echo "VolunmTotalCapacity=" & Capacity
        WScript.Echo "VolunmTotalFreeSpace=" & VolunmFreeSpace
        WScript.Echo "VolunmTotalUsedSpace=" & VolunmUsedSpace
     end sub
 end class


 class ObjVolumn
   Public VolunmLabel
   Public VolunmName
   Public VolunmSerialNumber
   Public VolunmFreeSpace
   Public VolunmCapacity
   Public VolunmUsedSpace
   Public VolunmPercentUsed
   Public VolunmDriveType
   Public VolunmFileSystem
     private sub class_Initialize
         ' Called automatically when class is created
     end sub
 
     private sub class_Terminate
         ' Called automatically when all references to class instance are removed
     end sub
 end class


 class ObjDisk
    Public DiskName
    Public DiskAvgDiskBytesPerRead
    Public DiskAvgDiskBytesPerTransfer
    Public DiskAvgDiskBytesPerWrite
    Public DiskAvgDisksecPerRead
    Public DiskAvgDisksecPerTransfer
    Public DiskAvgDisksecPerWrite
    Public DiskBytesPersec
    Public DiskReadBytesPersec
    Public DiskReadsPersec
    Public DiskTransfersPersec
    Public DiskWriteBytesPersec
    Public DiskWritesPersec
    Public DiskFreeMegabytes
    Public DiskPercentDiskReadTime
    Public DiskPercentDiskTime
    Public DiskPercentDiskWriteTime
    Public DiskPercentFreeSpace
    Public DiskPercentIdleTime
    Public VolunmLabel
    Public VolunmName
    Public VolunmSerialNumber
    Public VolunmFreeSpace
    Public VolunmCapacity
    Public VolunmUsedSpace
    Public VolunmPercentUsed
    Public VolunmDriveType
    Public VolunmFileSystem
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Print()
        WScript.Echo "DiskName=" & DiskName 
        WScript.Echo "DiskAvgDiskBytesPerRead=" & DiskAvgDiskBytesPerRead 
        WScript.Echo "DiskAvgDiskBytesPerTransfer=" & DiskAvgDiskBytesPerTransfer 
        WScript.Echo "DiskAvgDiskBytesPerWrite=" & DiskAvgDiskBytesPerWrite 
        WScript.Echo "DiskAvgDisksecPerRead=" & DiskAvgDisksecPerRead 
        WScript.Echo "DiskAvgDisksecPerTransfer=" & DiskAvgDisksecPerTransfer 
        WScript.Echo "DiskAvgDisksecPerWrite=" & DiskAvgDisksecPerWrite 
        WScript.Echo "DiskBytesPersec=" & DiskBytesPersec 
        WScript.Echo "DiskReadBytesPersec=" & DiskReadBytesPersec 
        WScript.Echo "DiskReadsPersec=" & DiskReadsPersec 
        WScript.Echo "DiskTransfersPersec=" & DiskTransfersPersec 
        WScript.Echo "DiskWriteBytesPersec=" & DiskWriteBytesPersec 
        WScript.Echo "DiskWritesPersec=" & DiskWritesPersec 
        WScript.Echo "DiskFreeMegabytes=" & DiskFreeMegabytes 
        WScript.Echo "DiskPercentDiskReadTime=" & DiskPercentDiskReadTime 
        WScript.Echo "DiskPercentDiskTime=" & DiskPercentDiskTime 
        WScript.Echo "DiskPercentDiskWriteTime=" & DiskPercentDiskWriteTime 
        WScript.Echo "DiskPercentFreeSpace=" & DiskPercentFreeSpace 
        WScript.Echo "DiskPercentIdleTime=" & DiskPercentIdleTime 
        WScript.Echo "VolunmLabel=" & VolunmLabel
        WScript.Echo "VolunmName=" & VolunmName
        WScript.Echo "VolunmSerialNumber=" & VolunmSerialNumber
        WScript.Echo "VolunmFreeSpace=" & VolunmFreeSpace
        WScript.Echo "VolunmCapacity=" & VolunmCapacity
        WScript.Echo "VolunmUsedSpace=" & VolunmUsedSpace
        WScript.Echo "VolunmPercentUsed=" & VolunmPercentUsed
        WScript.Echo "VolunmDriveType=" & VolunmDriveType
        WScript.Echo "VolunmFileSystem=" & VolunmFileSystem

     end sub
end class

 set objSyInfo = New ObjDiskInfo
objSyInfo.Collect
objSyInfo.Print