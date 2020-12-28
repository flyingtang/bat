Class ObjProcessInfo
    Public arrProcessObj 
    Private intProcessCount
    Public mapIdProcessObj
    Public mapProcessIdCmdnTask
    Public mapIdProcessPerfObj
    Private Sub Class_Initialize
        arrProcessObj = Array()
        intProcessCount = 0
        Set mapIdProcessObj = CreateObject("Scripting.Dictionary")
        Set mapProcessIdCmdnTask = CreateObject("Scripting.Dictionary")
        Set mapIdProcessPerfObj = CreateObject("Scripting.Dictionary")
    End Sub

    Private Sub Class_Terminate
      
    End Sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

        Set perfColItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Process",,48)
       
        Dim objPerfItem 'as Win32_PerfFormattedData_PerfProc_Process
        For Each objPerfItem in perfColItems
            set tmp = New ObjPerfProcess
            tmp.ProcessProcessId = objPerfItem.IDProcess
            tmp.ProcessPercentProcessorTime = objPerfItem.PercentProcessorTime
            tmp.ProcessWorkingSet = objPerfItem.WorkingSet
            if  objPerfItem.IDProcess <> 0 Then 
                mapIdProcessPerfObj.Add objPerfItem.IDProcess, tmp
            end if 
        Next

        strResult = createobject("wscript.shell").exec("TASKLIST /V /FO LIST").StdOut.ReadAll
        arrSplitStr =  Split(strResult, vbCrLf)
        Dim  arrSplit() 
        ReDim arrSplit(9, 1) 
        currentRows = 0
        for i = 0 to UBound(arrSplitStr) step 1
            if i <> 0 Then
                size  = UBound( arrSplit, 2) 
                if arrSplitStr(i) = "" Then
                    ReDim Preserve arrSplit(9, size+1) 
                    currentRows = 0
                else
                    arrSplit(currentRows, size-1) = (Split(arrSplitStr(i), ":", 2))(1)
                    currentRows = currentRows + 1
                end if
            end if
        next
        
        for i = 0 to  UBound( arrSplit, 2)  step 1
            set tmp = New ObjCmdTask
            tmp.pid=CInt(arrSplit(1, i))
            tmp.memused=arrSplit(4, i)
            tmp.cpuused=arrSplit(7, i)  
            if Not mapProcessIdCmdnTask.Exists(tmp.pid) Then 
                  mapProcessIdCmdnTask.Add tmp.pid, tmp
            end if 
        next

        Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)
        Dim objItem 'as Win32_Process 
        For Each objItem in colItems
            count = intProcessCount
            intProcessCount = intProcessCount+1
            ReDim Preserve arrProcessObj(intProcessCount)
            set arrProcessObj(count) = New ObjProcess
            arrProcessObj(count).ProcessProcessId = objItem.ProcessId
            arrProcessObj(count).ProcessName = objItem.Name
            arrProcessObj(count).ProcessParentProcessId = objItem.ParentProcessId
            arrProcessObj(count).ProcessThreadCount = objItem.ThreadCount
            arrProcessObj(count).ProcessHandleCount = objItem.HandleCount
            arrProcessObj(count).ProcessKernelModeTime = objItem.KernelModeTime
            arrProcessObj(count).ProcessUserModeTime = objItem.UserModeTime
            arrProcessObj(count).ProcessCpuTime =  (objItem.KernelModeTime + objItem.UserModeTime) / 10000000
            arrProcessObj(count).ProcessWorkingSetSize =objItem.WorkingSetSize
                    ' if TotalVisibleMemorySize > 0 Then
            '     WScript.Echo "ProcessProMemPercent=" & (objItem.WorkingSetSize\1024) \ TotalVisibleMemorySize
            ' else
            '     WScript.Echo "ProcessProMemPercent="
            
            ' WScript.Echo "arrProcessObj(count).ProcessProcessId===" & arrProcessObj(count).ProcessProcessId
            ' WScript.Echo mapIdProcessObj.Item(arrProcessObj(count).ProcessProcessId).ProcessName

            if objItem.ProcessId <> 0 Then
                if mapIdProcessPerfObj.Exists(objItem.ProcessId) Then 
                    set tmp =  mapIdProcessPerfObj.Item(objItem.ProcessId)   
                    arrProcessObj(count).ProcessPercentProcessorTime = tmp.ProcessPercentProcessorTime
                    arrProcessObj(count).ProcessWorkingSet = tmp.ProcessWorkingSet
                    'TODO 
                    if  mapProcessIdCmdnTask.Exists(cint(objItem.ProcessId)) Then 
                        set ttmp = mapProcessIdCmdnTask.Item(objItem.ProcessId)
                        arrProcessObj(count).ProcessMemused = ttmp.memused
                        arrProcessObj(count).ProcessCpuused = ttmp.cpuused
                    end if 
               end if 
            end if 
            mapIdProcessObj.Add arrProcessObj(count).ProcessProcessId, arrProcessObj(count)
        Next
    end sub

    sub Print()
        for each processObj in arrProcessObj
            if IsEmpty(processObj) Then
            else
                processObj.ToString
                ' if processObj.ProcessProcessId <> 0 Then
                '     WScript.Echo "processObj.ProcessProcessId=== " & processObj.ProcessProcessId 
                '     WScript.Echo "ProcesxxxxxxxxxxxxxxxxxsWorkingSet" & mapIdProcessObj.Item(processObj.ProcessProcessId).ProcessWorkingSet
                ' end if 
            end if
        next    
    end sub
End Class

class ObjProcess
    Public ProcessProcessId
    Public ProcessName
    Public ProcessParentProcessId
    Public ProcessThreadCount
    Public ProcessHandleCount
    Public ProcessCpuTime
    Public ProcessKernelModeTime
    Public ProcessUserModeTime
    Public ProcessWorkingSetSize
    Public ProcessPercentProcessorTime
    Public ProcessWorkingSet
    Public ProcessMemused
    Public ProcessCpuused
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    Public sub ToString()
        WScript.Echo "ProcessProcessId=" & ProcessProcessId
        WScript.Echo "ProcessName=" & ProcessName
		WScript.Echo "ProcessParentProcessId=" & ProcessParentProcessId
		WScript.Echo "ProcessThreadCount=" & ProcessThreadCount
		WScript.Echo "ProcessHandleCount=" & ProcessHandleCount
        WScript.Echo "ProcessCpuTime=" & ProcessCpuTime
        WScript.Echo "ProcessKernelModeTime=" & ProcessKernelModeTime
        WScript.Echo "ProcessUserModeTime=" & ProcessUserModeTime
        WScript.Echo "ProcessWorkingSetSize="  & ProcessWorkingSetSize
        WScript.Echo "ProcessPercentProcessorTime="  & ProcessPercentProcessorTime
        WScript.Echo "ProcessWorkingSet="  & ProcessWorkingSet
        WScript.Echo "ProcessMemused="  & ProcessMemused
        WScript.Echo "ProcessCpuused="  & ProcessCpuused
    end sub
end class

class ObjPerfProcess
    Public ProcessProcessId
    Public ProcessPercentProcessorTime
    Public ProcessWorkingSet
    Public ProcessName
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

end class

class ObjCmdTask
    Public pid
    Public memused
    Public cpuused
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub
end class

class ObjSystemInfo
    Public SysCaption
    Public SysCSName
    Public SysDescription
    Public SysFreePhysicalMemory
    Public SysFreeSpaceInPagingFiles
    Public SysFreeVirtualMemory
    Public SysLargeSystemCache
    Public SysManufacturer
    Public SysMaxNumberOfProcesses
    Public SysMaxProcessMemorySize
    Public SysName
    Public SysNumberOfLicensedUsers
    Public SysNumberOfProcesses
    Public SysNumberOfUsers
    Public SysSerialNumber
    Public SysTotalSwapSpaceSize
    Public SysTotalVirtualMemorySize
    Public SysTotalVisibleMemorySize
    Public SysVersion
    Public Sysruntime
    Public SysPercentUsedMemory
    Public SysUsedMemory
    Private intDaySecond  'day
    Private intHourSecond 'day
    Private intMinuteSecond'day

    private sub class_Initialize
        ' Called automatically when class is created
        intDaySecond = 86400 
        intHourSecond = 3600
        intMinuteSecond = 60 

    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
        Dim objItem 'as Win32_OperatingSystem
        Dim TotalVisibleMemorySize
        For Each objItem in colItems
             SysCaption =  objItem.Caption
             SysCSName = objItem.CSName
             SysDescription = objItem.Description
             SysFreePhysicalMemory = objItem.FreePhysicalMemory
             SysFreeSpaceInPagingFiles = objItem.FreeSpaceInPagingFiles
             SysFreeVirtualMemory =  objItem.FreeVirtualMemory
             SysLargeSystemCache = objItem.LargeSystemCache
             SysManufacturer = objItem.Manufacturer
             SysMaxNumberOfProcesses =  objItem.MaxNumberOfProcesses
             SysMaxProcessMemorySize = objItem.MaxProcessMemorySize
             SysName = Split(objItem.Name, "|")(0)
             SysNumberOfLicensedUsers = objItem.NumberOfLicensedUsers
             SysNumberOfProcesses =  objItem.NumberOfProcesses
             SysNumberOfUsers =  objItem.NumberOfUsers
             SysSerialNumber = objItem.SerialNumber
             SysTotalSwapSpaceSize = objItem.TotalSwapSpaceSize
             SysTotalVirtualMemorySize = objItem.TotalVirtualMemorySize
             SysTotalVisibleMemorySize = objItem.TotalVisibleMemorySize
             SysVersion = objItem.Version
             Sysruntime =   GetRuntimeStr(GetRuntimeSecond(objItem.LastBootUpTime))
             SysUsedMemory =  objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory
             SysPercentUsedMemory =  ((objItem.TotalVisibleMemorySize - objItem.FreePhysicalMemory) \ objItem.TotalVisibleMemorySize) * 100
        Next
    end sub

    function GetRuntimeStr(second)

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
    sub Print()
        WScript.Echo "SysCaption=" & SysCaption
        WScript.Echo "SysCSName=" & SysCSName
        WScript.Echo "SysDescription=" & SysDescription
        WScript.Echo "SysFreePhysicalMemory=" & SysFreePhysicalMemory
        WScript.Echo "SysFreeSpaceInPagingFiles=" & SysFreeSpaceInPagingFiles
        WScript.Echo "SysFreeVirtualMemory=" & SysFreeVirtualMemory
        WScript.Echo "SysLargeSystemCache=" & SysLargeSystemCache
        WScript.Echo "SysManufacturer=" & SysManufacturer
        WScript.Echo "SysMaxNumberOfProcesses=" & SysMaxNumberOfProcesses
        WScript.Echo "SysMaxProcessMemorySize=" & SysMaxProcessMemorySize
        WScript.Echo "SysName=" & SysName
        WScript.Echo "SysNumberOfLicensedUsers=" & SysNumberOfLicensedUsers
        WScript.Echo "SysNumberOfProcesses=" & SysNumberOfProcesses
        WScript.Echo "SysNumberOfUsers=" & SysNumberOfUsers
        WScript.Echo "SysSerialNumber=" & SysSerialNumber
        WScript.Echo "SysTotalSwapSpaceSize=" & SysTotalSwapSpaceSize
        WScript.Echo "SysTotalVirtualMemorySize=" & SysTotalVirtualMemorySize
        WScript.Echo "SysTotalVisibleMemorySize=" & SysTotalVisibleMemorySize
        WScript.Echo "SysVersion= " & SysVersion
        WScript.Echo "Sysruntime=" &  Sysruntime
        WScript.Echo "SysUsedMemory=" &  SysUsedMemory
        WScript.Echo "SysPercentUsedMemory=" &  SysPercentUsedMemory
    end sub
end class

class ObjCpuInfo
    Public CpuAddressWidth
    Public CpuCaption
    Public CpuCpuStatus
    Public CpuCreationClassName
    Public CpuCurrentClockSpeed
    Public CpuCurrentVoltage
    Public CpuDataWidth
    Public CpuDescription
    Public CpuLoadPercentage
    Public CpuManufacturer
    Public CpuMaxClockSpeed
    Public CpuName
    Public CpuProcessorId
    Public CpuProcessorType
    Public CpuRevision
    Public CpuRole
    Public CpuSocketDesignation
    Public CpuSystemName
     Public CpuCount
    private sub class_Initialize
        CpuCount = 0
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
        Dim objItem 'as Win32_Processor
        For Each objItem in colItems
            CpuAddressWidth = objItem.AddressWidth
            CpuCaption = objItem.Caption
            CpuCpuStatus = objItem.CpuStatus
            CpuCreationClassName = objItem.CreationClassName
            CpuCurrentClockSpeed = objItem.CurrentClockSpeed
            CpuCurrentVoltage = objItem.CurrentVoltage
            CpuDataWidth = objItem.DataWidth
            CpuDescription = objItem.Description
            CpuLoadPercentage = objItem.LoadPercentage
            CpuManufacturer = objItem.Manufacturer
            CpuMaxClockSpeed = objItem.MaxClockSpeed
            CpuName = objItem.Name
            CpuProcessorId = objItem.ProcessorId
            CpuProcessorType = objItem.ProcessorType
            CpuRevision = objItem.Revision
            CpuRole = objItem.Role
            CpuSocketDesignation = objItem.SocketDesignation
            CpuSystemName = objItem.SystemName
            CpuCount = CpuCount + 1
        Next
    end sub

    sub Print()
        WScript.Echo "CpuAddressWidth=" & CpuAddressWidth  
        WScript.Echo "CpuCaption=" & CpuCaption 
        WScript.Echo "CpuCpuStatus=" & CpuCpuStatus 
        WScript.Echo "CpuCreationClassName=" & CpuCreationClassName 
        WScript.Echo "CpuCurrentClockSpeed=" & CpuCurrentClockSpeed 
        WScript.Echo "CpuCurrentVoltage=" & CpuCurrentVoltage 
        WScript.Echo "CpuDataWidth=" & CpuDataWidth 
        WScript.Echo "CpuDescription=" & CpuDescription 
        WScript.Echo "CpuLoadPercentage=" & CpuLoadPercentage 
        WScript.Echo "CpuManufacturer=" & CpuManufacturer 
        WScript.Echo "CpuMaxClockSpeed=" & CpuMaxClockSpeed 
        WScript.Echo "CpuName=" & CpuName 
        WScript.Echo "CpuProcessorId=" & CpuProcessorId 
        WScript.Echo "CpuProcessorType=" & CpuProcessorType 
        WScript.Echo "CpuRevision=" & CpuRevision 
        WScript.Echo "CpuRole=" & CpuRole 
        WScript.Echo "CpuSocketDesignation=" & CpuSocketDesignation 
        WScript.Echo "CpuSystemName=" & CpuSystemName 
        WScript.Echo "CpuCount=" & CpuCount 
   end sub
 end class


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
       ' io ???
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

class ObjNetworkInterface
    Public NetworkInterfaceBytesReceivedPersec
    Public NetworkInterfaceBytesSentPersec
    Public NetworkInterfaceBytesTotalPersec
    Public NetworkInterfaceCurrentBandwidth
    Public NetworkInterfaceName
    Public NetworkInterfacePacketsPersec
    Public NetworkInterfacePacketsReceivedDiscarded
    Public NetworkInterfacePacketsReceivedErrors
    Public NetworkInterfacePacketsReceivedNonUnicastPersec
    Public NetworkInterfacePacketsReceivedPersec
    Public NetworkInterfacePacketsReceivedUnicastPersec
    Public NetworkInterfacePacketsReceivedUnknown
    Public NetworkInterfacePacketsSentNonUnicastPersec
    Public NetworkInterfacePacketsSentPersec
    Public NetworkInterfacePacketsSentUnicastPersec
    Public NetworkInterfaceTCPActiveRSCConnections
    Public NetworkInterfaceTCPRSCAveragePacketSize
    Public NetworkInterfaceTCPRSCCoalescedPacketsPersec
    Public NetworkInterfaceTCPRSCExceptionsPersec
    Public NetworkInterfaceTimestamp_Object
    Public NetworkInterfaceTimestamp_PerfTime
    Public NetworkInterfaceTimestamp_Sys100NS

    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        
    end sub

    sub Print()
        
    end sub

end class

class ObjNetworkAdaptor
    Public NetworkAdaptorAdapterType
    Public NetworkAdaptorAdapterTypeId
    Public NetworkAdaptorCaption
    Public NetworkAdaptorDescription
    Public NetworkAdaptorMACAddress
    Public NetworkAdaptorManufacturer
    Public NetworkAdaptorMaxSpeed
    Public NetworkAdaptorName
    Public NetworkAdaptorNetConnectionID
    Public NetworkAdaptorNetConnectionStatus
    Public NetworkAdaptorNetEnabled
    Public NetworkAdaptorNetworkAddresses
    Public NetworkAdaptorPermanentAddress
    Public NetworkAdaptorPhysicalAdapter
    Public NetworkAdaptorPNPDeviceID
    Public NetworkAdaptorProductName
    Public NetworkAdaptorServiceName
    Public NetworkAdaptorSpeed
    Public NetworkAdaptorStatus
    Public NetworkAdaptorStatusInfo

    Public NetworkAdaptorBytesReceivedPersec
    Public NetworkAdaptorBytesSentPersec
    Public NetworkAdaptorBytesTotalPersec
    Public NetworkAdaptorCurrentBandwidth
    Public NetworkAdaptorFrequency_Object
    Public NetworkAdaptorFrequency_PerfTime
    Public NetworkAdaptorFrequency_Sys100NS
    Public NetworkAdaptorPacketsPersec
    Public NetworkAdaptorPacketsReceivedDiscarded
    Public NetworkAdaptorPacketsReceivedErrors
    Public NetworkAdaptorPacketsReceivedNonUnicastPersec
    Public NetworkAdaptorPacketsReceivedPersec
    Public NetworkAdaptorPacketsReceivedUnicastPersec
    Public NetworkAdaptorPacketsReceivedUnknown
    Public NetworkAdaptorPacketsSentNonUnicastPersec
    Public NetworkAdaptorPacketsSentPersec
    Public NetworkAdaptorPacketsSentUnicastPersec
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        
    end sub

    sub Print()
       WScript.Echo "NetworkAdaptorAdapterType=" & NetworkAdaptorAdapterType
       WScript.Echo "NetworkAdaptorAdapterTypeId=" & NetworkAdaptorAdapterTypeId
       WScript.Echo "NetworkAdaptorCaption=" & NetworkAdaptorCaption
       WScript.Echo "NetworkAdaptorDescription=" & NetworkAdaptorDescription
       WScript.Echo "NetworkAdaptorMACAddress=" & NetworkAdaptorMACAddress
       WScript.Echo "NetworkAdaptorManufacturer=" & NetworkAdaptorManufacturer
       WScript.Echo "NetworkAdaptorMaxSpeed=" & NetworkAdaptorMaxSpeed
       WScript.Echo "NetworkAdaptorName=" & NetworkAdaptorName
       WScript.Echo "NetworkAdaptorNetConnectionID=" & NetworkAdaptorNetConnectionID
       WScript.Echo "NetworkAdaptorNetConnectionStatus=" & NetworkAdaptorNetConnectionStatus
       WScript.Echo "NetworkAdaptorNetEnabled=" & NetworkAdaptorNetEnabled
       WScript.Echo "NetworkAdaptorNetworkAddresses=" & NetworkAdaptorNetworkAddresses
       WScript.Echo "NetworkAdaptorPermanentAddress=" & NetworkAdaptorPermanentAddress
       WScript.Echo "NetworkAdaptorPhysicalAdapter=" & NetworkAdaptorPhysicalAdapter
       WScript.Echo "NetworkAdaptorPNPDeviceID=" & NetworkAdaptorPNPDeviceID
       WScript.Echo "NetworkAdaptorProductName=" & NetworkAdaptorProductName
       WScript.Echo "NetworkAdaptorServiceName=" & NetworkAdaptorServiceName
       WScript.Echo "NetworkAdaptorSpeed=" & NetworkAdaptorSpeed
       WScript.Echo "NetworkAdaptorStatus=" & NetworkAdaptorStatus
       WScript.Echo "NetworkAdaptorStatusInfo=" & NetworkAdaptorStatusInfo

       WScript.Echo "NetworkAdaptorBytesReceivedPersec=" & NetworkAdaptorBytesReceivedPersec
       WScript.Echo "NetworkAdaptorBytesSentPersec=" & NetworkAdaptorBytesSentPersec
       WScript.Echo "NetworkAdaptorBytesTotalPersec=" & NetworkAdaptorBytesTotalPersec
       WScript.Echo "NetworkAdaptorCurrentBandwidth=" & NetworkAdaptorCurrentBandwidth
       WScript.Echo "NetworkAdaptorFrequency_Object=" & NetworkAdaptorFrequency_Object
       WScript.Echo "NetworkAdaptorFrequency_PerfTime=" & NetworkAdaptorFrequency_PerfTime
       WScript.Echo "NetworkAdaptorFrequency_Sys100NS=" & NetworkAdaptorFrequency_Sys100NS
       WScript.Echo "NetworkAdaptorPacketsPersec=" & NetworkAdaptorPacketsPersec
       WScript.Echo "NetworkAdaptorPacketsReceivedDiscarded=" & NetworkAdaptorPacketsReceivedDiscarded
       WScript.Echo "NetworkAdaptorPacketsReceivedErrors=" & NetworkAdaptorPacketsReceivedErrors
       WScript.Echo "NetworkAdaptorPacketsReceivedNonUnicastPersec=" & NetworkAdaptorPacketsReceivedNonUnicastPersec
       WScript.Echo "NetworkAdaptorPacketsReceivedPersec=" & NetworkAdaptorPacketsReceivedPersec
       WScript.Echo "NetworkAdaptorPacketsReceivedUnicastPersec=" & NetworkAdaptorPacketsReceivedUnicastPersec
       WScript.Echo "NetworkAdaptorPacketsReceivedUnknown=" & NetworkAdaptorPacketsReceivedUnknown
       WScript.Echo "NetworkAdaptorPacketsSentNonUnicastPersec=" & NetworkAdaptorPacketsSentNonUnicastPersec
       WScript.Echo "NetworkAdaptorPacketsSentPersec=" & NetworkAdaptorPacketsSentPersec
       WScript.Echo "NetworkAdaptorPacketsSentUnicastPersec=" & NetworkAdaptorPacketsSentUnicastPersec
    end sub
end class



class ObjNetworkAdaptorIO
    Public NetworkInterfaceBytesReceivedPersec
    Public NetworkInterfaceBytesSentPersec
    Public NetworkInterfaceBytesTotalPersec
    Public NetworkInterfaceCurrentBandwidth
    Public NetworkInterfaceName
    Public NetworkInterfacePacketsPersec
    Public NetworkInterfacePacketsReceivedDiscarded
    Public NetworkInterfacePacketsReceivedErrors
    Public NetworkInterfacePacketsReceivedNonUnicastPersec
    Public NetworkInterfacePacketsReceivedPersec
    Public NetworkInterfacePacketsReceivedUnicastPersec
    Public NetworkInterfacePacketsReceivedUnknown
    Public NetworkInterfacePacketsSentNonUnicastPersec
    Public NetworkInterfacePacketsSentPersec
    Public NetworkInterfacePacketsSentUnicastPersec
    Public NetworkInterfaceTimestamp_Object
    Public NetworkInterfaceTimestamp_PerfTime
    Public NetworkInterfaceTimestamp_Sys100NS
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        
    end sub

    sub Print()
       WScript.Echo "NetworkInterfaceBytesReceivedPersec=" & NetworkInterfaceBytesReceivedPersec
       WScript.Echo "NetworkInterfaceBytesSentPersec=" & NetworkInterfaceBytesSentPersec
       WScript.Echo "NetworkInterfaceBytesTotalPersec=" & NetworkInterfaceBytesTotalPersec
       WScript.Echo "NetworkInterfaceCurrentBandwidth=" & NetworkInterfaceCurrentBandwidth
       WScript.Echo "NetworkInterfaceName=" & NetworkInterfaceName
       WScript.Echo "NetworkInterfacePacketsPersec=" & NetworkInterfacePacketsPersec
       WScript.Echo "NetworkInterfacePacketsReceivedDiscarded=" & NetworkInterfacePacketsReceivedDiscarded
       WScript.Echo "NetworkInterfacePacketsReceivedErrors=" & NetworkInterfacePacketsReceivedErrors
       WScript.Echo "NetworkInterfacePacketsReceivedNonUnicastPersec=" & NetworkInterfacePacketsReceivedNonUnicastPersec
       WScript.Echo "NetworkInterfacePacketsReceivedPersec=" & NetworkInterfacePacketsReceivedPersec
       WScript.Echo "NetworkInterfacePacketsReceivedUnicastPersec=" & NetworkInterfacePacketsReceivedUnicastPersec
       WScript.Echo "NetworkInterfacePacketsReceivedUnknown=" & NetworkInterfacePacketsReceivedUnknown
       WScript.Echo "NetworkInterfacePacketsSentNonUnicastPersec=" & NetworkInterfacePacketsSentNonUnicastPersec
       WScript.Echo "NetworkInterfacePacketsSentPersec=" & NetworkInterfacePacketsSentPersec
       WScript.Echo "NetworkInterfacePacketsSentUnicastPersec=" & NetworkInterfacePacketsSentUnicastPersec
       WScript.Echo "NetworkInterfaceTimestamp_Object=" & NetworkInterfaceTimestamp_Object
       WScript.Echo "NetworkInterfaceTimestamp_PerfTime=" & NetworkInterfaceTimestamp_PerfTime
       WScript.Echo "NetworkInterfaceTimestamp_Sys100NS=" & NetworkInterfaceTimestamp_Sys100NS
    end sub

end class

class  ObjNetworkAdaptorIOInfo
    Public ObjNetworkAdaptorIOs
    private sub class_Initialize
        ObjNetworkAdaptorIOs = Array()
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub


    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_NetworkInterface",,48)
        
        Dim objItem 'as Win32_PerfRawData_Tcpip_NetworkInterface
        For Each objItem in colItems
            size = UBound(ObjNetworkAdaptorIOs)
            if size < 0 Then
                size = 0
            end if
            Redim Preserve ObjNetworkAdaptorIOs(size + 1)

            set tmp = New ObjNetworkAdaptorIO
            tmp.NetworkInterfaceBytesReceivedPersec = objItem.BytesReceivedPersec
            tmp.NetworkInterfaceBytesSentPersec = objItem.BytesSentPersec
            tmp.NetworkInterfaceBytesTotalPersec = objItem.BytesTotalPersec
            tmp.NetworkInterfaceCurrentBandwidth = objItem.CurrentBandwidth
            tmp.NetworkInterfaceName = objItem.Name
            tmp.NetworkInterfacePacketsPersec = objItem.PacketsReceivedNonUnicastPersec
            tmp.NetworkInterfacePacketsReceivedDiscarded = objItem.PacketsReceivedDiscarded
            tmp.NetworkInterfacePacketsReceivedErrors = objItem.PacketsReceivedErrors
            tmp.NetworkInterfacePacketsReceivedNonUnicastPersec = objItem.PacketsReceivedNonUnicastPersec
            tmp.NetworkInterfacePacketsReceivedPersec = objItem.PacketsReceivedPersec
            tmp.NetworkInterfacePacketsReceivedUnicastPersec = objItem.PacketsReceivedUnicastPersec
            tmp.NetworkInterfacePacketsReceivedUnknown = objItem.PacketsReceivedUnknown
            tmp.NetworkInterfacePacketsSentNonUnicastPersec = objItem.PacketsSentNonUnicastPersec
            tmp.NetworkInterfacePacketsSentPersec = objItem.PacketsSentPersec
            tmp.NetworkInterfacePacketsSentUnicastPersec = objItem.PacketsSentUnicastPersec
            tmp.NetworkInterfaceTimestamp_Object = objItem.Timestamp_Object
            tmp.NetworkInterfaceTimestamp_PerfTime = objItem.Timestamp_PerfTime
            tmp.NetworkInterfaceTimestamp_Sys100NS = objItem.Timestamp_Sys100NS
            set ObjNetworkAdaptorIOs(size) = tmp
        Next
    end sub

    sub Print()
        For Each tmp In ObjNetworkAdaptorIOs
            if not IsEmpty(tmp) Then
                call tmp.Print
            end if
        Next 
    end sub
end class

class ObjNetworkAdaptorInfo
    Public ObjNetworkAdaptors
    ' Public mapNameInterfaceObj
    private sub class_Initialize
      
        ObjNetworkAdaptors = Array()
        ' Set mapNameInterfaceObj = CreateObject("Scripting.Dictionary")
    end sub

    private sub class_Terminate
      
    end sub

 

    sub Collect()
        ' Call CollectInterface
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)
        
        Dim objItem 'as Win32_NetworkAdapter
        For Each objItem in colItems    
            size = UBound(ObjNetworkAdaptors)
            if size < 0 Then
                size = 0
            end if
            Redim Preserve ObjNetworkAdaptors(size + 1)
            set tmp = New ObjNetworkAdaptor
            tmp.NetworkAdaptorAdapterType = objItem.AdapterType
            tmp.NetworkAdaptorAdapterTypeId = objItem.AdapterTypeId
            tmp.NetworkAdaptorCaption = objItem.Caption
            tmp.NetworkAdaptorDescription = objItem.Description
            tmp.NetworkAdaptorMACAddress = objItem.MACAddress
            tmp.NetworkAdaptorManufacturer = objItem.Manufacturer
            tmp.NetworkAdaptorMaxSpeed = objItem.MaxSpeed
            tmp.NetworkAdaptorName = objItem.Name
            tmp.NetworkAdaptorNetConnectionID = objItem.NetConnectionID
            tmp.NetworkAdaptorNetConnectionStatus = objItem.NetConnectionStatus

            ' if objItem.NetEnabled = True Then
            '     tmp.NetworkAdaptorNetEnabled = "up"
            ' else
            '     tmp.NetworkAdaptorNetEnabled = "down"
            ' end if

            ' tmp.NetworkAdaptorNetEnabled = objItem.NetEnabled
            tmp.NetworkAdaptorNetworkAddresses = objItem.NetworkAddresses
            tmp.NetworkAdaptorPermanentAddress = objItem.PermanentAddress
            ' tmp.NetworkAdaptorPhysicalAdapter = objItem.PhysicalAdapter
            tmp.NetworkAdaptorPNPDeviceID = objItem.PNPDeviceID
            tmp.NetworkAdaptorProductName = objItem.ProductName
            tmp.NetworkAdaptorServiceName = objItem.ServiceName
            tmp.NetworkAdaptorSpeed = objItem.Speed
            tmp.NetworkAdaptorStatus = objItem.Status
            tmp.NetworkAdaptorStatusInfo = objItem.StatusInfo

            ' if mapNameInterfaceObj.Exists(tmp.NetworkAdaptorName) Then
            '     set obj = mapNameInterfaceObj.Item(tmp.NetworkAdaptorName) 
            '     tmp.NetworkAdaptorBytesReceivedPersec = obj.NetworkInterfaceBytesReceivedPersec 
            '     tmp.NetworkAdaptorBytesSentPersec = obj.NetworkInterfaceBytesSentPersec 
            '     tmp.NetworkAdaptorBytesTotalPersec = obj.NetworkInterfaceBytesTotalPersec 
            '     tmp.NetworkAdaptorCurrentBandwidth = obj.NetworkInterfaceCurrentBandwidth 
            '     tmp.NetworkAdaptorFrequency_Object = obj.NetworkInterfaceTimestamp_Object 
            '     tmp.NetworkAdaptorFrequency_PerfTime = obj.NetworkInterfaceTimestamp_PerfTime 
            '     tmp.NetworkAdaptorFrequency_Sys100NS = obj.NetworkInterfaceTimestamp_Sys100NS
            '     tmp.NetworkAdaptorPacketsPersec = obj.NetworkInterfacePacketsPersec 
            '     tmp.NetworkAdaptorPacketsReceivedDiscarded = obj.NetworkInterfacePacketsReceivedDiscarded 
            '     tmp.NetworkAdaptorPacketsReceivedErrors = obj.NetworkInterfacePacketsReceivedErrors 
            '     tmp.NetworkAdaptorPacketsReceivedNonUnicastPersec = obj.NetworkInterfacePacketsReceivedNonUnicastPersec 
            '     tmp.NetworkAdaptorPacketsReceivedPersec = obj.NetworkInterfacePacketsReceivedPersec 
            '     tmp.NetworkAdaptorPacketsReceivedUnicastPersec = obj.NetworkInterfacePacketsReceivedUnicastPersec 
            '     tmp.NetworkAdaptorPacketsReceivedUnknown = obj.NetworkInterfacePacketsReceivedUnknown 
            '     tmp.NetworkAdaptorPacketsSentNonUnicastPersec = obj.NetworkInterfacePacketsSentNonUnicastPersec 
            '     tmp.NetworkAdaptorPacketsSentPersec = obj.NetworkInterfacePacketsSentPersec 
            '     tmp.NetworkAdaptorPacketsSentUnicastPersec = obj.NetworkInterfacePacketsSentUnicastPersec 

            ' end if
            set ObjNetworkAdaptors(size) = tmp
        Next
    end sub

    sub Print()
        For Each tmp In ObjNetworkAdaptors
            if not IsEmpty(tmp) Then
                call tmp.Print
            end if
        Next 
    end sub

end class

class ObjService
    Public ServiceAcceptPause
    Public ServiceAcceptStop
    Public ServiceCaption
    Public ServiceCheckPoint
    Public ServiceCreationClassName
    Public ServiceDelayedAutoStart
    Public ServiceDescription
    Public ServiceDesktopInteract
    Public ServiceDisplayName
    Public ServiceErrorControl
    Public ServiceExitCode
    Public ServiceInstallDate
    Public ServiceName
    Public ServicePathName
    Public ServiceProcessId
    Public ServiceServiceSpecificExitCode
    Public ServiceServiceType
    Public ServiceStarted
    Public ServiceStartMode
    Public ServiceStartName
    Public ServiceState
    Public ServiceStatus
    Public ServiceSystemCreationClassName
    Public ServiceSystemName
    Public ServiceTagId
    Public ServiceWaitHint
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        
    end sub

    sub Print()
        WScript.Echo   "ServiceAcceptPause=" & ServiceAcceptPause
        WScript.Echo   "ServiceAcceptStop=" &     ServiceAcceptStop
        WScript.Echo   "ServiceCaption=" &     ServiceCaption
        WScript.Echo   "ServiceCheckPoint=" &     ServiceCheckPoint
        WScript.Echo   "ServiceCreationClassName=" &     ServiceCreationClassName
        WScript.Echo   "ServiceDelayedAutoStart=" &     ServiceDelayedAutoStart
        WScript.Echo   "ServiceDescription=" &     ServiceDescription
        WScript.Echo   "ServiceDesktopInteract=" &     ServiceDesktopInteract
        WScript.Echo   "ServiceDisplayName=" &     ServiceDisplayName
        WScript.Echo   "ServiceErrorControl=" &     ServiceErrorControl
        WScript.Echo   "ServiceExitCode=" &     ServiceExitCode
        WScript.Echo   "ServiceInstallDate=" &     ServiceInstallDate
        WScript.Echo   "ServiceName=" &     ServiceName
        WScript.Echo   "ServicePathName=" &     ServicePathName
        WScript.Echo   "ServiceProcessId=" &     ServiceProcessId
        WScript.Echo   "ServiceServiceSpecificExitCode=" &     ServiceServiceSpecificExitCode
        WScript.Echo   "ServiceServiceType=" &     ServiceServiceType
        WScript.Echo   "ServiceStarted=" &     ServiceStarted
        WScript.Echo   "ServiceStartMode=" &     ServiceStartMode
        WScript.Echo   "ServiceStartName=" &     ServiceStartName
        WScript.Echo   "ServiceState=" &     ServiceState
        WScript.Echo   "ServiceStatus=" &     ServiceStatus
        WScript.Echo   "ServiceSystemCreationClassName=" &     ServiceSystemCreationClassName
        WScript.Echo   "ServiceSystemName=" &     ServiceSystemName
        WScript.Echo   "ServiceTagId=" &     ServiceTagId
        WScript.Echo   "ServiceWaitHint=" &     ServiceWaitHint
    end sub

end class


class ObjServiceInfo
    Public ObjServices
    private sub class_Initialize
        ' Called automatically when class is created
        ObjServices = Array()
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Service",,48)

        Dim objItem 'as Win32_Service
        For Each objItem in colItems
            size = UBound(ObjServices)
            if size < 0 Then
                size = 0
            end if
            Redim Preserve ObjServices(size + 1)
            set tmp = New ObjService
            tmp.ServiceAcceptPause = objItem.AcceptPause 
            tmp.ServiceAcceptStop = objItem.AcceptStop 
            tmp.ServiceCaption = objItem.Caption 
            tmp.ServiceCheckPoint = objItem.CheckPoint 
            tmp.ServiceCreationClassName = objItem.CreationClassName 
            tmp.ServiceDelayedAutoStart = objItem.DelayedAutoStart 
            tmp.ServiceDescription = objItem.Description 
            tmp.ServiceDesktopInteract = objItem.DesktopInteract 
            tmp.ServiceDisplayName = objItem.DisplayName 
            tmp.ServiceErrorControl = objItem.ErrorControl 
            tmp.ServiceExitCode = objItem.ExitCode 
            tmp.ServiceInstallDate = objItem.InstallDate 
            tmp.ServiceName = objItem.Name 
            tmp.ServicePathName = objItem.PathName 
            tmp.ServiceProcessId = objItem.ProcessId 
            tmp.ServiceServiceSpecificExitCode = objItem.ServiceSpecificExitCode 
            tmp.ServiceServiceType = objItem.ServiceType 
            tmp.ServiceStarted = objItem.Started 
            tmp.ServiceStartMode = objItem.StartMode 
            tmp.ServiceStartName = objItem.StartName 
            tmp.ServiceState = objItem.State 
            tmp.ServiceStatus = objItem.Status 
            tmp.ServiceSystemCreationClassName = objItem.SystemCreationClassName 
            tmp.ServiceSystemName = objItem.SystemName 
            tmp.ServiceTagId = objItem.TagId 
            tmp.ServiceWaitHint = objItem.WaitHint 
            set ObjServices(size) = tmp
        Next
    end sub


    sub Print()
        Call ObjServices(0).Print
        ' For Each tmp In ObjServices
        '     if not IsEmpty(tmp) Then
        '         call tmp.Print
        '     end if
        ' Next 
    end sub
end class


class ObjSchTaskInfo
    Public ObjSchTasks
    Public  rowCont 
    private sub class_Initialize
        
        ObjSchTasks = Array()
        rowCont = 29
    end sub

    private sub class_Terminate
        
    end sub

    sub Collect()
        strResult = createobject("wscript.shell").exec("SCHTASKS /Query /FO LIST /V").StdOut.ReadAll
    
        arrSplitStr =  Split(strResult, vbCrLf)

        Dim  arrSplit() ,curentfolder
        ReDim arrSplit(rowCont, 1) 
        currentRows = 0
        isFirst = 1
        for i = 0 to UBound(arrSplitStr) step 1
            ' ??????????,???????
            if i <> 0 Then
                size  = UBound( arrSplit, 2) 
                if arrSplitStr(i) = "" Then
                    ReDim Preserve arrSplit(rowCont, size+1) 
                    currentRows = 0
                    isFirst = 1
                else
                    tmp = Split(arrSplitStr(i), ":", 2)
                    if isFirst = 1 Then
                        isFirst = 0
                        if tmp(0) = "?????" Then
                            curentfolder = tmp(1)
                        end if
                        arrSplit(currentRows, size-1) = curentfolder
                    else   
                        
                        arrSplit(currentRows, size-1) = tmp(1)
                    end if 
                    currentRows = currentRows + 1    
                end if
            end if
        next
        
      
        for i = 0 to  UBound( arrSplit, 2)  step 1
            size = UBound(ObjSchTasks)
            if size < 0 Then
                size = 0
            end if
            Redim Preserve ObjSchTasks(size + 1)
            set tmp = New ObjSchTask
             tmp.SchTaskName = arrSplit(2, i)
             tmp.SchTaskNextRuntime = arrSplit(3, i)
             tmp.SchTaskMode = arrSplit(4, i)
             tmp.SchTaskLastRuntime = arrSplit(4, i)
             tmp.SchTaskLastResult= arrSplit(7, i)
             tmp.SchTask = arrSplit(9, i)
             tmp.SchTaskStatus = arrSplit(12, i)
            set ObjSchTasks(size) = tmp
        next
    end sub

    sub Print()
        For Each tmp In ObjSchTasks
            if not IsEmpty(tmp) Then
                call tmp.Print
            end if
        Next 
    end sub
end class



class ObjSchTask
    Public SchTaskName 
    Public SchTaskNextRuntime
    Public SchTaskLastRuntime 
    Public SchTaskLastResult 
    Public SchTaskMode 
    Public SchTaskStatus 
    Public SchTask
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        
    end sub

    sub Print()
       WScript.Echo "SchTaskName=" & SchTaskName 
       WScript.Echo "SchTaskNextRuntim=" & SchTaskNextRuntime
       WScript.Echo "SchTaskLastRuntime=" & SchTaskLastRuntime 
       WScript.Echo "SchTaskLastResult=" & SchTaskLastResult 
       WScript.Echo "SchTaskName=" & SchTaskName 
       WScript.Echo "SchTaskMode=" & SchTaskMode 
       WScript.Echo "SchTaskStatus=" & SchTaskStatus 
       WScript.Echo "SchTask=" & SchTask
    end sub
end class


class ObjTcpInfo
    Public TcpConnectionFailures
    Public TcpConnectionsActive
    Public TcpConnectionsEstablished
    Public TcpConnectionsPassive
    Public TcpConnectionsReset
    Public TcpSegmentsPersec
    Public TcpSegmentsReceivedPersec
    Public TcpSegmentsRetransmittedPersec
    Public TcpSegmentsSentPersec

    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_Tcpip_TCPv4",,48)
        
        Dim objItem 'as Win32_PerfRawData_Tcpip_TCPv4
        For Each objItem in colItems
            TcpConnectionFailures = objItem.ConnectionFailures
            TcpConnectionsActive = objItem.ConnectionsActive
            TcpConnectionsEstablished = objItem.ConnectionsEstablished
            TcpConnectionsPassive = objItem.ConnectionsPassive
            TcpConnectionsReset = objItem.ConnectionsReset
            TcpSegmentsPersec = objItem.SegmentsPersec
            TcpSegmentsReceivedPersec = objItem.SegmentsReceivedPersec
            TcpSegmentsRetransmittedPersec = objItem.SegmentsRetransmittedPersec
            TcpSegmentsSentPersec = objItem.SegmentsSentPersec
        Next
    end sub

    sub Print()
        WScript.Echo "TcpConnectionFailures=" & TcpConnectionFailures
        WScript.Echo "TcpConnectionsActive=" & TcpConnectionsActive
        WScript.Echo "TcpConnectionsEstablished=" & TcpConnectionsEstablished
        WScript.Echo "TcpConnectionsPassive=" & TcpConnectionsPassive
        WScript.Echo "TcpConnectionsReset=" & TcpConnectionsReset
        WScript.Echo "TcpSegmentsPersec=" & TcpSegmentsPersec
        WScript.Echo "TcpSegmentsReceivedPersec=" & TcpSegmentsReceivedPersec
        WScript.Echo "TcpSegmentsRetransmittedPersec=" & TcpSegmentsRetransmittedPersec
        WScript.Echo "TcpSegmentsSentPersec=" & TcpSegmentsSentPersec
    end sub

end class



class ObjAccontInfo
    Public ObjAcconts
    private sub class_Initialize
        ObjAcconts = Array()
    end sub

    private sub class_Terminate
       
    end sub

    sub Collect()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Account",,48)
        For Each objItem in colItems
            size = UBound(ObjAcconts)
            if size < 0 Then
                size = 0
            end if
            Redim Preserve ObjAcconts(size + 1)
            set tmp = New ObjAccount
            
            tmp.AccountCaption = objItem.Caption  
            tmp.AccountDescription = objItem.Description  
            tmp.AccountDomain = objItem.Domain  
            tmp.AccountInstallDate = objItem.InstallDate  
            tmp.AccountLocalAccount = objItem.LocalAccount  
            tmp.AccountName = objItem.Name  
            tmp.AccountSID = objItem.SID  
            tmp.AccountSIDType = objItem.SIDType  
            tmp.AccountStatus = objItem.Status  

            set ObjAcconts(size) = tmp
        Next
    end sub

    sub Print()
        For Each tmp In ObjAcconts
            if not IsEmpty(tmp) Then
                call tmp.Print
            end if
        Next 
    end sub
   
end class


class ObjAccount
    Public AccountCaption
    Public AccountDescription
    Public AccountDomain
    Public AccountInstallDate
    Public AccountLocalAccount
    Public AccountName
    Public AccountSID
    Public AccountSIDType
    Public AccountStatus
    private sub class_Initialize
        ' Called automatically when class is created
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Print()
        WScript.Echo "AccountCaption=" &AccountCaption
        WScript.Echo "AccountDescription=" &  AccountDescription
        WScript.Echo "AccountDomain=" &  AccountDomain
        WScript.Echo "AccountInstallDate=" &  AccountInstallDate
        WScript.Echo "AccountLocalAccount=" &  AccountLocalAccount
        WScript.Echo "AccountName=" & AccountName
        WScript.Echo "AccountSID=" &  AccountSID
        WScript.Echo "AccountSIDType=" & AccountSIDType
        WScript.Echo "AccountStatus=" &  AccountStatus
    end sub
end class




Dim oobj
For Each oobj In Array(New ObjSystemInfo,New ObjProcessInfo,New ObjCpuInfo,New ObjDiskInfo, New ObjNetworkAdaptorInfo, New ObjNetworkAdaptorIOInfo,New ObjSchTaskInfo, New ObjTcpInfo,New ObjAccontInfo)
    call oobj.Collect
    call oobj.Print
Next 
