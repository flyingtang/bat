Class ObjProcessInfo
    Public arrProcessObj 
    Private intProcessCount
    Public mapIdProcessObj
    Private Sub Class_Initialize
        arrProcessObj = Array()
        intProcessCount = 0
        Set mapIdProcessObj = CreateObject("Scripting.Dictionary")
        WScript.Echo " Called automatically when class is created"
    End Sub

    Private Sub Class_Terminate
        ' Called automatically when all references to class instance are removed
        WScript.Echo " Called automatically when all references to class instance are removed"
      
    End Sub



    sub CollectProcess()
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

        ' 动态信息
        Set perfColItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Process",,48)
        Set mapIdProcessPerfObj = CreateObject("Scripting.Dictionary")
        Dim objPerfItem 'as Win32_PerfFormattedData_PerfProc_Process
        For Each objPerfItem in perfColItems
            set tmp = New ObjPerfProcess
            tmp.ProcessProcessId = objPerfItem.IDProcess
            tmp.ProcessPercentProcessorTime = objPerfItem.PercentProcessorTime
            tmp.ProcessWorkingSet = objPerfItem.WorkingSet
            ' 有两个ID 为0 的东东
            if  objPerfItem.IDProcess <> 0 Then 
                mapIdProcessPerfObj.Add objPerfItem.IDProcess, tmp
            end if 
        Next


        ' 命令行读取的数据
        strResult = createobject("wscript.shell").exec("TASKLIST /V /FO LIST").StdOut.ReadAll
        arrSplitStr =  Split(strResult, vbCrLf)
        Dim  arrSplit() 
        ReDim arrSplit(9, 1) 
        currentRows = 0
        for i = 0 to UBound(arrSplitStr) step 1
            ' 第一行是空行,所以去除
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
        
        Set mapProcessIdCmdnTask = CreateObject("Scripting.Dictionary")
        for i = 0 to  UBound( arrSplit, 2)  step 1
            set tmp = New CmdTask
            tmp.pid=CInt(arrSplit(1, i))
            tmp.memused=arrSplit(4, i)
            tmp.cpuused=arrSplit(7, i)  
            if Not mapProcessIdCmdnTask.Exists(tmp.pid) Then 
                  mapProcessIdCmdnTask.Add tmp.pid, tmp
            end if 
        next

        ' 基本信息
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
                set tmp =  mapIdProcessPerfObj.Item(objItem.ProcessId)   
                arrProcessObj(count).ProcessPercentProcessorTime = tmp.ProcessPercentProcessorTime
                arrProcessObj(count).ProcessWorkingSet = tmp.ProcessWorkingSet
                'TODO 问题
               if  mapProcessIdCmdnTask.Exists(cint(objItem.ProcessId)) Then 
                set ttmp = mapProcessIdCmdnTask.Item(objItem.ProcessId)
                arrProcessObj(count).ProcessMemused = ttmp.memused
                arrProcessObj(count).ProcessCpuused = ttmp.cpuused
               end if 
            end if 
            mapIdProcessObj.Add arrProcessObj(count).ProcessProcessId, arrProcessObj(count)
        Next
    end sub

    sub PrintProcess()
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

' cmd 任务对象
class CmdTask
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




set  objProInfo = New ObjProcessInfo
objProInfo.CollectProcess
objProInfo.PrintProcess