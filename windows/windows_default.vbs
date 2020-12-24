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
            end if 
            mapIdProcessObj.Add arrProcessObj(count).ProcessProcessId, arrProcessObj(count)
        Next

        ' 命令行读取的数据
        set oExec = createobject("wscript.shell").exec("TASKLIST /V /FO LIST")
        sResults = oExec.Stdout.ReadAll
        count = 0
        do while oExec.Stdout.AtEndOfLine <> True
            line = oExec.ReadLine
            count = count + 1
            WScript.Echo "count: " & count
        loop


    end sub

    sub PrintProcess()
        for each processObj in arrProcessObj
            if IsEmpty(processObj) Then
            else
                processObj.ToString
                if processObj.ProcessProcessId <> 0 Then
                    WScript.Echo "processObj.ProcessProcessId=== " & processObj.ProcessProcessId 
                    WScript.Echo "ProcesxxxxxxxxxxxxxxxxxsWorkingSet" & mapIdProcessObj.Item(processObj.ProcessProcessId).ProcessWorkingSet
                end if 
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
    Public ProcessWorkingSetSize
    Public ProcessPercentProcessorTime
    Public ProcessWorkingSet
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
        WScript.Echo "ProcessWorkingSetSize:"  & ProcessWorkingSetSize
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






set  objProInfo = New ObjProcessInfo
WScript.Echo objProInfo.mapIdProcessObj.Item("2584")
objProInfo.CollectProcess
objProInfo.PrintProcess