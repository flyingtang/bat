class ObjSchTaskInfo
    Public ObjSchTasks
    Public  rowCont 
    private sub class_Initialize
        ' Called automatically when class is created
        ObjSchTasks = Array()
        rowCont = 29
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
    end sub

    sub Collect()
        strResult = createobject("wscript.shell").exec("SCHTASKS /Query /FO LIST /V").StdOut.ReadAll
        ' WScript.Echo strResult
        arrSplitStr =  Split(strResult, vbCrLf)

        Dim  arrSplit() ,curentfolder, arrSplitKey()
        ReDim arrSplit(rowCont, 1) 
        ReDim arrSplitKey(rowCont, 1) 
        currentRows = 0
        isFirst = 1
        for i = 0 to UBound(arrSplitStr) step 1
            ' ��һ���ǿ���,����ȥ��
            if i <> 0 Then
                size  = UBound(arrSplit, 2) 
                if arrSplitStr(i) = "" Then
                    ReDim Preserve arrSplit(rowCont, size+1) 
                    ReDim Preserve arrSplitKey(rowCont, size+1) 
                    currentRows = 0
                    isFirst = 1
                else
                    tmp = Split(arrSplitStr(i), ":", 2)
                    if isFirst = 1 Then
                        isFirst = 0
                        if tmp(0) = "�ļ���" Then
                            curentfolder = tmp(1)
                        end if
                        arrSplit(currentRows, size-1) = curentfolder
                        arrSplitKey(currentRows, size-1) = tmp(0)
                    else   
                        arrSplit(currentRows, size-1) = tmp(1)
                        arrSplitKey(currentRows, size-1) = tmp(0)
                    end if 
                    currentRows = currentRows + 1    
                end if
            end if
        next
        
        intCoumnSize = UBound(arrSplitKey, 2)
        for i = 0 to  intCoumnSize  step 1
            if i < intCoumnSize -1 Then
                size = UBound(ObjSchTasks)
                if size < 0 Then
                    size = 0
                end if
                Redim Preserve ObjSchTasks(size + 1)
                set tmp = New ObjSchTask

                for j = 0 to rowCont step 1
                    strKeyName = arrSplitKey(j, i)
                    select case strKeyName
                        case "������"
                            tmp.SchTaskName = arrSplit(j, i)
                        case "�´�����ʱ��"
                            tmp.SchTaskNextRuntime = arrSplit(j, i)
                        case "ģʽ"
                            tmp.SchTaskMode = arrSplit(j, i)
                        case "�ϴ�����ʱ��"
                            tmp.SchTaskLastRuntime = arrSplit(j, i)
                        case "�ϴν��"
                            tmp.SchTaskLastResult = arrSplit(j, i)
                        case "Ҫ���е�����"
                            tmp.SchTask = arrSplit(j, i)
                        case "�ƻ�����״̬"
                            tmp.SchTaskStatus = arrSplit(j, i)
                        case "�ƻ�������"
                            tmp.SchTaskType = arrSplit(j, i)
                        case else
                            
                    end select
                next
                set ObjSchTasks(size) = tmp
            end if 
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
    Public SchTaskType
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
       WScript.Echo "SchTaskType=" & SchTaskType
    end sub
end class

set objSyInfo = New ObjSchTaskInfo
objSyInfo.Collect
objSyInfo.Print