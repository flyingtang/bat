
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

        Dim  arrSplit() ,curentfolder
        ReDim arrSplit(rowCont, 1) 
        currentRows = 0
        isFirst = 1
        for i = 0 to UBound(arrSplitStr) step 1
            ' 第一行是空行,所以去除
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
                        if tmp(0) = "文件夹" Then
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
            set ObjSchTasks(size) = tmp
        next
    end sub

    sub Print()
        Call ObjSchTasks(0).Print
        ' For Each tmp In ObjServices
        '     if not IsEmpty(tmp) Then
        '         call tmp.Print
        '     end if
        ' Next 
    end sub
end class



class ObjSchTask
    Public SchTaskName 
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
    end sub
end class



set objSyInfo = New ObjSchTaskInfo
objSyInfo.Collect
objSyInfo.Print