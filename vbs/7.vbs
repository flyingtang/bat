

' 命令行读取的数据

' sResults = oExec.Stdout.ReadAll

' WScript.Echo strResult
' dim str
' set objtl = New ObjTaskList
' objtl.ProcessProcessId = objPerfItem.IDProcess
' objtl.ProcessPercentProcessorTime = objPerfItem.PercentProcessorTime
' objtl.ProcessWorkingSet = objPerfItem.WorkingSet
' ' 有两个ID 为0 的东东
' if  objPerfItem.IDProcess <> 0 Then 
'     mapIdProcessPerfObj.Add objPerfItem.IDProcess, tmp
' end if 
' WScript.Echo "========="
' dim str, row,column,index
' willRow = 1
' column = 9
' currentrow = 0
' currentColunm = 0
' dim myArr
' myArr = Array(2,2)
' ReDim Preserve  myArr(2 To 2, 2 To 3)

' Split(strResult, vbCrLf)

' for each str in  Split(strResult, vbCrLf)
'     ' WScript.Echo str
'     if str = "" Then
'         currentrow = willRow
'         currentColunm = 0
'         willRow = willRow+1
'         WScript.Echo willRow, column
'         ReDim Preserve MyArr(willRow, column)
'     Else
'         MyArr(currentrow, currentColunm) = str
'         currentColunm = currentColunm + 1
'     end if 
' next

' For Each ss In myArr
'     for each s in ss
'         WScript.Echo "s===" & s
'     next
' Next ' Element

' WScript.Echo "===========s========"
' arrSplitStr =  Split(strResult, vbCrLf)

' for i = 0 to UBound(arrSplitStr) step 1
'     if i <> 0 Then
'         WScript.Echo arrSplitStr(i), i
'     end if
' next
' WScript.Echo "===========end========" & UBound(arrSplitStr) & LBound(arrSplitStr)

' Dim intTotal, intSpaceCount, intcoumnCount, rowCount
' intTotal=0
' intSpaceCount=0
' coumnCount=9
' isFirestRow = 1


' For Each str In arrSplitStr
'     if isFirestRow= 1 && str = "" Then
'     else
'          if str = "" Then
'             isFirestRow = 0
'             intSpaceCount = intSpaceCount+1
'         end if
'     end if 
'     intTotal = intTotal + 1  
' Next ' Element
' rowCount =(intTotal - spaceCount ) / coumnCount

' WScript.Echo intTotal, intSpaceCount, rowCount,coumnCount

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
            arrSplit(currentRows, size-1) = (Split(arrSplitStr(i), ":"))(1)
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
