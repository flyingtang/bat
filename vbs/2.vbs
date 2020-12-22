Option Explicit
Dim strComputer
Dim objWMIService
Dim colProcesses strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery _
("Select * from Win32_Process") 'for each objProcess in colProcesses
' If InStr(LCase(objProcess.Name), "notepad.exe") = 1 Then
' sngProcessTime = (CSng(objProcess.KernelModeTime) + CSng(objProcess.UserModeTime)) / 10000000
' Wscript.Echo "CPU Usage " & count & "=" & sngProcessTime
' Wscript.Echo "Process ID " & count & "=" & objProcess.ProcessID & ":" & _
' objProcess.Name
' end if
'Next
Dim pn
Dim pinfo()
Dim ctr
ctr = 0
Dim obj
'判断有多少个进程
For Each obj In colProcesses
'pn = pn + obj.Description + vbcrlf
ctr = ctr + 1
Next
ReDim pinfo(ctr, 1)
ctr = 0
'获取进程的信息
For Each obj In colProcesses
'pn = pn + obj.Description + vbcrlf
ctr = ctr + 1
pinfo(ctr, 0) = obj.Description
pinfo(ctr, 1) = CStr((CSng(obj.KernelModeTime) + CSng(obj.UserModeTime)) / 10000000)
Next
'以下是对进程的冒泡排序
Dim i, j, temp
For i = 1 To ctr - 1 Step 1
For j = ctr - 1 To i Step -1
'msgbox pinfo(j - 1, 1),,"1"
'msgbox pinfo(j, 1),,"2"
If CDbl(pinfo(j - 1, 1)) < CDbl(pinfo(j, 1)) Then
temp = pinfo(j - 1, 1)
pinfo(j - 1, 1) = pinfo(j, 1)
pinfo(j, 1) = temp
temp = pinfo(j - 1, 0)
pinfo(j - 1, 0) = pinfo(j, 0)
pinfo(j, 0) = temp
End If
Next
Next
pn = ""
For ctr = 1 To 10
pn = pn & pinfo(ctr, 0) & vbCrLf
Next
MsgBox pn