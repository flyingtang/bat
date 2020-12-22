Dim Dic
Set Dic = CreateObject("Scripting.Dictionary")
' Dic.Add "Name", "Sirrah" '向Dictionary对象中添加键值对
' Dic.Add "Age", 23



' WScript.Echo  DIc.Item("Age")'判断键是否存在　



Class ProcessObject
    Public sttProcessName
    Public intProcessID
    Public intProcessMemUsed
    Public intCpuUsed
    Public Sub Class_Initialize
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate
        ' Called automatically when all references to class instance are removed
    End Sub


End Class


set objProcess = New ProcessObject
objProcess.sttProcessName="xiaogang2"
objProcess.intProcessID="xiaogang1"
objProcess.intProcessMemUsed="xiaogan1"
objProcess.intProcessMemUsed="xiaogang1"

Dic.Add "xiaogang", objProcess

one  = Dic.Item("xiaogang").sttProcessName

WScript.Echo  one