On Error Resume Next
Set oShell = CreateObject("Wscript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
oShell.Run "REGEDIT.EXE /S uninst.reg", ,true
sDesktop = oShell.SpecialFolders("Desktop")
If oFSO.FileExists (sDesktop & "\Htaedit.lnk") Then oFSO.DeleteFile (sDesktop & "\Htaedit.lnk")
If oFSO.FileExists (sDesktop & "\Vbsedit.lnk") Then oFSO.DeleteFile (sDesktop & "\Vbsedit.lnk")
MsgBox "注册表清理完毕，请手工删除软件所有文件夹！",64 