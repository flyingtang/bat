Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Wscript.Shell")
sPath = WScript.ScriptFullName
sCurDir = oFSO.GetParentFolderName(sPath)
sHlpFile = sCurDir & "\Script全集.chm"
sWeb = "http://www.baidu.com/baidu?word=%1"
sHta = Chr(34)& sCurDir & "\Htaedit.exe" & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
sVbs = Chr(34)& sCurDir & "\Vbsedit.exe" & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
oShell.RegWrite "HKEY_CLASSES_ROOT\htafile\Shell\Edit with HtaEdit\command\", sHta, "REG_SZ"
oShell.RegWrite "HKEY_CLASSES_ROOT\vbsfile\Shell\Edit with VbsEdit\command\", sVbs, "REG_SZ"
oShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Adersoft\Htaedit\INSTALLDIR", sCurDir, "REG_SZ"
oShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Adersoft\Vbsedit\INSTALLDIR", sCurDir, "REG_SZ"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Help\File", sHlpFile, "REG_SZ"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Help\Web", sWeb, "REG_SZ"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Background", &H00FFFFFF, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Comments", &H00008000, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Fields", &H000000FF, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Keywords", &H00FF0000, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\LineNumbers", &H00808080, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\NormalText", &H00000000, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Numbers", &H00000080, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\SelBackground", &H00A00000, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\SelText", &H00FFFFFF, "REG_DWORD"
oShell.RegWrite "HKEY_CURRENT_USER\Software\Adersoft\VbsEdit\Colors\Strings", &H00FF00FF, "REG_DWORD"
rtnVal = MsgBox ("绿化完成，是否创建快捷方式?",36,"是否创建快捷方式")
If rtnVal = vbYES Then
	sDesktop = oShell.SpecialFolders("Desktop")
	set oShellLink1 = oShell.CreateShortcut(sDesktop & "\HtaEdit.lnk")
	oShellLink1.TargetPath = Chr(34)& sCurDir & "\Htaedit.exe" & Chr(34)
	oShellLink1.WorkingDirectory = Chr(34)& sCurDir & Chr(34)
	oShellLink1.Save
	set oShellLink2 = oShell.CreateShortcut(sDesktop & "\VbsEdit.lnk")
	oShellLink2.TargetPath = Chr(34)& sCurDir & "\Vbsedit.exe" & Chr(34)
	oShellLink2.WorkingDirectory = Chr(34) & sCurDir & Chr(34)
	oShellLink2.Save
End If
