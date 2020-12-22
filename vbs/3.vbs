function GetProcessInfo()
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Process",,48)

	Dim objItem 'as Win32_Process
	For Each objItem in colItems
		WScript.Echo "ProcessProcessId: " & objItem.ProcessId
		WScript.Echo "ProcessName: " & objItem.Name
		WScript.Echo "ProcessParentProcessId: " & objItem.ParentProcessId
		WScript.Echo "ProcessThreadCount: " & objItem.ThreadCount
		WScript.Echo "ProcessHandleCount: " & objItem.HandleCount
		WScript.Echo "ProcessCpuTime:"  &  (objItem.KernelModeTime + objItem.UserModeTime) / 10000000
        WScript.Echo "ProcessWorkingSetSize: " & objItem.WorkingSetSize
      if SysTotalVisibleMemorySize > 0 Then
        WScript.Echo "ProcessProMemPercent: " & rount(objItem.WorkingSetSize / SysTotalVisibleMemorySize)
      else
        WScript.Echo "ProcessProMemPercent: "
      end if
        
	Next
end function


GetProcessInfo()