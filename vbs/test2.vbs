function GetRuntimeStr(second)
	const intDaySecond = 86400 'day
	const intHourSecond = 3600 'day
	const intMinuteSecond = 60 'day

	intDay = second \ intDaySecond
	second = second mod intDaySecond    

	intHour = second \ intHourSecond
	second = second mod intHourSecond    

	intMinute = second \ intMinuteSecond
	second = second mod intMinuteSecond    

	WScript.Echo "abc" & intDay , intHour, intMinute, second

	if intDay > 0 Then
		WScript.Echo strRuntime & intDay & "��"
		strRuntime = strRuntime & intDay & "��"
	end if

	if intHour > 0 Then
		strRuntime = strRuntime  & intHour &"Сʱ"
	end if

	if intMinute > 0 Then
		strRuntime = strRuntime  & intMinute & "����"
	end if
	strRuntime = strRuntime  & second & "��"
	GetRuntimeStr=strRuntime
end function


dim ss
ss=520000
r = GetRuntimeStr(ss)
WScript.Echo "���: "& r