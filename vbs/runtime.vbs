

function formatWindowsDate (dt)
    tmp=dt
    WScript.Echo dt
    count = 4
    y=Left(tmp, count)
    tmp=Mid(tmp, count+1)
    WScript.Echo y & " " & tmp

    count = 2
    m=Left(tmp, count)
    tmp=Mid(tmp, count+1)
    WScript.Echo m & " " & tmp

    d=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    h=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    mi=Left(tmp, 2)
    tmp=Mid(tmp, count+1)

    s=Left(tmp, 2)
    tmp=Mid(tmp, count+1)
    
	formatWindowsDate= m&"/"&d&"/"&y&" "&h&":"&mi&":"&s
	' formatWindowsDate=#m/d/y h:mi:s#
end function

dt=formatWindowsDate("20201221122955")

datOld = CDate(dt)
datNow = Date()

intDiffSecond=DateDiff("s", datNow,datOld)

WScript.Echo intSecond
