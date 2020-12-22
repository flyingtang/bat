set oExec = createobject("wscript.shell").exec("TASKLIST /V /FO LIST")
sResults = oExec.Stdout.ReadAll
WScript.Echo sResults
count = 0
do while oExec.Stdout.AtEndOfLine <> True
 line = oExec.ReadLine
 count = count + 1
 WScript.Echo "count: " & count
loop

