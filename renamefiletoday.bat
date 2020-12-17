@echo off


if not exist %1% ( echo 命令用法如下：
    echo %0 filename
    echo filename 代表需要修改的文件名
    echo.
    echo.
    goto end
)

@REM ~x 取扩展名
set extension=%~x1

@REM  /F 使用文件解析
for /F "tokens=1-3 delims=/-" %%A in  ('date/T') do set date=%%A%%B%%C
ren "%1" "%date%%extension%"
set exension=
set date=
:end