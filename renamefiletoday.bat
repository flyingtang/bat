@echo off


if not exist %1% ( echo �����÷����£�
    echo %0 filename
    echo filename ������Ҫ�޸ĵ��ļ���
    echo.
    echo.
    goto end
)

@REM ~x ȡ��չ��
set extension=%~x1

@REM  /F ʹ���ļ�����
for /F "tokens=1-3 delims=/-" %%A in  ('date/T') do set date=%%A%%B%%C
ren "%1" "%date%%extension%"
set exension=
set date=
:end