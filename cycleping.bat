@echo off

:again
ping baidu.com > null

if not %errorlevel% EQU 0 goto again

start "可以正常与主机通信" echo 现在可以正常ping通百度主机