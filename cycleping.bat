@echo off

:again
ping baidu.com > null

if not %errorlevel% EQU 0 goto again

start "��������������ͨ��" echo ���ڿ�������pingͨ�ٶ�����