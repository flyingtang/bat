@echo off

set /a sum=0

for %%x in (*.txt)  do (

echo %%x �ļ����������£�
type %%x
echo .
set /a sum=sum+1
)

echo һ����ʾ��%sum%���ı��ļ�
