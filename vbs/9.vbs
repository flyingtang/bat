
Dim  Info() '���嶯̬����
ReDim Info( 8, 9 ) '��ʼ��

InfoSize = UBound( Info, 2 ) '�õ��ڶ�ά����±�
WScript.Echo InfoSize
ReDim Preserve Info( 8, InfoSize+1 ) '���¶���ڶ�ά