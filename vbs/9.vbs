
Dim  Info() '定义动态数组
ReDim Info( 8, 9 ) '初始化

InfoSize = UBound( Info, 2 ) '得到第二维最大下标
WScript.Echo InfoSize
ReDim Preserve Info( 8, InfoSize+1 ) '重新定义第二维