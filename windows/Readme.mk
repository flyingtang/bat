### 1、监控中心-监控管理-脚本设置

* 上传文件，后缀为.vbs
* 脚本解析为bat  例如  cmd /c cscript /Nologo windows_default.vbs
* 参数出参名称如下

### 2、参数出参列表

系统信息 

```ini
SysCaption=Microsoft Windows 10 企业版
SysCSName=DESKTOP-JK64T46
SysDescription=
# 系统物理内存剩余（G）
SysFreePhysicalMemory=
SysFreeSpaceInPagingFiles=9753800
# 系统虚拟内存剩余 （G）
SysFreeVirtualMemory=
SysLargeSystemCache=
SysManufacturer=Microsoft Corporation
# 系统进程上限
SysMaxNumberOfProcesses=-1
SysMaxProcessMemorySize=137438953344
SysName=Microsoft Windows 10 企业版
SysNumberOfLicensedUsers=
# 进程总数
SysNumberOfProcesses=175
# 系统用户数
SysNumberOfUsers=2
# 序列号
SysSerialNumber=00329-00000-00003-AA066
# 交换总量
SysTotalSwapSpaceSize=
# 虚拟内存总量
SysTotalVirtualMemorySize=26640156
# 内存总量 （G）
SysTotalVisibleMemorySize=
# 系统版本号
SysVersion= 10.0.10240
# 运行时间
Sysruntime=1天19小时11分钟36秒
# 内存使用率
SysPercentUsedMemory=74.2842540814371
```

CPU 信息

```ini
CpuAddressWidth=64
CpuCaption=Intel64 Family 6 Model 158 Stepping 10
CpuCpuStatus=1
CpuCreationClassName=Win32_Processor
CpuCurrentClockSpeed=3000
CpuCurrentVoltage=10
CpuDataWidth=64
CpuDescription=Intel64 Family 6 Model 158 Stepping 10
CpuLoadPercentage=43
CpuManufacturer=GenuineIntel
CpuMaxClockSpeed=3000
CpuName=Intel(R) Core(TM) i5-8500 CPU @ 3.00GHz
CpuProcessorId=BFEBFBFF000906EA
CpuProcessorType=3
CpuRevision=
CpuRole=CPU
CpuSocketDesignation=U3E1
CpuSystemName=DESKTOP-JK64T46
CpuCount=1
```



磁盘信息 列表 （备注： 目前 Disk 和 Volunm 可以当成一样东西）

```
​```ini
DiskName=G:
DiskAvgDiskBytesPerRead=0
DiskAvgDiskBytesPerTransfer=0
DiskAvgDiskBytesPerWrite=0
DiskAvgDisksecPerRead=0
DiskAvgDisksecPerTransfer=0
DiskAvgDisksecPerWrite=0
DiskBytesPersec=0
DiskReadBytesPersec=0
DiskReadsPersec=0
DiskTransfersPersec=0
DiskWriteBytesPersec=0
DiskWritesPersec=0
DiskFreeMegabytes=96553
DiskPercentDiskReadTime=0
DiskPercentDiskTime=0
DiskPercentDiskWriteTime=0
DiskPercentFreeSpace=29
DiskPercentIdleTime=100
VolunmLabel=数据
VolunmName=G:
VolunmSerialNumber=705221
VolunmFreeSpace=101243686912
# 券容量 单位字节
VolunmCapacity=338776756224
VolunmUsedSpace=237533069312
VolunmPercentUsed=70.1149252267303
VolunmDriveType=Local Disk
VolunmFileSystem=NTFS

# 总计信息
VolunmTotalCapacity=1120114757632
VolunmTotalFreeSpace=507359977472
VolunmTotalUsedSpace=612754780160
```

网卡信息 列表
```ini
# 上一版本
# NetworkAdaptorBytesReceivedPersec=0
# NetworkAdaptorBytesSentPersec=0
# NetworkAdaptorBytesTotalPersec=0
# NetworkAdaptorCaption=
# # 指以位/每秒估计的网络接口的当前带宽
# NetworkAdaptorCurrentBandwidth=3000000
# NetworkAdaptorDescription=
# NetworkAdaptorFrequency_Object=
# NetworkAdaptorFrequency_PerfTime=
# NetworkAdaptorFrequency_Sys100NS=
# NetworkAdaptorName=Bluetooth Device [Personal Area Network]
# NetworkAdaptorOffloadedConnections=0
# NetworkAdaptorOutputQueueLength=0
# NetworkAdaptorPacketsOutboundDiscarded=0
# NetworkAdaptorPacketsOutboundErrors=0
# NetworkAdaptorPacketsPersec=0
# NetworkAdaptorPacketsReceivedDiscarded=0
# NetworkAdaptorPacketsReceivedErrors=0
# NetworkAdaptorPacketsReceivedNonUnicastPersec=0
# NetworkAdaptorPacketsReceivedPersec=0
# NetworkAdaptorPacketsReceivedUnicastPersec=0
# NetworkAdaptorPacketsReceivedUnknown=0
# NetworkAdaptorPacketsSentNonUnicastPersec=0
# NetworkAdaptorPacketsSentPersec=0
# NetworkAdaptorPacketsSentUnicastPersec=0
# NetworkAdaptorTCPActiveRSCConnections=0
# NetworkAdaptorTCPRSCAveragePacketSize=0
# NetworkAdaptorTCPRSCCoalescedPacketsPersec=0
# NetworkAdaptorTCPRSCExceptionsPersec=0
# NetworkAdaptorTimestamp_Object=
# NetworkAdaptorTimestamp_PerfTime=
# NetworkAdaptorTimestamp_Sys100NS=
NetworkAdaptorAdapterType=
NetworkAdaptorAdapterTypeId=
NetworkAdaptorCaption=[00000000] Microsoft Kernel Debug Network Adapter
NetworkAdaptorDescription=Microsoft Kernel Debug Network Adapter
NetworkAdaptorMACAddress=
NetworkAdaptorManufacturer=Microsoft
NetworkAdaptorMaxSpeed=
NetworkAdaptorName=Microsoft Kernel Debug Network Adapter
NetworkAdaptorNetConnectionID=
NetworkAdaptorNetConnectionStatus=
NetworkAdaptorNetEnabled=
NetworkAdaptorNetworkAddresses=
NetworkAdaptorPermanentAddress=
NetworkAdaptorPhysicalAdapter=False
NetworkAdaptorPNPDeviceID=ROOT\KDNIC\0000
NetworkAdaptorProductName=Microsoft Kernel Debug Network Adapter
NetworkAdaptorServiceName=kdnic
```

网络io相关
```ini
NetworkAdaptorSpeed=
NetworkAdaptorStatus=
NetworkAdaptorStatusInfo=
NetworkAdaptorBytesReceivedPersec=
NetworkAdaptorBytesSentPersec=
NetworkAdaptorBytesTotalPersec=
NetworkAdaptorCurrentBandwidth=
NetworkAdaptorFrequency_Object=
NetworkAdaptorFrequency_PerfTime=
NetworkAdaptorFrequency_Sys100NS=
NetworkAdaptorPacketsPersec=
NetworkAdaptorPacketsReceivedDiscarded=
NetworkAdaptorPacketsReceivedErrors=
NetworkAdaptorPacketsReceivedNonUnicastPersec=
NetworkAdaptorPacketsReceivedPersec=
NetworkAdaptorPacketsReceivedUnicastPersec=
NetworkAdaptorPacketsReceivedUnknown=
NetworkAdaptorPacketsSentNonUnicastPersec=
NetworkAdaptorPacketsSentPersec=
NetworkAdaptorPacketsSentUnicastPersec=
```

服务信息 列表
```ini
ServiceAcceptPause=False
ServiceAcceptStop=True
ServiceCaption=Adobe Genuine Monitor Service
ServiceCheckPoint=0
ServiceCreationClassName=Win32_Service
ServiceDelayedAutoStart=False
ServiceDescription=Adobe Genuine Monitor Service
ServiceDesktopInteract=False
ServiceDisplayName=Adobe Genuine Monitor Service
ServiceErrorControl=Normal
ServiceExitCode=0
ServiceInstallDate=
ServiceName=AGMService
ServicePathName="C:\Program Files (x86)\Common Files\Adobe\AdobeGCClient\AGMService.exe"
ServiceProcessId=2096
ServiceServiceSpecificExitCode=0
ServiceServiceType=Own Process
ServiceStarted=True
ServiceStartMode=Auto
ServiceStartName=LocalSystem
ServiceState=Running
ServiceStatus=OK
ServiceSystemCreationClassName=Win32_ComputerSystem
ServiceSystemName=DESKTOP-JK64T46
ServiceTagId=0
ServiceWaitHint=0
```

计划任务
```ini
SchTaskName=                             zxAgentCheckDown
SchTaskNextRuntim=                       9:27:00, 2020-12-25
SchTaskLastRuntime=                       9:26:00, 2020-12-25
SchTaskLastResult=                           0
SchTaskName=                             zxAgentCheckDown
SchTaskMode=
SchTaskStatus=                       已启用
SchTask=                       C:\Program Files\zxops\check_down.bat
SchTaskType=                         每天
```


进程相关 列表
```ini
ProcessProcessId=500
ProcessName=csrss.exe
ProcessParentProcessId=444
ProcessThreadCount=17
ProcessHandleCount=538
ProcessCpuTime=37656250.546875
ProcessKernelModeTime=37656250
ProcessUserModeTime=5468750
ProcessWorkingSetSize=5115904
ProcessPercentProcessorTime=0
ProcessWorkingSet=5115904
ProcessMemused= 4,996 K
ProcessCpuused= 0:00:04
```


tcp 相关
```ini
TcpConnectionFailures=7324
TcpConnectionsActive=42227
TcpConnectionsEstablished=30
TcpConnectionsPassive=10497
TcpConnectionsReset=2948
TcpSegmentsPersec=4727466
TcpSegmentsReceivedPersec=2612280
TcpSegmentsRetransmittedPersec=244641
TcpSegmentsSentPersec=2115186
```

账户列表
```ini
AccountCaption=DESKTOP-JK64T46\__vmware__
AccountDescription=VMware User Group
AccountDomain=DESKTOP-JK64T46
AccountInstallDate=
AccountLocalAccount=True
AccountName=__vmware__
AccountSID=S-1-5-21-1378040621-233758120-3585781211-1002
AccountSIDType=4
AccountStatus=OK
```


pagefile
```ini

PageFileAllocatedBaseSize: 9728
PageFileCurrentUsage: 353
PageFilePeakUsage: 395
#页使用率
PageFilePercentUsed: 3.62870065789474

```