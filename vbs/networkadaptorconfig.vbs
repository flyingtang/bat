On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration",,48)

Dim objItem 'as Win32_NetworkAdapterConfiguration
For Each objItem in colItems
	WScript.Echo "ArpAlwaysSourceRoute: " & objItem.ArpAlwaysSourceRoute
	WScript.Echo "ArpUseEtherSNAP: " & objItem.ArpUseEtherSNAP
	WScript.Echo "Caption: " & objItem.Caption
	WScript.Echo "DatabasePath: " & objItem.DatabasePath
	WScript.Echo "DeadGWDetectEnabled: " & objItem.DeadGWDetectEnabled
	WScript.Echo "DefaultIPGateway: " & objItem.DefaultIPGateway
	WScript.Echo "DefaultTOS: " & objItem.DefaultTOS
	WScript.Echo "DefaultTTL: " & objItem.DefaultTTL
	WScript.Echo "Description: " & objItem.Description
	WScript.Echo "DHCPEnabled: " & objItem.DHCPEnabled
	WScript.Echo "DHCPLeaseExpires: " & objItem.DHCPLeaseExpires
	WScript.Echo "DHCPLeaseObtained: " & objItem.DHCPLeaseObtained
	WScript.Echo "DHCPServer: " & objItem.DHCPServer
	WScript.Echo "DNSDomain: " & objItem.DNSDomain
	WScript.Echo "DNSDomainSuffixSearchOrder: " & objItem.DNSDomainSuffixSearchOrder
	WScript.Echo "DNSEnabledForWINSResolution: " & objItem.DNSEnabledForWINSResolution
	WScript.Echo "DNSHostName: " & objItem.DNSHostName
	WScript.Echo "DNSServerSearchOrder: " & objItem.DNSServerSearchOrder
	WScript.Echo "DomainDNSRegistrationEnabled: " & objItem.DomainDNSRegistrationEnabled
	WScript.Echo "ForwardBufferMemory: " & objItem.ForwardBufferMemory
	WScript.Echo "FullDNSRegistrationEnabled: " & objItem.FullDNSRegistrationEnabled
	WScript.Echo "GatewayCostMetric: " & objItem.GatewayCostMetric
	WScript.Echo "IGMPLevel: " & objItem.IGMPLevel
	WScript.Echo "Index: " & objItem.Index
	WScript.Echo "InterfaceIndex: " & objItem.InterfaceIndex
	WScript.Echo "IPAddress: " & objItem.IPAddress
	WScript.Echo "IPConnectionMetric: " & objItem.IPConnectionMetric
	WScript.Echo "IPEnabled: " & objItem.IPEnabled
	WScript.Echo "IPFilterSecurityEnabled: " & objItem.IPFilterSecurityEnabled
	WScript.Echo "IPPortSecurityEnabled: " & objItem.IPPortSecurityEnabled
	WScript.Echo "IPSecPermitIPProtocols: " & objItem.IPSecPermitIPProtocols
	WScript.Echo "IPSecPermitTCPPorts: " & objItem.IPSecPermitTCPPorts
	WScript.Echo "IPSecPermitUDPPorts: " & objItem.IPSecPermitUDPPorts
	WScript.Echo "IPSubnet: " & objItem.IPSubnet
	WScript.Echo "IPUseZeroBroadcast: " & objItem.IPUseZeroBroadcast
	WScript.Echo "IPXAddress: " & objItem.IPXAddress
	WScript.Echo "IPXEnabled: " & objItem.IPXEnabled
	WScript.Echo "IPXFrameType: " & objItem.IPXFrameType
	WScript.Echo "IPXMediaType: " & objItem.IPXMediaType
	WScript.Echo "IPXNetworkNumber: " & objItem.IPXNetworkNumber
	WScript.Echo "IPXVirtualNetNumber: " & objItem.IPXVirtualNetNumber
	WScript.Echo "KeepAliveInterval: " & objItem.KeepAliveInterval
	WScript.Echo "KeepAliveTime: " & objItem.KeepAliveTime
	WScript.Echo "MACAddress: " & objItem.MACAddress
	WScript.Echo "MTU: " & objItem.MTU
	WScript.Echo "NumForwardPackets: " & objItem.NumForwardPackets
	WScript.Echo "PMTUBHDetectEnabled: " & objItem.PMTUBHDetectEnabled
	WScript.Echo "PMTUDiscoveryEnabled: " & objItem.PMTUDiscoveryEnabled
	WScript.Echo "ServiceName: " & objItem.ServiceName
	WScript.Echo "SettingID: " & objItem.SettingID
	WScript.Echo "TcpipNetbiosOptions: " & objItem.TcpipNetbiosOptions
	WScript.Echo "TcpMaxConnectRetransmissions: " & objItem.TcpMaxConnectRetransmissions
	WScript.Echo "TcpMaxDataRetransmissions: " & objItem.TcpMaxDataRetransmissions
	WScript.Echo "TcpNumConnections: " & objItem.TcpNumConnections
	WScript.Echo "TcpUseRFC1122UrgentPointer: " & objItem.TcpUseRFC1122UrgentPointer
	WScript.Echo "TcpWindowSize: " & objItem.TcpWindowSize
	WScript.Echo "WINSEnableLMHostsLookup: " & objItem.WINSEnableLMHostsLookup
	WScript.Echo "WINSHostLookupFile: " & objItem.WINSHostLookupFile
	WScript.Echo "WINSPrimaryServer: " & objItem.WINSPrimaryServer
	WScript.Echo "WINSScopeID: " & objItem.WINSScopeID
	WScript.Echo "WINSSecondaryServer: " & objItem.WINSSecondaryServer
	WScript.Echo ""
Next
