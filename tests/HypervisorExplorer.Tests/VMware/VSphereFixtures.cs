namespace HypervisorExplorer.Tests.VMware;

/// <summary>
/// Realistic vim25 SOAP responses for a small vCenter: DC1 / Cluster01 with two ESXi hosts, a distributed switch
/// (DSwitch01 + PG-App-100), a standard vSwitch0, two datastores (VMFS + NFS) and three VMs (web01 with a
/// two-level snapshot tree, db01 powered off, tmpl-ubuntu template).
/// </summary>
internal static class VSphereFixtures
{
    public const string InstanceUuid = "7d1e2a3b-4c5d-4e6f-8a9b-0c1d2e3f4a5b";

    public static string Envelope(string body) =>
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        + "<soapenv:Envelope xmlns:soapenc=\"http://schemas.xmlsoap.org/soap/encoding/\" "
        + "xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" "
        + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n<soapenv:Body>\n" + body + "\n</soapenv:Body>\n</soapenv:Envelope>";

    public static string Fault(string faultType, string message, string detailInner = "") =>
        Envelope($"""
            <soapenv:Fault><faultcode>ServerFaultCode</faultcode><faultstring>{message}</faultstring><detail><{faultType}Fault xmlns="urn:vim25" xsi:type="{faultType}">{detailInner}</{faultType}Fault></detail></soapenv:Fault>
            """);

    public static string Empty(string op) => Envelope($"<{op}Response xmlns=\"urn:vim25\"></{op}Response>");

    public static string RetrieveResult(string objects, string? token = null, bool continuation = false)
    {
        var op = continuation ? "ContinueRetrievePropertiesEx" : "RetrievePropertiesEx";
        var tok = token is null ? "" : $"<token>{token}</token>";
        return Envelope($"<{op}Response xmlns=\"urn:vim25\"><returnval>{tok}{objects}</returnval></{op}Response>");
    }

    /// <summary>An ObjectContent.</summary>
    public static string O(string type, string id, params string[] propSets) =>
        $"<objects><obj type=\"{type}\">{id}</obj>{string.Concat(propSets)}</objects>";

    /// <summary>A DynamicProperty (propSet) with a typed value.</summary>
    public static string P(string name, string xsiType, string inner) =>
        $"<propSet><name>{name}</name><val xsi:type=\"{xsiType}\">{inner}</val></propSet>";

    public static string S(string name, string value) => P(name, "xsd:string", value);
    public static string B(string name, bool value) => P(name, "xsd:boolean", value ? "true" : "false");
    public static string Ref(string name, string type, string value) =>
        $"<propSet><name>{name}</name><val type=\"{type}\" xsi:type=\"ManagedObjectReference\">{value}</val></propSet>";
    public static string Refs(string name, params (string Type, string Value)[] refs) =>
        P(name, "ArrayOfManagedObjectReference",
            string.Concat(refs.Select(r => $"<ManagedObjectReference type=\"{r.Type}\" xsi:type=\"ManagedObjectReference\">{r.Value}</ManagedObjectReference>")));

    public const string ServiceContent = """
        <RetrieveServiceContentResponse xmlns="urn:vim25"><returnval>
          <rootFolder type="Folder">group-d1</rootFolder>
          <propertyCollector type="PropertyCollector">propertyCollector</propertyCollector>
          <viewManager type="ViewManager">ViewManager</viewManager>
          <about>
            <name>VMware vCenter Server</name>
            <fullName>VMware vCenter Server 8.0.2 build-22617221</fullName>
            <vendor>VMware, Inc.</vendor>
            <version>8.0.2</version>
            <patchLevel></patchLevel>
            <build>22617221</build>
            <localeVersion>INTL</localeVersion>
            <localeBuild>000</localeBuild>
            <osType>linux-x64</osType>
            <productLineId>vpx</productLineId>
            <apiType>VirtualCenter</apiType>
            <apiVersion>8.0.2.0</apiVersion>
            <instanceUuid>7d1e2a3b-4c5d-4e6f-8a9b-0c1d2e3f4a5b</instanceUuid>
            <licenseProductName>VMware VirtualCenter Server</licenseProductName>
            <licenseProductVersion>8.0</licenseProductVersion>
          </about>
          <setting type="OptionManager">VpxSettings</setting>
          <userDirectory type="UserDirectory">UserDirectory</userDirectory>
          <sessionManager type="SessionManager">SessionManager</sessionManager>
          <authorizationManager type="AuthorizationManager">AuthorizationManager</authorizationManager>
          <perfManager type="PerformanceManager">PerfMgr</perfManager>
          <scheduledTaskManager type="ScheduledTaskManager">ScheduledTaskManager</scheduledTaskManager>
          <eventManager type="EventManager">EventManager</eventManager>
          <taskManager type="TaskManager">TaskManager</taskManager>
          <licenseManager type="LicenseManager">LicenseManager</licenseManager>
          <searchIndex type="SearchIndex">SearchIndex</searchIndex>
          <fileManager type="FileManager">FileManager</fileManager>
          <virtualDiskManager type="VirtualDiskManager">virtualDiskManager</virtualDiskManager>
        </returnval></RetrieveServiceContentResponse>
        """;

    public const string LoginResponse = """
        <LoginResponse xmlns="urn:vim25"><returnval><key>52b5f1d2-0d36-6a40-5d84-8b4f1f0c9e1a</key><userName>VSPHERE.LOCAL\Administrator</userName><fullName>Administrator vsphere.local</fullName><loginTime>2026-10-06T09:00:00.123Z</loginTime><lastActiveTime>2026-10-06T09:00:00.123Z</lastActiveTime><locale>en</locale><messageLocale>en</messageLocale><extensionSession>false</extensionSession><ipAddress>10.0.0.5</ipAddress><userAgent>HypervisorExplorer</userAgent><callCount>0</callCount></returnval></LoginResponse>
        """;

    public static string ContainerView(string type) =>
        Envelope($"<CreateContainerViewResponse xmlns=\"urn:vim25\"><returnval type=\"ContainerView\">session[52b5f1d2-0d36-6a40-5d84-8b4f1f0c9e1a]52{type}</returnval></CreateContainerViewResponse>");

    // ------------------------------------------------------------------ inventory tree (ManagedEntity: name, parent)

    private static string E(string type, string id, string name, string? parentType, string? parent) =>
        O(type, id, S("name", name), parentType is null ? "" : Ref("parent", parentType, parent!));

    public static string Entities() => RetrieveResult(string.Concat(
        E("Datacenter", "datacenter-1", "DC1", "Folder", "group-d1"),
        E("Folder", "group-v3", "vm", "Datacenter", "datacenter-1"),
        E("Folder", "group-h4", "host", "Datacenter", "datacenter-1"),
        E("Folder", "group-s5", "datastore", "Datacenter", "datacenter-1"),
        E("Folder", "group-n6", "network", "Datacenter", "datacenter-1"),
        E("Folder", "group-v100", "Linux", "Folder", "group-v3"),
        E("ClusterComputeResource", "domain-c10", "Cluster01", "Folder", "group-h4"),
        E("ResourcePool", "resgroup-11", "Resources", "ClusterComputeResource", "domain-c10"),
        E("ResourcePool", "resgroup-50", "Prod", "ResourcePool", "resgroup-11"),
        E("HostSystem", "host-20", "esx01.lab.local", "ClusterComputeResource", "domain-c10"),
        E("HostSystem", "host-21", "esx02.lab.local", "ClusterComputeResource", "domain-c10"),
        E("Datastore", "datastore-30", "ds-vmfs01", "Folder", "group-s5"),
        E("Datastore", "datastore-31", "nfs01", "Folder", "group-s5"),
        E("Network", "network-40", "VM Network", "Folder", "group-n6"),
        E("VmwareDistributedVirtualSwitch", "dvs-60", "DSwitch01", "Folder", "group-n6"),
        E("DistributedVirtualPortgroup", "dvportgroup-61", "PG-App-100", "Folder", "group-n6"),
        E("DistributedVirtualPortgroup", "dvportgroup-62", "DSwitch01-DVUplinks-60", "Folder", "group-n6"),
        E("VirtualMachine", "vm-100", "web01", "Folder", "group-v100"),
        E("VirtualMachine", "vm-101", "db01", "Folder", "group-v3"),
        E("VirtualMachine", "vm-102", "tmpl-ubuntu", "Folder", "group-v3")));

    // ------------------------------------------------------------------ hosts

    private const string Esx01Network = """
        <vswitch><name>vSwitch0</name><key>key-vim.host.VirtualSwitch-vSwitch0</key><numPorts>2560</numPorts><numPortsAvailable>2546</numPortsAvailable><mtu>1500</mtu>
          <portgroup>key-vim.host.PortGroup-VM Network</portgroup><portgroup>key-vim.host.PortGroup-Management Network</portgroup>
          <pnic>key-vim.host.PhysicalNic-vmnic0</pnic>
          <spec><numPorts>2560</numPorts><bridge xsi:type="HostVirtualSwitchBondBridge"><nicDevice>vmnic0</nicDevice></bridge>
            <policy><security><allowPromiscuous>false</allowPromiscuous><macChanges>false</macChanges><forgedTransmits>false</forgedTransmits></security>
              <nicTeaming><policy>loadbalance_srcid</policy><reversePolicy>true</reversePolicy><notifySwitches>true</notifySwitches><rollingOrder>false</rollingOrder>
                <failureCriteria><checkSpeed>minimum</checkSpeed><speed>10</speed><checkDuplex>false</checkDuplex><fullDuplex>false</fullDuplex><checkErrorPercent>false</checkErrorPercent><percentage>0</percentage><checkBeacon>false</checkBeacon></failureCriteria>
                <nicOrder><activeNic>vmnic0</activeNic></nicOrder></nicTeaming>
              <offloadPolicy><csumOffload>true</csumOffload><tcpSegmentation>true</tcpSegmentation><zeroCopyXmit>true</zeroCopyXmit></offloadPolicy>
              <shapingPolicy><enabled>false</enabled></shapingPolicy></policy><mtu>1500</mtu></spec>
        </vswitch>
        <proxySwitch><dvsUuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</dvsUuid><dvsName>DSwitch01</dvsName><key>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</key><numPorts>512</numPorts><numPortsAvailable>500</numPortsAvailable>
          <uplinkPort><key>16</key><value>Uplink 1</value></uplinkPort>
          <pnic>key-vim.host.PhysicalNic-vmnic1</pnic><mtu>9000</mtu>
          <spec><backing xsi:type="DistributedVirtualSwitchHostMemberPnicBacking"><pnicSpec><pnicDevice>vmnic1</pnicDevice><uplinkPortKey>16</uplinkPortKey><uplinkPortgroupKey>dvportgroup-62</uplinkPortgroupKey></pnicSpec></backing></spec>
        </proxySwitch>
        <portgroup><key>key-vim.host.PortGroup-VM Network</key><vswitch>key-vim.host.VirtualSwitch-vSwitch0</vswitch>
          <computedPolicy><security><allowPromiscuous>false</allowPromiscuous><macChanges>false</macChanges><forgedTransmits>false</forgedTransmits></security><nicTeaming><policy>loadbalance_srcid</policy><reversePolicy>true</reversePolicy><notifySwitches>true</notifySwitches><rollingOrder>false</rollingOrder></nicTeaming><offloadPolicy><csumOffload>true</csumOffload><tcpSegmentation>true</tcpSegmentation><zeroCopyXmit>true</zeroCopyXmit></offloadPolicy><shapingPolicy><enabled>false</enabled></shapingPolicy></computedPolicy>
          <spec><name>VM Network</name><vlanId>0</vlanId><vswitchName>vSwitch0</vswitchName><policy><security></security><nicTeaming></nicTeaming><offloadPolicy></offloadPolicy><shapingPolicy></shapingPolicy></policy></spec></portgroup>
        <portgroup><key>key-vim.host.PortGroup-Management Network</key><port><key>key-vim.host.PortGroup.Port-33554436</key><mac>00:50:56:6a:11:01</mac><type>host</type></port><vswitch>key-vim.host.VirtualSwitch-vSwitch0</vswitch>
          <computedPolicy><security><allowPromiscuous>false</allowPromiscuous><macChanges>true</macChanges><forgedTransmits>true</forgedTransmits></security><nicTeaming><policy>failover_explicit</policy><reversePolicy>true</reversePolicy><notifySwitches>true</notifySwitches><rollingOrder>false</rollingOrder></nicTeaming><offloadPolicy><csumOffload>true</csumOffload><tcpSegmentation>true</tcpSegmentation><zeroCopyXmit>true</zeroCopyXmit></offloadPolicy><shapingPolicy><enabled>false</enabled></shapingPolicy></computedPolicy>
          <spec><name>Management Network</name><vlanId>10</vlanId><vswitchName>vSwitch0</vswitchName><policy></policy></spec></portgroup>
        <pnic><key>key-vim.host.PhysicalNic-vmnic0</key><device>vmnic0</device><pci>0000:18:00.0</pci><driver>ixgben</driver><linkSpeed><speedMb>10000</speedMb><duplex>true</duplex></linkSpeed><wakeOnLanSupported>false</wakeOnLanSupported><mac>b4:96:91:aa:00:01</mac></pnic>
        <pnic><key>key-vim.host.PhysicalNic-vmnic1</key><device>vmnic1</device><pci>0000:18:00.1</pci><driver>ixgben</driver><linkSpeed><speedMb>25000</speedMb><duplex>true</duplex></linkSpeed><wakeOnLanSupported>false</wakeOnLanSupported><mac>b4:96:91:aa:00:02</mac></pnic>
        <pnic><key>key-vim.host.PhysicalNic-vmnic2</key><device>vmnic2</device><pci>0000:19:00.0</pci><driver>ntg3</driver><wakeOnLanSupported>true</wakeOnLanSupported><mac>b4:96:91:aa:00:03</mac></pnic>
        <vnic><device>vmk0</device><key>key-vim.host.VirtualNic-vmk0</key><portgroup>Management Network</portgroup>
          <spec><ip><dhcp>false</dhcp><ipAddress>10.0.10.21</ipAddress><subnetMask>255.255.255.0</subnetMask><ipV6Config><ipV6Address><ipAddress>fe80::250:56ff:fe6a:1101</ipAddress><prefixLength>64</prefixLength><origin>other</origin><dadState>preferred</dadState></ipV6Address><autoConfigurationEnabled>false</autoConfigurationEnabled><dhcpV6Enabled>false</dhcpV6Enabled></ipV6Config></ip><mac>00:50:56:6a:11:01</mac><mtu>1500</mtu><tsoEnabled>true</tsoEnabled><netStackInstanceKey>defaultTcpipStack</netStackInstanceKey></spec><port>key-vim.host.PortGroup.Port-33554436</port></vnic>
        <vnic><device>vmk1</device><key>key-vim.host.VirtualNic-vmk1</key><portgroup></portgroup>
          <spec><ip><dhcp>false</dhcp><ipAddress>10.0.20.21</ipAddress><subnetMask>255.255.255.0</subnetMask></ip><mac>00:50:56:6b:22:01</mac><distributedVirtualPort><switchUuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</switchUuid><portgroupKey>dvportgroup-61</portgroupKey><portKey>12</portKey></distributedVirtualPort><mtu>9000</mtu></spec></vnic>
        <dnsConfig xsi:type="HostDnsConfig"><dhcp>false</dhcp><hostName>esx01</hostName><domainName>lab.local</domainName><address>10.0.0.10</address><address>10.0.0.11</address><searchDomain>lab.local</searchDomain></dnsConfig>
        <ipRouteConfig xsi:type="HostIpRouteConfig"><defaultGateway>10.0.10.1</defaultGateway><gatewayDevice>vmk0</gatewayDevice></ipRouteConfig>
        <atBootIpV6Enabled>true</atBootIpV6Enabled><ipV6Enabled>true</ipV6Enabled>
        """;

    private static string HostObject(string id, string name, string serial, bool maintenance, string networkInner, bool rich) => O("HostSystem", id,
        S("name", name),
        Ref("parent", "ClusterComputeResource", "domain-c10"),
        P("overallStatus", "ManagedEntityStatus", maintenance ? "yellow" : "green"),
        P("configStatus", "ManagedEntityStatus", "green"),
        P("runtime.connectionState", "HostSystemConnectionState", "connected"),
        B("runtime.inMaintenanceMode", maintenance),
        B("runtime.inQuarantineMode", false),
        P("runtime.bootTime", "xsd:dateTime", "2026-09-01T06:12:44.511Z"),
        P("runtime.powerState", "HostSystemPowerState", "poweredOn"),
        P("summary.hardware", "HostHardwareSummary", $"""
            <vendor>Dell Inc.</vendor><model>PowerEdge R650</model><uuid>4c4c4544-0042-3510-8051-{serial}</uuid>
            <otherIdentifyingInfo><identifierValue>{serial}</identifierValue><identifierType><label>Service tag</label><summary>Service tag of the system</summary><key>ServiceTag</key></identifierType></otherIdentifyingInfo>
            <otherIdentifyingInfo><identifierValue>Dell System</identifierValue><identifierType><label>OEM specific string</label><summary>OEM specific string</summary><key>OemSpecificString</key></identifierType></otherIdentifyingInfo>
            <memorySize>274877906944</memorySize><cpuModel>Intel(R) Xeon(R) Gold 6342 CPU @ 2.80GHz</cpuModel><cpuMhz>2800</cpuMhz>
            <numCpuPkgs>2</numCpuPkgs><numCpuCores>48</numCpuCores><numCpuThreads>96</numCpuThreads><numNics>3</numNics><numHBAs>2</numHBAs>
            """),
        P("summary.quickStats", "HostListSummaryQuickStats", "<overallCpuUsage>13440</overallCpuUsage><overallMemoryUsage>131072</overallMemoryUsage><distributedCpuFairness>0</distributedCpuFairness><distributedMemoryFairness>0</distributedMemoryFairness><uptime>3036000</uptime>"),
        P("summary.config.product", "AboutInfo", "<name>VMware ESXi</name><fullName>VMware ESXi 8.0.2 build-22380479</fullName><vendor>VMware, Inc.</vendor><version>8.0.2</version><build>22380479</build><localeVersion>INTL</localeVersion><localeBuild>000</localeBuild><osType>vmnix-x86</osType><productLineId>embeddedEsx</productLineId><apiType>HostAgent</apiType><apiVersion>8.0.2.0</apiVersion><licenseProductName>VMware ESX Server</licenseProductName><licenseProductVersion>8.0</licenseProductVersion>"),
        S("summary.currentEVCModeKey", "intel-icelake"),
        S("summary.maxEVCModeKey", "intel-icelake"),
        P("hardware.biosInfo", "HostBIOSInfo", "<biosVersion>1.10.2</biosVersion><releaseDate>2023-05-17T00:00:00Z</releaseDate><vendor>Dell Inc.</vendor><majorRelease>1</majorRelease><minorRelease>10</minorRelease>"),
        P("hardware.systemInfo", "HostSystemInfo", $"<vendor>Dell Inc.</vendor><model>PowerEdge R650</model><uuid>4c4c4544-0042-3510-8051-{serial}</uuid><serialNumber>{serial}</serialNumber>"),
        P("hardware.cpuPowerManagementInfo", "HostCpuPowerManagementInfo", "<currentPolicy>Balanced</currentPolicy><hardwareSupport>ACPI P-states, ACPI C-states</hardwareSupport>"),
        B("capability.vmotionSupported", true),
        B("capability.storageVMotionSupported", true),
        P("config.hyperThread", "HostHyperThreadScheduleInfo", "<available>true</available><active>true</active><config>true</config>"),
        P("config.network", "HostNetworkInfo", networkInner),
        rich ? P("config.storageDevice.hostBusAdapter", "ArrayOfHostHostBusAdapter", """
            <HostHostBusAdapter xsi:type="HostBlockHba"><key>key-vim.host.BlockHba-vmhba0</key><device>vmhba0</device><bus>3</bus><status>unknown</status><model>Dell HBA355i Front</model><driver>lsi_msgpt35</driver><pci>0000:03:00.0</pci><storageProtocol>scsi</storageProtocol></HostHostBusAdapter>
            <HostHostBusAdapter xsi:type="HostFibreChannelHba"><key>key-vim.host.FibreChannelHba-vmhba2</key><device>vmhba2</device><bus>59</bus><status>online</status><model>QLogic QLE2772 Dual Port 32Gb Fibre Channel to PCIe Adapter</model><driver>qlnativefc</driver><pci>0000:3b:00.0</pci><storageProtocol>scsi</storageProtocol><portWorldWideName>2377900762156915201</portWorldWideName><nodeWorldWideName>2305843168118987265</nodeWorldWideName><portType>fabric</portType><speed>32</speed></HostHostBusAdapter>
            """) : "",
        rich ? P("config.storageDevice.scsiLun", "ArrayOfScsiLun", """
            <ScsiLun xsi:type="HostScsiDisk"><deviceName>/vmfs/devices/disks/naa.600a098038304437415d4b6a59684a52</deviceName><deviceType>disk</deviceType><key>key-vim.host.ScsiDisk-0200000000600a098038304437415d4b6a59684a524c554e202020</key><uuid>0200000000600a098038304437415d4b6a59684a524c554e202020</uuid><canonicalName>naa.600a098038304437415d4b6a59684a52</canonicalName><displayName>NETAPP Fibre Channel Disk (naa.600a098038304437415d4b6a59684a52)</displayName><lunType>disk</lunType><vendor>NETAPP  </vendor><model>LUN C-Mode      </model><revision>9130</revision><scsiLevel>6</scsiLevel><serialNumber>80D7A]KjYhJR</serialNumber><operationalState>ok</operationalState><queueDepth>64</queueDepth><vStorageSupport>vStorageSupported</vStorageSupport><capacity><blockSize>512</blockSize><block>4294967296</block></capacity><localDisk>false</localDisk><ssd>true</ssd></ScsiLun>
            <ScsiLun xsi:type="ScsiLun"><deviceName>/vmfs/devices/cdrom/mpx.vmhba1:C0:T0:L0</deviceName><deviceType>cdrom</deviceType><key>key-vim.host.ScsiLun-cdrom0</key><uuid>0005000000766d686261313a303a30</uuid><canonicalName>mpx.vmhba1:C0:T0:L0</canonicalName><lunType>cdrom</lunType><operationalState>ok</operationalState></ScsiLun>
            """) : "",
        rich ? P("config.storageDevice.multipathInfo", "HostMultipathInfo", """
            <lun><key>key-vim.host.MultipathInfo.LogicalUnit-0200000000600a098038304437415d4b6a59684a524c554e202020</key><id>0200000000600a098038304437415d4b6a59684a524c554e202020</id><lun>key-vim.host.ScsiDisk-0200000000600a098038304437415d4b6a59684a524c554e202020</lun>
              <path><key>key-vim.host.MultipathInfo.Path-vmhba2:C0:T0:L1</key><name>vmhba2:C0:T0:L1</name><pathState>active</pathState><state>active</state><isWorkingPath>true</isWorkingPath><adapter>key-vim.host.FibreChannelHba-vmhba2</adapter><lun>key-vim.host.MultipathInfo.LogicalUnit-0200000000600a098038304437415d4b6a59684a524c554e202020</lun></path>
              <path><key>key-vim.host.MultipathInfo.Path-vmhba2:C0:T1:L1</key><name>vmhba2:C0:T1:L1</name><pathState>standby</pathState><state>standby</state><isWorkingPath>false</isWorkingPath><adapter>key-vim.host.FibreChannelHba-vmhba2</adapter><lun>key-vim.host.MultipathInfo.LogicalUnit-0200000000600a098038304437415d4b6a59684a524c554e202020</lun></path>
              <policy xsi:type="HostMultipathInfoLogicalUnitPolicy"><policy>VMW_PSP_RR</policy></policy><storageArrayTypePolicy><policy>VMW_SATP_ALUA</policy></storageArrayTypePolicy></lun>
            <lun><key>key-vim.host.MultipathInfo.LogicalUnit-cdrom0</key><id>0005000000766d686261313a303a30</id><lun>key-vim.host.ScsiLun-cdrom0</lun><path><key>p</key><name>vmhba1:C0:T0:L0</name><pathState>active</pathState></path><policy xsi:type="HostMultipathInfoFixedLogicalUnitPolicy"><policy>VMW_PSP_FIXED</policy><prefer>vmhba1:C0:T0:L0</prefer></policy></lun>
            """) : "",
        rich ? P("config.fileSystemVolume.mountInfo", "ArrayOfHostFileSystemMountInfo", """
            <HostFileSystemMountInfo><mountInfo><path>/vmfs/volumes/650a1b2c-11aa22bb-33cc-b49691aa0001</path><accessMode>readWrite</accessMode><mounted>true</mounted><accessible>true</accessible></mountInfo>
              <volume xsi:type="HostVmfsVolume"><type>VMFS</type><name>ds-vmfs01</name><capacity>2198754820096</capacity><blockSizeMb>1</blockSizeMb><blockSize>1024</blockSize><maxBlocks>67108864</maxBlocks><majorVersion>6</majorVersion><version>6.82</version><uuid>650a1b2c-11aa22bb-33cc-b49691aa0001</uuid><extent><diskName>naa.600a098038304437415d4b6a59684a52</diskName><partition>1</partition></extent><vmfsUpgradable>false</vmfsUpgradable><ssd>true</ssd><local>false</local></volume></HostFileSystemMountInfo>
            <HostFileSystemMountInfo><mountInfo><path>/vmfs/volumes/a1b2c3d4-e5f60718</path><accessMode>readWrite</accessMode><mounted>true</mounted><accessible>true</accessible></mountInfo>
              <volume xsi:type="HostNasVolume"><type>NFS</type><name>nfs01</name><capacity>4398046511104</capacity><remoteHost>nas01.lab.local</remoteHost><remotePath>/vol/vmware</remotePath></volume></HostFileSystemMountInfo>
            """) : "",
        P("config.dateTimeInfo", "HostDateTimeInfo", "<timeZone><key>UTC</key><name>UTC</name><description>UTC</description><gmtOffset>0</gmtOffset></timeZone><systemClockProtocol>ntp</systemClockProtocol><ntpConfig><server>0.pool.ntp.org</server><server>1.pool.ntp.org</server></ntpConfig>"),
        P("config.powerSystemInfo", "PowerSystemInfo", "<currentPolicy><key>2</key><name>Balanced</name><shortName>dynamic</shortName><description>Reduce energy consumption with minimal performance compromise</description></currentPolicy>"),
        P("config.service", "HostServiceInfo", "<service><key>TSM-SSH</key><label>SSH</label><required>false</required><uninstallable>false</uninstallable><running>false</running><policy>off</policy></service><service><key>ntpd</key><label>NTP Daemon</label><required>false</required><uninstallable>false</uninstallable><running>true</running><ruleset>ntpClient</ruleset><policy>on</policy></service>"));

    public static string Hosts() => RetrieveResult(
        HostObject("host-20", "esx01.lab.local", "5JK8XG3", false, Esx01Network, true)
        + HostObject("host-21", "esx02.lab.local", "6LM9YH4", true, """
            <vswitch><name>vSwitch0</name><key>key-vim.host.VirtualSwitch-vSwitch0</key><numPorts>2560</numPorts><numPortsAvailable>2550</numPortsAvailable><mtu>1500</mtu><pnic>key-vim.host.PhysicalNic-vmnic0</pnic><spec><numPorts>2560</numPorts><policy><security><allowPromiscuous>false</allowPromiscuous><macChanges>false</macChanges><forgedTransmits>false</forgedTransmits></security><nicTeaming><policy>loadbalance_srcid</policy></nicTeaming></policy></spec></vswitch>
            <portgroup><key>key-vim.host.PortGroup-VM Network</key><spec><name>VM Network</name><vlanId>0</vlanId><vswitchName>vSwitch0</vswitchName><policy></policy></spec></portgroup>
            <pnic><key>key-vim.host.PhysicalNic-vmnic0</key><device>vmnic0</device><pci>0000:18:00.0</pci><driver>ixgben</driver><linkSpeed><speedMb>10000</speedMb><duplex>true</duplex></linkSpeed><mac>b4:96:91:bb:00:01</mac></pnic>
            <vnic><device>vmk0</device><key>key-vim.host.VirtualNic-vmk0</key><portgroup>Management Network</portgroup><spec><ip><dhcp>false</dhcp><ipAddress>10.0.10.22</ipAddress><subnetMask>255.255.255.0</subnetMask></ip><mac>00:50:56:6a:11:02</mac><mtu>1500</mtu></spec></vnic>
            <dnsConfig xsi:type="HostDnsConfig"><dhcp>false</dhcp><hostName>esx02</hostName><domainName>lab.local</domainName><address>10.0.0.10</address></dnsConfig>
            <ipRouteConfig xsi:type="HostIpRouteConfig"><defaultGateway>10.0.10.1</defaultGateway></ipRouteConfig>
            """, false));

    // ------------------------------------------------------------------ clusters, pools, datastores, DVS

    public static string ComputeResources() => RetrieveResult(O("ClusterComputeResource", "domain-c10",
        S("name", "Cluster01"),
        Ref("parent", "Folder", "group-h4"),
        P("overallStatus", "ManagedEntityStatus", "green"),
        P("configStatus", "ManagedEntityStatus", "green"),
        P("summary", "ClusterComputeResourceSummary", "<totalCpu>268800</totalCpu><totalMemory>549755813888</totalMemory><numCpuCores>96</numCpuCores><numCpuThreads>192</numCpuThreads><effectiveCpu>240000</effectiveCpu><effectiveMemory>480000</effectiveMemory><numHosts>2</numHosts><numEffectiveHosts>1</numEffectiveHosts><overallStatus>green</overallStatus><currentFailoverLevel>1</currentFailoverLevel><numVmotions>42</numVmotions><currentEVCModeKey>intel-icelake</currentEVCModeKey>"),
        P("configurationEx", "ClusterConfigInfoEx", """
            <dasConfig><enabled>true</enabled><vmMonitoring>vmMonitoringOnly</vmMonitoring><hostMonitoring>enabled</hostMonitoring><vmComponentProtecting>disabled</vmComponentProtecting><failoverLevel>1</failoverLevel>
              <admissionControlPolicy xsi:type="ClusterFailoverResourcesAdmissionControlPolicy"><cpuFailoverResourcesPercent>50</cpuFailoverResourcesPercent><memoryFailoverResourcesPercent>50</memoryFailoverResourcesPercent><failoverLevel>1</failoverLevel><autoComputePercentages>true</autoComputePercentages></admissionControlPolicy>
              <admissionControlEnabled>true</admissionControlEnabled>
              <defaultVmSettings><restartPriority>medium</restartPriority><isolationResponse>none</isolationResponse>
                <vmToolsMonitoringSettings><enabled>true</enabled><vmMonitoring>vmMonitoringOnly</vmMonitoring><clusterSettings>true</clusterSettings><failureInterval>30</failureInterval><minUpTime>120</minUpTime><maxFailures>3</maxFailures><maxFailureWindow>3600</maxFailureWindow></vmToolsMonitoringSettings></defaultVmSettings>
              <hBDatastoreCandidatePolicy>allFeasibleDsWithUserPreference</hBDatastoreCandidatePolicy></dasConfig>
            <dasVmConfig><key type="VirtualMachine">vm-100</key><restartPriority>high</restartPriority><dasSettings><restartPriority>high</restartPriority><isolationResponse>clusterIsolationResponse</isolationResponse></dasSettings></dasVmConfig>
            <drsConfig><enabled>true</enabled><enableVmBehaviorOverrides>true</enableVmBehaviorOverrides><defaultVmBehavior>fullyAutomated</defaultVmBehavior><vmotionRate>3</vmotionRate></drsConfig>
            <rule xsi:type="ClusterAntiAffinityRuleSpec"><key>1</key><status>green</status><enabled>true</enabled><name>separate-web-db</name><mandatory>false</mandatory><userCreated>true</userCreated><inCompliance>true</inCompliance><vm type="VirtualMachine">vm-100</vm><vm type="VirtualMachine">vm-101</vm></rule>
            <dpmConfigInfo><enabled>false</enabled><defaultDpmBehavior>automated</defaultDpmBehavior><hostPowerActionRate>3</hostPowerActionRate></dpmConfigInfo>
            <vmSwapPlacement>vmDirectory</vmSwapPlacement>
            """),
        Refs("host", ("HostSystem", "host-20"), ("HostSystem", "host-21")),
        Ref("resourcePool", "ResourcePool", "resgroup-11")));

    private static string PoolSummary(string name, long cpuRes, long memLimit, bool expandable) => $"""
        <name>{name}</name>
        <config><changeVersion>1</changeVersion>
          <cpuAllocation><reservation>{cpuRes}</reservation><expandableReservation>{(expandable ? "true" : "false")}</expandableReservation><limit>-1</limit><shares><shares>4000</shares><level>normal</level></shares><overheadLimit>-1</overheadLimit></cpuAllocation>
          <memoryAllocation><reservation>0</reservation><expandableReservation>true</expandableReservation><limit>{memLimit}</limit><shares><shares>163840</shares><level>normal</level></shares></memoryAllocation></config>
        <runtime><memory><reservationUsed>0</reservationUsed><reservationUsedForVm>0</reservationUsedForVm><unreservedForPool>412316860416</unreservedForPool><unreservedForVm>412316860416</unreservedForVm><overallUsage>8589934592</overallUsage><maxUsage>412316860416</maxUsage></memory>
          <cpu><reservationUsed>0</reservationUsed><reservationUsedForVm>0</reservationUsedForVm><unreservedForPool>240000</unreservedForPool><unreservedForVm>240000</unreservedForVm><overallUsage>1200</overallUsage><maxUsage>240000</maxUsage></cpu><overallStatus>green</overallStatus></runtime>
        <quickStats><overallCpuUsage>1200</overallCpuUsage><overallCpuDemand>1250</overallCpuDemand><guestMemoryUsage>1638</guestMemoryUsage><hostMemoryUsage>8250</hostMemoryUsage><distributedCpuEntitlement>1300</distributedCpuEntitlement><distributedMemoryEntitlement>8400</distributedMemoryEntitlement><staticCpuEntitlement>11200</staticCpuEntitlement><staticMemoryEntitlement>8300</staticMemoryEntitlement><privateMemory>8100</privateMemory><sharedMemory>12</sharedMemory><swappedMemory>0</swappedMemory><balloonedMemory>0</balloonedMemory><overheadMemory>60</overheadMemory><consumedOverheadMemory>45</consumedOverheadMemory><compressedMemory>0</compressedMemory></quickStats>
        <configuredMemoryMB>16384</configuredMemoryMB>
        """;

    public static string ResourcePools() => RetrieveResult(
        O("ResourcePool", "resgroup-11", S("name", "Resources"), Ref("parent", "ClusterComputeResource", "domain-c10"),
            Ref("owner", "ClusterComputeResource", "domain-c10"), Refs("vm", ("VirtualMachine", "vm-101")),
            Refs("resourcePool", ("ResourcePool", "resgroup-50")),
            P("summary", "ResourcePoolSummary", PoolSummary("Resources", 240000, -1, false)),
            P("overallStatus", "ManagedEntityStatus", "green"))
        + O("ResourcePool", "resgroup-50", S("name", "Prod"), Ref("parent", "ResourcePool", "resgroup-11"),
            Ref("owner", "ClusterComputeResource", "domain-c10"), Refs("vm", ("VirtualMachine", "vm-100")),
            P("resourcePool", "ArrayOfManagedObjectReference", ""),
            P("summary", "ResourcePoolSummary", PoolSummary("Prod", 2000, 32768, true)),
            P("overallStatus", "ManagedEntityStatus", "green")));

    public static string Datastores() => RetrieveResult(
        O("Datastore", "datastore-30", S("name", "ds-vmfs01"), Ref("parent", "Folder", "group-s5"),
            P("overallStatus", "ManagedEntityStatus", "green"), P("configStatus", "ManagedEntityStatus", "gray"),
            P("summary", "DatastoreSummary", """<datastore type="Datastore">datastore-30</datastore><name>ds-vmfs01</name><url>ds:///vmfs/volumes/650a1b2c-11aa22bb-33cc-b49691aa0001/</url><capacity>2198754820096</capacity><freeSpace>1099377410048</freeSpace><uncommitted>214748364800</uncommitted><accessible>true</accessible><multipleHostAccess>true</multipleHostAccess><type>VMFS</type><maintenanceMode>normal</maintenanceMode>"""),
            P("info", "VmfsDatastoreInfo", """
                <name>ds-vmfs01</name><url>/vmfs/volumes/650a1b2c-11aa22bb-33cc-b49691aa0001</url><freeSpace>1099377410048</freeSpace><maxFileSize>70368744177664</maxFileSize><maxVirtualDiskCapacity>68169720922112</maxVirtualDiskCapacity><maxMemoryFileSize>70368744177664</maxMemoryFileSize><timestamp>2026-10-06T08:59:00Z</timestamp><containerId>650a1b2c-11aa22bb-33cc-b49691aa0001</containerId><maxPhysicalRDMFileSize>70368744177664</maxPhysicalRDMFileSize><maxVirtualRDMFileSize>68169720922112</maxVirtualRDMFileSize>
                <vmfs><type>VMFS</type><name>ds-vmfs01</name><capacity>2198754820096</capacity><blockSizeMb>1</blockSizeMb><blockSize>1024</blockSize><unmapGranularity>1024</unmapGranularity><unmapPriority>low</unmapPriority><maxBlocks>67108864</maxBlocks><majorVersion>6</majorVersion><version>6.82</version><uuid>650a1b2c-11aa22bb-33cc-b49691aa0001</uuid><extent><diskName>naa.600a098038304437415d4b6a59684a52</diskName><partition>1</partition></extent><vmfsUpgradable>false</vmfsUpgradable><ssd>true</ssd><local>false</local></vmfs>
                """),
            P("host", "ArrayOfDatastoreHostMount", """<DatastoreHostMount><key type="HostSystem">host-20</key><mountInfo><path>/vmfs/volumes/650a1b2c</path><accessMode>readWrite</accessMode><mounted>true</mounted><accessible>true</accessible></mountInfo></DatastoreHostMount><DatastoreHostMount><key type="HostSystem">host-21</key><mountInfo><path>/vmfs/volumes/650a1b2c</path><accessMode>readWrite</accessMode><mounted>true</mounted><accessible>true</accessible></mountInfo></DatastoreHostMount>"""),
            Refs("vm", ("VirtualMachine", "vm-100"), ("VirtualMachine", "vm-101")),
            P("iormConfiguration", "StorageIORMInfo", "<enabled>true</enabled><congestionThresholdMode>automatic</congestionThresholdMode><congestionThreshold>30</congestionThreshold><percentOfPeakThroughput>90</percentOfPeakThroughput><statsCollectionEnabled>true</statsCollectionEnabled><reservationEnabled>true</reservationEnabled><statsAggregationDisabled>false</statsAggregationDisabled>"))
        + O("Datastore", "datastore-31", S("name", "nfs01"), Ref("parent", "Folder", "group-s5"),
            P("overallStatus", "ManagedEntityStatus", "yellow"), P("configStatus", "ManagedEntityStatus", "gray"),
            P("summary", "DatastoreSummary", """<datastore type="Datastore">datastore-31</datastore><name>nfs01</name><url>ds:///vmfs/volumes/a1b2c3d4-e5f60718/</url><capacity>4398046511104</capacity><freeSpace>439804651110</freeSpace><uncommitted>0</uncommitted><accessible>true</accessible><multipleHostAccess>true</multipleHostAccess><type>NFS</type><maintenanceMode>normal</maintenanceMode>"""),
            P("info", "NasDatastoreInfo", """<name>nfs01</name><url>/vmfs/volumes/a1b2c3d4-e5f60718</url><freeSpace>439804651110</freeSpace><maxFileSize>62672162783232</maxFileSize><nas><type>NFS</type><name>nfs01</name><capacity>4398046511104</capacity><remoteHost>nas01.lab.local</remoteHost><remotePath>/vol/vmware</remotePath><userName></userName><remoteHostNames>nas01.lab.local</remoteHostNames></nas>"""),
            P("host", "ArrayOfDatastoreHostMount", """<DatastoreHostMount><key type="HostSystem">host-20</key><mountInfo><path>/vmfs/volumes/a1b2c3d4-e5f60718</path><accessMode>readWrite</accessMode><mounted>true</mounted><accessible>true</accessible></mountInfo></DatastoreHostMount>"""),
            Refs("vm", ("VirtualMachine", "vm-100"), ("VirtualMachine", "vm-102")),
            P("iormConfiguration", "StorageIORMInfo", "<enabled>false</enabled><congestionThreshold>30</congestionThreshold>")));

    private const string DvsPortSetting = """
        <blocked><inherited>false</inherited><value>false</value></blocked>
        <inShapingPolicy><inherited>false</inherited><enabled><inherited>false</inherited><value>false</value></enabled><averageBandwidth><inherited>false</inherited><value>100000000</value></averageBandwidth><peakBandwidth><inherited>false</inherited><value>100000000</value></peakBandwidth><burstSize><inherited>false</inherited><value>104857600</value></burstSize></inShapingPolicy>
        <outShapingPolicy><inherited>false</inherited><enabled><inherited>false</inherited><value>false</value></enabled><averageBandwidth><inherited>false</inherited><value>100000000</value></averageBandwidth><peakBandwidth><inherited>false</inherited><value>100000000</value></peakBandwidth><burstSize><inherited>false</inherited><value>104857600</value></burstSize></outShapingPolicy>
        """;

    public static string Dvs() => RetrieveResult(O("VmwareDistributedVirtualSwitch", "dvs-60",
        S("name", "DSwitch01"), Ref("parent", "Folder", "group-n6"),
        S("uuid", "50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8"),
        P("overallStatus", "ManagedEntityStatus", "green"),
        P("summary", "DVSSummary", """<name>DSwitch01</name><uuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</uuid><numPorts>32</numPorts><productInfo><name>DVS</name><vendor>VMware, Inc.</vendor><version>8.0.0</version><build>22380479</build></productInfo><hostMember type="HostSystem">host-20</hostMember><hostMember type="HostSystem">host-21</hostMember><vm type="VirtualMachine">vm-100</vm><vm type="VirtualMachine">vm-101</vm><host type="HostSystem">host-20</host><portgroupName>PG-App-100</portgroupName><portgroupName>DSwitch01-DVUplinks-60</portgroupName><description>Production DVS</description><contact><name>NetOps</name><contact>netops@lab.local</contact></contact><numHosts>2</numHosts>"""),
        P("config", "VMwareDVSConfigInfo", $"""
            <uuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</uuid><name>DSwitch01</name><numStandalonePorts>0</numStandalonePorts><numPorts>32</numPorts><maxPorts>2147483647</maxPorts>
            <uplinkPortPolicy xsi:type="DVSNameArrayUplinkPortPolicy"><uplinkPortName>Uplink 1</uplinkPortName><uplinkPortName>Uplink 2</uplinkPortName></uplinkPortPolicy><uplinkPortgroup type="DistributedVirtualPortgroup">dvportgroup-62</uplinkPortgroup>
            <defaultPortConfig xsi:type="VMwareDVSPortSetting">{DvsPortSetting}<vlan xsi:type="VmwareDistributedVirtualSwitchVlanIdSpec"><inherited>false</inherited><vlanId>0</vlanId></vlan></defaultPortConfig>
            <host><config><host type="HostSystem">host-20</host><maxProxySwitchPorts>512</maxProxySwitchPorts></config><productInfo><name>DVS</name><vendor>VMware, Inc.</vendor><version>8.0.0</version></productInfo><uplinkPortKey>16</uplinkPortKey></host>
            <productInfo><name>DVS</name><vendor>VMware, Inc.</vendor><version>8.0.0</version><build>22380479</build></productInfo>
            <configVersion>12</configVersion><contact><name>NetOps</name><contact>netops@lab.local</contact></contact><description>Production DVS</description>
            <createTime>2024-02-10T14:22:31.000Z</createTime><networkResourceManagementEnabled>true</networkResourceManagementEnabled>
            <linkDiscoveryProtocolConfig><protocol>cdp</protocol><operation>listen</operation></linkDiscoveryProtocolConfig>
            <maxMtu>9000</maxMtu>
            <lacpGroupConfig><key>lag1</key><name>lag1</name><mode>active</mode><uplinkNum>2</uplinkNum><loadbalanceAlgorithm>srcDestIpTcpUdpPortVlan</loadbalanceAlgorithm></lacpGroupConfig>
            <lacpApiVersion>multipleLag</lacpApiVersion>
            """)));

    public static string Portgroups() => RetrieveResult(
        O("DistributedVirtualPortgroup", "dvportgroup-61", S("name", "PG-App-100"), S("key", "dvportgroup-61"),
            P("config", "DVPortgroupConfigInfo", $"""
                <key>dvportgroup-61</key><name>PG-App-100</name><numPorts>8</numPorts><distributedVirtualSwitch type="VmwareDistributedVirtualSwitch">dvs-60</distributedVirtualSwitch>
                <defaultPortConfig xsi:type="VMwareDVSPortSetting">{DvsPortSetting}
                  <vlan xsi:type="VmwareDistributedVirtualSwitchVlanIdSpec"><inherited>false</inherited><vlanId>100</vlanId></vlan>
                  <uplinkTeamingPolicy><inherited>false</inherited><policy><inherited>false</inherited><value>loadbalance_srcid</value></policy><reversePolicy><inherited>false</inherited><value>true</value></reversePolicy><notifySwitches><inherited>false</inherited><value>true</value></notifySwitches><rollingOrder><inherited>false</inherited><value>false</value></rollingOrder>
                    <failureCriteria><inherited>false</inherited><checkSpeed><inherited>false</inherited><value>minimum</value></checkSpeed><speed><inherited>false</inherited><value>10</value></speed><checkDuplex><inherited>false</inherited><value>false</value></checkDuplex><fullDuplex><inherited>false</inherited><value>false</value></fullDuplex><checkErrorPercent><inherited>false</inherited><value>false</value></checkErrorPercent><percentage><inherited>false</inherited><value>0</value></percentage><checkBeacon><inherited>false</inherited><value>false</value></checkBeacon></failureCriteria>
                    <uplinkPortOrder><inherited>false</inherited><activeUplinkPort>Uplink 1</activeUplinkPort><standbyUplinkPort>Uplink 2</standbyUplinkPort></uplinkPortOrder></uplinkTeamingPolicy>
                  <securityPolicy><inherited>false</inherited><allowPromiscuous><inherited>false</inherited><value>false</value></allowPromiscuous><macChanges><inherited>false</inherited><value>false</value></macChanges><forgedTransmits><inherited>false</inherited><value>false</value></forgedTransmits></securityPolicy>
                  <macManagementPolicy><inherited>false</inherited><allowPromiscuous>false</allowPromiscuous><macChanges>false</macChanges><forgedTransmits>false</forgedTransmits></macManagementPolicy>
                </defaultPortConfig>
                <description></description><type>earlyBinding</type>
                <policy xsi:type="VMwareDVSPortgroupPolicy"><blockOverrideAllowed>true</blockOverrideAllowed><shapingOverrideAllowed>false</shapingOverrideAllowed><vendorConfigOverrideAllowed>false</vendorConfigOverrideAllowed><livePortMovingAllowed>false</livePortMovingAllowed><portConfigResetAtDisconnect>true</portConfigResetAtDisconnect><vlanOverrideAllowed>false</vlanOverrideAllowed><uplinkTeamingOverrideAllowed>false</uplinkTeamingOverrideAllowed><securityPolicyOverrideAllowed>false</securityPolicyOverrideAllowed></policy>
                <autoExpand>true</autoExpand><uplink>false</uplink>
                """))
        + O("DistributedVirtualPortgroup", "dvportgroup-62", S("name", "DSwitch01-DVUplinks-60"), S("key", "dvportgroup-62"),
            P("config", "DVPortgroupConfigInfo", """
                <key>dvportgroup-62</key><name>DSwitch01-DVUplinks-60</name><numPorts>4</numPorts><distributedVirtualSwitch type="VmwareDistributedVirtualSwitch">dvs-60</distributedVirtualSwitch>
                <defaultPortConfig xsi:type="VMwareDVSPortSetting"><vlan xsi:type="VmwareDistributedVirtualSwitchTrunkVlanSpec"><inherited>false</inherited><vlanId><start>0</start><end>4094</end></vlanId></vlan></defaultPortConfig>
                <type>earlyBinding</type><uplink>true</uplink>
                """)));

    // ------------------------------------------------------------------ VMs

    private const string Web01Hardware = """
        <numCPU>4</numCPU><numCoresPerSocket>2</numCoresPerSocket><autoCoresPerSocket>false</autoCoresPerSocket><memoryMB>8192</memoryMB><virtualICH7MPresent>false</virtualICH7MPresent><virtualSMCPresent>false</virtualSMCPresent>
        <device xsi:type="VirtualIDEController"><key>200</key><deviceInfo><label>IDE 0</label><summary>IDE 0</summary></deviceInfo><busNumber>0</busNumber><device>3002</device></device>
        <device xsi:type="VirtualPCIController"><key>100</key><deviceInfo><label>PCI controller 0</label><summary>PCI controller 0</summary></deviceInfo><busNumber>0</busNumber><device>500</device><device>12000</device><device>1000</device><device>4000</device><device>4001</device></device>
        <device xsi:type="VirtualMachineVideoCard"><key>500</key><deviceInfo><label>Video card </label><summary>Video card</summary></deviceInfo><controllerKey>100</controllerKey><unitNumber>0</unitNumber><videoRamSizeInKB>8192</videoRamSizeInKB><numDisplays>2</numDisplays><useAutoDetect>false</useAutoDetect><enable3DSupport>false</enable3DSupport><use3dRenderer>automatic</use3dRenderer><graphicsMemorySizeInKB>262144</graphicsMemorySizeInKB></device>
        <device xsi:type="ParaVirtualSCSIController"><key>1000</key><deviceInfo><label>SCSI controller 0</label><summary>VMware paravirtual SCSI</summary></deviceInfo><slotInfo xsi:type="VirtualDevicePciBusSlotInfo"><pciSlotNumber>160</pciSlotNumber></slotInfo><controllerKey>100</controllerKey><unitNumber>3</unitNumber><busNumber>0</busNumber><device>2000</device><device>2001</device><hotAddRemove>true</hotAddRemove><sharedBus>noSharing</sharedBus><scsiCtlrUnitNumber>7</scsiCtlrUnitNumber></device>
        <device xsi:type="VirtualUSBController"><key>7000</key><deviceInfo><label>USB controller </label><summary>Auto connect Disabled</summary></deviceInfo><controllerKey>100</controllerKey><unitNumber>22</unitNumber><busNumber>0</busNumber><device>13000</device><autoConnectDevices>false</autoConnectDevices><ehciEnabled>true</ehciEnabled></device>
        <device xsi:type="VirtualDisk"><key>2000</key><deviceInfo><label>Hard disk 1</label><summary>52,428,800 KB</summary></deviceInfo>
          <backing xsi:type="VirtualDiskFlatVer2BackingInfo"><fileName>[ds-vmfs01] web01/web01-000002.vmdk</fileName><datastore type="Datastore">datastore-30</datastore><backingObjectId></backingObjectId><diskMode>persistent</diskMode><split>false</split><writeThrough>false</writeThrough><thinProvisioned>true</thinProvisioned><eagerlyScrub>false</eagerlyScrub><uuid>6000C29a-1b2c-3d4e-5f60-718293a4b5c6</uuid><contentId>8d2f0a3b1c4e5f6a7b8c9d0e1f2a3b4c</contentId><changeId>52 1a 2b 3c-4d 5e 6f 70/1</changeId>
            <parent><fileName>[ds-vmfs01] web01/web01-000001.vmdk</fileName><datastore type="Datastore">datastore-30</datastore><diskMode>persistent</diskMode><thinProvisioned>true</thinProvisioned><uuid>6000C29a-1b2c-3d4e-5f60-718293a4b5c6</uuid>
              <parent><fileName>[ds-vmfs01] web01/web01.vmdk</fileName><datastore type="Datastore">datastore-30</datastore><diskMode>persistent</diskMode><thinProvisioned>true</thinProvisioned></parent></parent>
            <deltaDiskFormat>seSparseFormat</deltaDiskFormat><digestEnabled>false</digestEnabled><sharing>sharingNone</sharing></backing>
          <controllerKey>1000</controllerKey><unitNumber>0</unitNumber><capacityInKB>52428800</capacityInKB><capacityInBytes>53687091200</capacityInBytes>
          <shares><shares>1000</shares><level>normal</level></shares><storageIOAllocation><limit>-1</limit><shares><shares>1000</shares><level>normal</level></shares><reservation>0</reservation></storageIOAllocation>
          <diskObjectId>100-2000</diskObjectId><nativeUnmanagedLinkedClone>false</nativeUnmanagedLinkedClone></device>
        <device xsi:type="VirtualDisk"><key>2001</key><deviceInfo><label>Hard disk 2</label><summary>104,857,600 KB</summary></deviceInfo>
          <backing xsi:type="VirtualDiskFlatVer2BackingInfo"><fileName>[nfs01] web01/web01_1.vmdk</fileName><datastore type="Datastore">datastore-31</datastore><diskMode>independent_persistent</diskMode><split>false</split><writeThrough>false</writeThrough><thinProvisioned>false</thinProvisioned><eagerlyScrub>true</eagerlyScrub><uuid>6000C29b-2c3d-4e5f-6071-8293a4b5c6d7</uuid><digestEnabled>false</digestEnabled><sharing>sharingMultiWriter</sharing></backing>
          <controllerKey>1000</controllerKey><unitNumber>1</unitNumber><capacityInKB>104857600</capacityInKB><capacityInBytes>107374182400</capacityInBytes>
          <storageIOAllocation><limit>500</limit><shares><shares>2000</shares><level>high</level></shares><reservation>0</reservation></storageIOAllocation></device>
        <device xsi:type="VirtualCdrom"><key>3002</key><deviceInfo><label>CD/DVD drive 1</label><summary>ISO [ds-vmfs01] iso/ubuntu-22.04.iso</summary></deviceInfo>
          <backing xsi:type="VirtualCdromIsoBackingInfo"><fileName>[ds-vmfs01] iso/ubuntu-22.04.iso</fileName><datastore type="Datastore">datastore-30</datastore></backing>
          <connectable><startConnected>false</startConnected><allowGuestControl>true</allowGuestControl><connected>true</connected><status>ok</status></connectable><controllerKey>200</controllerKey><unitNumber>0</unitNumber></device>
        <device xsi:type="VirtualVmxnet3"><key>4000</key><deviceInfo><label>Network adapter 1</label><summary>DVSwitch: 50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</summary></deviceInfo>
          <backing xsi:type="VirtualEthernetCardDistributedVirtualPortBackingInfo"><port><switchUuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</switchUuid><portgroupKey>dvportgroup-61</portgroupKey><portKey>3</portKey><connectionCookie>1234567890</connectionCookie></port></backing>
          <connectable><migrateConnect>unset</migrateConnect><startConnected>true</startConnected><allowGuestControl>true</allowGuestControl><connected>true</connected><status>ok</status></connectable>
          <slotInfo xsi:type="VirtualDevicePciBusSlotInfo"><pciSlotNumber>192</pciSlotNumber></slotInfo><controllerKey>100</controllerKey><unitNumber>7</unitNumber>
          <addressType>assigned</addressType><macAddress>00:50:56:aa:bb:01</macAddress><wakeOnLanEnabled>true</wakeOnLanEnabled><resourceAllocation><reservation>0</reservation><share><shares>50</shares><level>normal</level></share><limit>-1</limit></resourceAllocation><uptCompatibilityEnabled>true</uptCompatibilityEnabled></device>
        <device xsi:type="VirtualE1000e"><key>4001</key><deviceInfo><label>Network adapter 2</label><summary>VM Network</summary></deviceInfo>
          <backing xsi:type="VirtualEthernetCardNetworkBackingInfo"><deviceName>VM Network</deviceName><useAutoDetect>false</useAutoDetect><network type="Network">network-40</network></backing>
          <connectable><startConnected>false</startConnected><allowGuestControl>true</allowGuestControl><connected>false</connected><status>untried</status></connectable>
          <controllerKey>100</controllerKey><unitNumber>8</unitNumber><addressType>manual</addressType><macAddress>00:50:56:3f:00:02</macAddress><wakeOnLanEnabled>true</wakeOnLanEnabled></device>
        <device xsi:type="VirtualUSB"><key>13000</key><deviceInfo><label>USB 1</label><summary>Autoconnect Device</summary></deviceInfo>
          <backing xsi:type="VirtualUSBUSBBackingInfo"><deviceName>path:1/0/1 version:2</deviceName></backing><connectable><startConnected>true</startConnected><allowGuestControl>false</allowGuestControl><connected>true</connected></connectable>
          <controllerKey>7000</controllerKey><unitNumber>0</unitNumber><connected>true</connected><vendor>1921</vendor><product>21889</product><family>storage</family><speed>high</speed></device>
        """;

    private const string Web01Guest = """
        <toolsStatus>toolsOk</toolsStatus><toolsVersionStatus>guestToolsCurrent</toolsVersionStatus><toolsVersionStatus2>guestToolsCurrent</toolsVersionStatus2><toolsRunningStatus>guestToolsRunning</toolsRunningStatus><toolsVersion>12352</toolsVersion><toolsInstallType>guestToolsTypeOpenVMTools</toolsInstallType>
        <guestId>ubuntu64Guest</guestId><guestFamily>linuxGuest</guestFamily><guestFullName>Ubuntu Linux (64-bit)</guestFullName><guestDetailedData>bitness='64' distroName='Ubuntu' distroVersion='22.04' familyName='Linux' kernelVersion='5.15.0-91-generic' prettyName='Ubuntu 22.04.3 LTS'</guestDetailedData>
        <hostName>web01.lab.local</hostName><ipAddress>10.0.100.11</ipAddress>
        <net><network>PG-App-100</network><ipAddress>10.0.100.11</ipAddress><ipAddress>fe80::250:56ff:feaa:bb01</ipAddress><macAddress>00:50:56:aa:bb:01</macAddress><connected>true</connected><deviceConfigId>4000</deviceConfigId>
          <ipConfig><ipAddress><ipAddress>10.0.100.11</ipAddress><prefixLength>24</prefixLength><state>preferred</state></ipAddress><ipAddress><ipAddress>fe80::250:56ff:feaa:bb01</ipAddress><prefixLength>64</prefixLength><state>unknown</state></ipAddress></ipConfig></net>
        <net><network>VM Network</network><macAddress>00:50:56:3f:00:02</macAddress><connected>false</connected><deviceConfigId>4001</deviceConfigId></net>
        <disk><diskPath>/</diskPath><capacity>52710469632</capacity><freeSpace>31626281779</freeSpace><filesystemType>ext4</filesystemType><mappings><key>2000</key></mappings></disk>
        <disk><diskPath>/boot</diskPath><capacity>1023303680</capacity><freeSpace>767477760</freeSpace><filesystemType>ext4</filesystemType><mappings><key>2000</key></mappings></disk>
        <screen><width>1280</width><height>800</height></screen>
        <guestState>running</guestState><appHeartbeatStatus>appStatusGray</appHeartbeatStatus><guestKernelCrashed>false</guestKernelCrashed><appState>none</appState><guestOperationsReady>true</guestOperationsReady><interactiveGuestOperationsReady>false</interactiveGuestOperationsReady><guestStateChangeSupported>true</guestStateChangeSupported>
        <hwVersion>vmx-19</hwVersion><customizationInfo><customizationStatus>TOOLSDEPLOYPKG_IDLE</customizationStatus></customizationInfo>
        """;

    private static string Allocation(long reservation, long limit, int shares, string level) =>
        $"<reservation>{reservation}</reservation><expandableReservation>false</expandableReservation><limit>{limit}</limit><shares><shares>{shares}</shares><level>{level}</level></shares>";

    public static string Web01 => O("VirtualMachine", "vm-100",
        S("name", "web01"),
        Ref("parent", "Folder", "group-v100"),
        Ref("resourcePool", "ResourcePool", "resgroup-50"),
        Refs("datastore", ("Datastore", "datastore-30"), ("Datastore", "datastore-31")),
        Refs("network", ("DistributedVirtualPortgroup", "dvportgroup-61"), ("Network", "network-40")),
        P("overallStatus", "ManagedEntityStatus", "green"),
        P("configStatus", "ManagedEntityStatus", "green"),
        P("guestHeartbeatStatus", "ManagedEntityStatus", "green"),
        B("config.template", false),
        S("config.uuid", "4231a1b2-c3d4-e5f6-0718-293a4b5c6d7e"),
        S("config.instanceUuid", "5031a1b2-c3d4-e5f6-0718-293a4b5c6d7e"),
        S("config.guestFullName", "Ubuntu Linux (64-bit)"),
        S("config.guestId", "ubuntu64Guest"),
        S("config.version", "vmx-19"),
        S("config.changeVersion", "2026-09-30T12:01:02.345678Z"),
        P("config.createDate", "xsd:dateTime", "2024-03-12T10:15:00.000Z"),
        S("config.annotation", "Web front end &amp; API"),
        S("config.firmware", "efi"),
        P("config.bootOptions", "VirtualMachineBootOptions", "<bootDelay>5000</bootDelay><enterBIOSSetup>false</enterBIOSSetup><efiSecureBootEnabled>true</efiSecureBootEnabled><bootRetryEnabled>true</bootRetryEnabled><bootRetryDelay>10000</bootRetryDelay><networkBootProtocol>ipv4</networkBootProtocol>"),
        P("config.hardware", "VirtualHardware", Web01Hardware),
        P("config.cpuAllocation", "ResourceAllocationInfo", Allocation(1000, -1, 4000, "normal")),
        P("config.memoryAllocation", "ResourceAllocationInfo", Allocation(2048, 16384, 81920, "normal")),
        B("config.cpuHotAddEnabled", true),
        B("config.cpuHotRemoveEnabled", false),
        B("config.memoryHotAddEnabled", true),
        B("config.memoryReservationLockedToMax", false),
        P("config.latencySensitivity", "LatencySensitivity", "<level>normal</level>"),
        B("config.changeTrackingEnabled", true),
        P("config.files", "VirtualMachineFileInfo", "<vmPathName>[ds-vmfs01] web01/web01.vmx</vmPathName><snapshotDirectory>[ds-vmfs01] web01/</snapshotDirectory><suspendDirectory>[ds-vmfs01] web01/</suspendDirectory><logDirectory>[ds-vmfs01] web01/</logDirectory>"),
        P("config.tools", "ToolsConfigInfo", "<toolsVersion>12352</toolsVersion><toolsInstallType>guestToolsTypeOpenVMTools</toolsInstallType><afterPowerOn>true</afterPowerOn><afterResume>true</afterResume><beforeGuestStandby>true</beforeGuestStandby><beforeGuestShutdown>true</beforeGuestShutdown><toolsUpgradePolicy>manual</toolsUpgradePolicy><syncTimeWithHostAllowed>true</syncTimeWithHostAllowed><syncTimeWithHost>false</syncTimeWithHost>"),
        P("config.scheduledHardwareUpgradeInfo", "ScheduledHardwareUpgradeInfo", "<upgradePolicy>never</upgradePolicy><scheduledHardwareUpgradeStatus>none</scheduledHardwareUpgradeStatus>"),
        P("runtime.powerState", "VirtualMachinePowerState", "poweredOn"),
        P("runtime.connectionState", "VirtualMachineConnectionState", "connected"),
        Ref("runtime.host", "HostSystem", "host-20"),
        P("runtime.bootTime", "xsd:dateTime", "2026-09-20T08:00:00.000Z"),
        B("runtime.consolidationNeeded", true),
        P("runtime.faultToleranceState", "VirtualMachineFaultToleranceState", "notConfigured"),
        S("runtime.minRequiredEVCModeKey", "intel-sandybridge"),
        P("runtime.dasVmProtection", "VirtualMachineRuntimeInfoDasProtectionState", "<dasProtected>true</dasProtected>"),
        P("runtime.maxCpuUsage", "xsd:int", "11200"),
        P("runtime.maxMemoryUsage", "xsd:int", "8192"),
        P("runtime.memoryOverhead", "xsd:long", "62914560"),
        P("guest", "GuestInfo", Web01Guest),
        P("summary.storage", "VirtualMachineStorageSummary", "<committed>32749125632</committed><uncommitted>128849018880</uncommitted><unshared>32212254720</unshared><timestamp>2026-10-06T08:55:00Z</timestamp>"),
        P("summary.quickStats", "VirtualMachineQuickStats", "<overallCpuUsage>840</overallCpuUsage><overallCpuDemand>900</overallCpuDemand><guestMemoryUsage>1638</guestMemoryUsage><hostMemoryUsage>8250</hostMemoryUsage><guestHeartbeatStatus>green</guestHeartbeatStatus><distributedCpuEntitlement>950</distributedCpuEntitlement><distributedMemoryEntitlement>4096</distributedMemoryEntitlement><staticCpuEntitlement>11200</staticCpuEntitlement><staticMemoryEntitlement>8300</staticMemoryEntitlement><grantedMemory>8192</grantedMemory><privateMemory>8100</privateMemory><sharedMemory>12</sharedMemory><swappedMemory>64</swappedMemory><balloonedMemory>256</balloonedMemory><consumedOverheadMemory>45</consumedOverheadMemory><ftLogBandwidth>-1</ftLogBandwidth><ftSecondaryLatency>-1</ftSecondaryLatency><ftLatencyStatus>gray</ftLatencyStatus><compressedMemory>0</compressedMemory><uptimeSeconds>1386000</uptimeSeconds><ssdSwappedMemory>0</ssdSwappedMemory>"),
        P("snapshot", "VirtualMachineSnapshotInfo", """
            <currentSnapshot type="VirtualMachineSnapshot">snapshot-2</currentSnapshot>
            <rootSnapshotList><snapshot type="VirtualMachineSnapshot">snapshot-1</snapshot><vm type="VirtualMachine">vm-100</vm><name>Before patch</name><description>Pre-patch baseline</description><id>1</id><createTime>2026-08-01T22:00:00.000Z</createTime><state>poweredOff</state><quiesced>false</quiesced><replaySupported>false</replaySupported>
              <childSnapshotList><snapshot type="VirtualMachineSnapshot">snapshot-2</snapshot><vm type="VirtualMachine">vm-100</vm><name>After patch</name><description></description><id>2</id><createTime>2026-08-02T06:30:00.000Z</createTime><state>poweredOn</state><quiesced>true</quiesced><replaySupported>false</replaySupported></childSnapshotList>
            </rootSnapshotList>
            """),
        P("layoutEx.file", "ArrayOfVirtualMachineFileLayoutExFileInfo", """
            <VirtualMachineFileLayoutExFileInfo><key>0</key><name>[ds-vmfs01] web01/web01.vmx</name><type>config</type><size>3712</size><uniqueSize>3712</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>1</key><name>[ds-vmfs01] web01/web01.vmsd</name><type>snapshotList</type><size>914</size><uniqueSize>914</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>2</key><name>[ds-vmfs01] web01/web01.vmdk</name><type>diskDescriptor</type><size>600</size><uniqueSize>600</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>3</key><name>[ds-vmfs01] web01/web01-flat.vmdk</name><type>diskExtent</type><size>10737418240</size><uniqueSize>10737418240</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>4</key><name>[ds-vmfs01] web01/web01-000001.vmdk</name><type>diskDescriptor</type><size>400</size><uniqueSize>400</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>5</key><name>[ds-vmfs01] web01/web01-000001-sesparse.vmdk</name><type>diskExtent</type><size>1073741824</size><uniqueSize>1073741824</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>6</key><name>[ds-vmfs01] web01/web01-000002.vmdk</name><type>diskDescriptor</type><size>400</size><uniqueSize>400</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>7</key><name>[ds-vmfs01] web01/web01-000002-sesparse.vmdk</name><type>diskExtent</type><size>536870912</size><uniqueSize>536870912</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>8</key><name>[ds-vmfs01] web01/web01-Snapshot1.vmsn</name><type>snapshotData</type><size>2097152</size><uniqueSize>2097152</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>9</key><name>[ds-vmfs01] web01/web01-Snapshot2.vmsn</name><type>snapshotData</type><size>3145728</size><uniqueSize>3145728</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>10</key><name>[nfs01] web01/web01_1.vmdk</name><type>diskDescriptor</type><size>500</size><uniqueSize>500</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            <VirtualMachineFileLayoutExFileInfo><key>11</key><name>[nfs01] web01/web01_1-flat.vmdk</name><type>diskExtent</type><size>21474836480</size><uniqueSize>21474836480</uniqueSize><accessible>true</accessible></VirtualMachineFileLayoutExFileInfo>
            """),
        P("layoutEx.disk", "ArrayOfVirtualMachineFileLayoutExDiskLayout", """
            <VirtualMachineFileLayoutExDiskLayout><key>2000</key><chain><fileKey>2</fileKey><fileKey>3</fileKey></chain><chain><fileKey>4</fileKey><fileKey>5</fileKey></chain><chain><fileKey>6</fileKey><fileKey>7</fileKey></chain></VirtualMachineFileLayoutExDiskLayout>
            <VirtualMachineFileLayoutExDiskLayout><key>2001</key><chain><fileKey>10</fileKey><fileKey>11</fileKey></chain></VirtualMachineFileLayoutExDiskLayout>
            """),
        P("layoutEx.snapshot", "ArrayOfVirtualMachineFileLayoutExSnapshotLayout", """
            <VirtualMachineFileLayoutExSnapshotLayout><key type="VirtualMachineSnapshot">snapshot-1</key><dataKey>8</dataKey><memoryKey>-1</memoryKey><disk><key>2000</key><chain><fileKey>2</fileKey><fileKey>3</fileKey></chain></disk></VirtualMachineFileLayoutExSnapshotLayout>
            <VirtualMachineFileLayoutExSnapshotLayout><key type="VirtualMachineSnapshot">snapshot-2</key><dataKey>9</dataKey><memoryKey>-1</memoryKey><disk><key>2000</key><chain><fileKey>2</fileKey><fileKey>3</fileKey></chain><chain><fileKey>4</fileKey><fileKey>5</fileKey></chain></disk></VirtualMachineFileLayoutExSnapshotLayout>
            """));

    public static string Db01 => O("VirtualMachine", "vm-101",
        S("name", "db01"),
        Ref("parent", "Folder", "group-v3"),
        Ref("resourcePool", "ResourcePool", "resgroup-11"),
        Refs("datastore", ("Datastore", "datastore-30")),
        Refs("network", ("DistributedVirtualPortgroup", "dvportgroup-61")),
        P("overallStatus", "ManagedEntityStatus", "green"),
        P("configStatus", "ManagedEntityStatus", "yellow"),
        P("guestHeartbeatStatus", "ManagedEntityStatus", "gray"),
        B("config.template", false),
        S("config.uuid", "4231b2c3-d4e5-f607-1829-3a4b5c6d7e8f"),
        S("config.instanceUuid", "5031b2c3-d4e5-f607-1829-3a4b5c6d7e8f"),
        S("config.guestFullName", "Microsoft Windows Server 2022 (64-bit)"),
        S("config.version", "vmx-15"),
        S("config.changeVersion", "2026-07-01T09:00:00.000000Z"),
        P("config.createDate", "xsd:dateTime", "2022-01-05T09:30:00.000Z"),
        S("config.firmware", "bios"),
        P("config.bootOptions", "VirtualMachineBootOptions", "<bootDelay>0</bootDelay><enterBIOSSetup>false</enterBIOSSetup><efiSecureBootEnabled>false</efiSecureBootEnabled><bootRetryEnabled>false</bootRetryEnabled><bootRetryDelay>10000</bootRetryDelay>"),
        P("config.hardware", "VirtualHardware", """
            <numCPU>2</numCPU><numCoresPerSocket>1</numCoresPerSocket><memoryMB>16384</memoryMB>
            <device xsi:type="VirtualLsiLogicSASController"><key>1000</key><deviceInfo><label>SCSI controller 0</label><summary>LSI Logic SAS</summary></deviceInfo><controllerKey>100</controllerKey><unitNumber>3</unitNumber><busNumber>0</busNumber><device>2000</device><hotAddRemove>true</hotAddRemove><sharedBus>noSharing</sharedBus><scsiCtlrUnitNumber>7</scsiCtlrUnitNumber></device>
            <device xsi:type="VirtualDisk"><key>2000</key><deviceInfo><label>Hard disk 1</label><summary>83,886,080 KB</summary></deviceInfo>
              <backing xsi:type="VirtualDiskFlatVer2BackingInfo"><fileName>[ds-vmfs01] db01/db01.vmdk</fileName><datastore type="Datastore">datastore-30</datastore><diskMode>persistent</diskMode><split>false</split><writeThrough>false</writeThrough><thinProvisioned>false</thinProvisioned><eagerlyScrub>false</eagerlyScrub><uuid>6000C29c-3d4e-5f60-7182-93a4b5c6d7e8</uuid><sharing>sharingNone</sharing></backing>
              <controllerKey>1000</controllerKey><unitNumber>0</unitNumber><capacityInKB>83886080</capacityInKB><capacityInBytes>85899345920</capacityInBytes></device>
            <device xsi:type="VirtualVmxnet3"><key>4000</key><deviceInfo><label>Network adapter 1</label><summary>DVSwitch</summary></deviceInfo>
              <backing xsi:type="VirtualEthernetCardDistributedVirtualPortBackingInfo"><port><switchUuid>50 1a 2b 3c 4d 5e 6f 70-81 92 a3 b4 c5 d6 e7 f8</switchUuid><portgroupKey>dvportgroup-61</portgroupKey><portKey>4</portKey></port></backing>
              <connectable><startConnected>true</startConnected><allowGuestControl>true</allowGuestControl><connected>false</connected><status>untried</status></connectable>
              <controllerKey>100</controllerKey><unitNumber>7</unitNumber><addressType>assigned</addressType><macAddress>00:50:56:aa:bb:02</macAddress></device>
            <device xsi:type="VirtualMachineVideoCard"><key>500</key><deviceInfo><label>Video card </label><summary>Video card</summary></deviceInfo><videoRamSizeInKB>4096</videoRamSizeInKB><numDisplays>1</numDisplays></device>
            """),
        P("config.cpuAllocation", "ResourceAllocationInfo", Allocation(0, -1, 2000, "normal")),
        P("config.memoryAllocation", "ResourceAllocationInfo", Allocation(0, -1, 163840, "normal")),
        B("config.cpuHotAddEnabled", false),
        B("config.memoryHotAddEnabled", false),
        P("config.latencySensitivity", "LatencySensitivity", "<level>high</level>"),
        B("config.changeTrackingEnabled", false),
        P("config.files", "VirtualMachineFileInfo", "<vmPathName>[ds-vmfs01] db01/db01.vmx</vmPathName><snapshotDirectory>[ds-vmfs01] db01/</snapshotDirectory><suspendDirectory>[ds-vmfs01] db01/</suspendDirectory><logDirectory>[ds-vmfs01] db01/</logDirectory>"),
        P("config.extraConfig[\"disk.EnableUUID\"]", "OptionValue", "<key>disk.EnableUUID</key><value xsi:type=\"xsd:string\">TRUE</value>"),
        P("runtime.powerState", "VirtualMachinePowerState", "poweredOff"),
        P("runtime.connectionState", "VirtualMachineConnectionState", "connected"),
        Ref("runtime.host", "HostSystem", "host-21"),
        B("runtime.consolidationNeeded", false),
        P("runtime.faultToleranceState", "VirtualMachineFaultToleranceState", "notConfigured"),
        P("guest", "GuestInfo", "<toolsStatus>toolsNotRunning</toolsStatus><toolsVersionStatus2>guestToolsSupportedOld</toolsVersionStatus2><toolsRunningStatus>guestToolsNotRunning</toolsRunningStatus><toolsVersion>11365</toolsVersion><guestState>notRunning</guestState><guestOperationsReady>false</guestOperationsReady>"),
        P("summary.storage", "VirtualMachineStorageSummary", "<committed>85899345920</committed><uncommitted>0</uncommitted><unshared>85899345920</unshared>"),
        P("summary.quickStats", "VirtualMachineQuickStats", "<overallCpuUsage>0</overallCpuUsage><guestMemoryUsage>0</guestMemoryUsage><hostMemoryUsage>0</hostMemoryUsage><balloonedMemory>0</balloonedMemory><uptimeSeconds>0</uptimeSeconds>"));

    public static string Template => O("VirtualMachine", "vm-102",
        S("name", "tmpl-ubuntu"),
        Ref("parent", "Folder", "group-v3"),
        Refs("datastore", ("Datastore", "datastore-31")),
        P("network", "ArrayOfManagedObjectReference", ""),
        P("overallStatus", "ManagedEntityStatus", "green"),
        P("configStatus", "ManagedEntityStatus", "green"),
        P("guestHeartbeatStatus", "ManagedEntityStatus", "gray"),
        B("config.template", true),
        S("config.uuid", "4231c3d4-e5f6-0718-293a-4b5c6d7e8f90"),
        S("config.instanceUuid", "5031c3d4-e5f6-0718-293a-4b5c6d7e8f90"),
        S("config.guestFullName", "Ubuntu Linux (64-bit)"),
        S("config.version", "vmx-19"),
        S("config.firmware", "efi"),
        P("config.hardware", "VirtualHardware", """
            <numCPU>1</numCPU><numCoresPerSocket>1</numCoresPerSocket><memoryMB>2048</memoryMB>
            <device xsi:type="VirtualAHCIController"><key>15000</key><deviceInfo><label>SATA controller 0</label><summary>AHCI</summary></deviceInfo><busNumber>0</busNumber><device>16000</device></device>
            <device xsi:type="VirtualDisk"><key>16000</key><deviceInfo><label>Hard disk 1</label><summary>16,777,216 KB</summary></deviceInfo>
              <backing xsi:type="VirtualDiskFlatVer2BackingInfo"><fileName>[nfs01] tmpl-ubuntu/tmpl-ubuntu.vmdk</fileName><datastore type="Datastore">datastore-31</datastore><diskMode>persistent</diskMode><thinProvisioned>true</thinProvisioned><uuid>6000C29d-4e5f-6071-8293-a4b5c6d7e8f9</uuid></backing>
              <controllerKey>15000</controllerKey><unitNumber>0</unitNumber><capacityInKB>16777216</capacityInKB><capacityInBytes>17179869184</capacityInBytes></device>
            """),
        P("config.files", "VirtualMachineFileInfo", "<vmPathName>[nfs01] tmpl-ubuntu/tmpl-ubuntu.vmtx</vmPathName>"),
        P("runtime.powerState", "VirtualMachinePowerState", "poweredOff"),
        P("runtime.connectionState", "VirtualMachineConnectionState", "connected"),
        Ref("runtime.host", "HostSystem", "host-21"),
        P("guest", "GuestInfo", "<toolsStatus>toolsNotRunning</toolsStatus><toolsRunningStatus>guestToolsNotRunning</toolsRunningStatus><guestState>notRunning</guestState>"),
        P("summary.storage", "VirtualMachineStorageSummary", "<committed>3221225472</committed><uncommitted>13958643712</uncommitted><unshared>3221225472</unshared>"));

    // ------------------------------------------------------------------ licences

    public static string Licenses() => RetrieveResult(O("LicenseManager", "LicenseManager",
        P("licenses", "ArrayOfLicenseManagerLicenseInfo", """
            <LicenseManagerLicenseInfo><licenseKey>AB12C-DE34F-GH56J-KL78M-NP90Q</licenseKey><editionKey>vc.standard.instance</editionKey><name>vCenter Server 8 Standard</name><total>1</total><used>1</used><costUnit>server</costUnit>
              <properties><key>ProductName</key><value xsi:type="xsd:string">VMware VirtualCenter Server</value></properties>
              <properties><key>ProductVersion</key><value xsi:type="xsd:string">8.0</value></properties>
              <properties><key>feature</key><value xsi:type="KeyValue"><key>vpxStd</key><value>vCenter Server Standard</value></value></properties>
              <labels><key>Owner</key><value>IT</value></labels></LicenseManagerLicenseInfo>
            <LicenseManagerLicenseInfo><licenseKey>ZX98Y-WV76U-TS54R-QP32O-NM10L</licenseKey><editionKey>esx.enterprisePlus.cpuPackageCoreLimited</editionKey><name>vSphere 8 Enterprise Plus</name><total>64</total><used>48</used><costUnit>cpuPackage:32core</costUnit>
              <properties><key>expirationDate</key><value xsi:type="xsd:dateTime">2027-12-31T00:00:00Z</value></properties>
              <properties><key>feature</key><value xsi:type="KeyValue"><key>vmotion</key><value>vMotion</value></value></properties>
              <properties><key>feature</key><value xsi:type="KeyValue"><key>drs</key><value>vSphere DRS</value></value></properties></LicenseManagerLicenseInfo>
            """),
        Ref("licenseAssignmentManager", "LicenseAssignmentManager", "LicenseAssignmentManager")));

    public static string AssignedLicenses() => Envelope("""
        <QueryAssignedLicensesResponse xmlns="urn:vim25">
          <returnval><entityId>host-20</entityId><scope>7d1e2a3b-4c5d-4e6f-8a9b-0c1d2e3f4a5b</scope><entityDisplayName>esx01.lab.local</entityDisplayName><assignedLicense><licenseKey>ZX98Y-WV76U-TS54R-QP32O-NM10L</licenseKey><editionKey>esx.enterprisePlus.cpuPackageCoreLimited</editionKey><name>vSphere 8 Enterprise Plus</name><total>64</total><used>48</used><costUnit>cpuPackage:32core</costUnit></assignedLicense></returnval>
          <returnval><entityId>7d1e2a3b-4c5d-4e6f-8a9b-0c1d2e3f4a5b</entityId><entityDisplayName>vc01.lab.local</entityDisplayName><assignedLicense><licenseKey>AB12C-DE34F-GH56J-KL78M-NP90Q</licenseKey><name>vCenter Server 8 Standard</name><total>1</total><used>1</used><costUnit>server</costUnit></assignedLicense></returnval>
        </QueryAssignedLicensesResponse>
        """);
}
