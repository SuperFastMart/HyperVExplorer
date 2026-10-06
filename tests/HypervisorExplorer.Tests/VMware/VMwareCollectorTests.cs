using System.Net;
using HypervisorExplorer.Collectors.VMware;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Tests.VMware;

public class VMwareCollectorTests
{
    private static ConnectionRequest Request(string password = "secret", string? group = null) => new()
    {
        Platform = Platform.VMware,
        Address = "vc01.lab.local",
        Username = "administrator@vsphere.local",
        Secret = password,
        Group = group,
    };

    private static async Task<(InventorySnapshot Snapshot, FakeVSphereHandler Handler, List<string> Progress)> CollectAsync(
        Action<FakeVSphereHandler>? setup = null)
    {
        var handler = new FakeVSphereHandler();
        setup?.Invoke(handler);
        var progress = new List<string>();
        var collector = new VMwareCollector(() => handler);
        var snap = await collector.CollectAsync(Request(), new SyncProgress(progress), CancellationToken.None);
        return (snap, handler, progress);
    }

    private sealed class SyncProgress(List<string> sink) : IProgress<string>
    {
        public void Report(string value) => sink.Add(value);
    }

    private static object? Cell(TableDefinition table, object row, string header) =>
        table.Columns.Single(c => c.Header == header).GetValue(row);

    private static object RowFor(TableDefinition table, Inventory inv, string firstHeader, string value) =>
        table.Rows(inv).Single(r => Equals(Cell(table, r, firstHeader), value));

    // ------------------------------------------------------------------ protocol

    [Fact]
    public async Task Collect_FollowsSoapFlow_PagesAndAlwaysLogsOut()
    {
        var (snap, h, progress) = await CollectAsync();

        Assert.Equal("RetrieveServiceContent", h.Calls[0].Operation);
        Assert.Equal("Login", h.Calls[1].Operation);
        Assert.Equal("Logout", h.Calls[^1].Operation);
        Assert.Equal(1, h.Count("Logout"));
        Assert.Equal(1, h.Count("ContinueRetrievePropertiesEx"));
        Assert.Equal(h.Count("CreateContainerView"), h.Count("DestroyView"));
        Assert.Equal(8, h.Count("CreateContainerView"));
        Assert.Equal(new Uri("https://vc01.lab.local/sdk"), h.Calls[0].Uri);

        // SOAPAction switches to the detected API version; session cookie is replayed after Login.
        Assert.Equal("urn:vim25", h.Calls[0].SoapAction);
        Assert.All(h.Calls.Skip(1), c => Assert.Equal("urn:vim25/8.0.2.0", c.SoapAction));
        Assert.All(h.Calls.Skip(2), c => Assert.Contains("vmware_soap_session", c.Cookie));

        // Property specs traverse the container view and page with maxObjects.
        var vmCall = h.Calls.First(c => c.Operation == "RetrievePropertiesEx" && c.Body.Contains("<type>VirtualMachine</type>"));
        Assert.Contains("<selectSet xsi:type=\"TraversalSpec\"><name>traverseView</name><type>ContainerView</type><path>view</path>", vmCall.Body);
        Assert.Contains("<maxObjects>500</maxObjects>", vmCall.Body);
        Assert.Contains("<pathSet>config.hardware</pathSet>", vmCall.Body);
        Assert.Contains("<token>1</token>", h.Calls.Single(c => c.Operation == "ContinueRetrievePropertiesEx").Body);

        Assert.Equal(3, snap.VirtualMachines.Count);
        Assert.Equal(2, snap.Hosts.Count);
        Assert.Equal(2, snap.Datastores.Count);
        Assert.Single(snap.Clusters);
        Assert.Contains(progress, p => p.Contains("Retrieved 3 VMs"));
        Assert.Empty(snap.Warnings);
    }

    [Fact]
    public async Task InvalidLogin_MapsToAuthenticationFailed_WithoutLogout()
    {
        var h = new FakeVSphereHandler();
        var collector = new VMwareCollector(() => h);
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            collector.CollectAsync(Request("wrong"), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.AuthenticationFailed, ex.Kind);
        Assert.Contains("incorrect user name or password", ex.Message);
        Assert.Contains("administrator@vsphere.local", ex.Hint);
        Assert.Equal(0, h.Count("Logout"));
    }

    [Fact]
    public async Task FaultAfterLogin_StillLogsOut_AndMapsNoPermission()
    {
        var h = new FakeVSphereHandler
        {
            Override = (op, type, _) => op == "RetrievePropertiesEx" && type == "HostSystem"
                ? FakeVSphereHandler.ServerFault("NoPermission", "Permission to perform this operation was denied.",
                    "<object type=\"Folder\">group-d1</object><privilegeId>System.View</privilegeId>")
                : null,
        };
        var collector = new VMwareCollector(() => h);
        var ex = await Assert.ThrowsAsync<CollectionException>(() => collector.CollectAsync(Request(), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.PermissionDenied, ex.Kind);
        Assert.Contains("System.View", ex.Message);
        Assert.Equal(1, h.Count("Logout"));
        Assert.Equal(h.Count("CreateContainerView"), h.Count("DestroyView"));
    }

    [Fact]
    public async Task InvalidProperty_IsDroppedAndRetried()
    {
        var rejected = 0;
        var (snap, h, _) = await CollectAsync(handler => handler.Override = (op, type, body) =>
            op == "RetrievePropertiesEx" && type == "VirtualMachine" && body.Contains("config.createDate") && rejected++ == 0
                ? FakeVSphereHandler.ServerFault("InvalidProperty", "", "<name>config.createDate</name>")
                : null);
        Assert.Equal(1, rejected);
        Assert.Equal(3, snap.VirtualMachines.Count);
        Assert.Contains(snap.Warnings, w => w.Contains("config.createDate"));
        var retry = h.Calls.Where(c => c.Operation == "RetrievePropertiesEx" && c.Body.Contains("<type>VirtualMachine</type>")).Last();
        Assert.DoesNotContain("config.createDate", retry.Body);
    }

    [Fact]
    public async Task Unreachable_MapsToUnreachableWithPortHint()
    {
        var collector = new VMwareCollector(() => new ThrowingHandler());
        var ex = await Assert.ThrowsAsync<CollectionException>(() => collector.CollectAsync(Request(), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.Unreachable, ex.Kind);
        Assert.Contains("443", ex.Hint);
    }

    private sealed class ThrowingHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            throw new HttpRequestException("Connection refused", new System.Net.Sockets.SocketException(61));
    }

    [Fact]
    public async Task NonSoapResponse_MapsToProtocol()
    {
        var h = new FakeVSphereHandler { Override = (_, _, _) => (HttpStatusCode.NotFound, "<html><body>Not found</body>") };
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new VMwareCollector(() => h).CollectAsync(Request(), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.Protocol, ex.Kind);
    }

    [Theory]
    [InlineData("AB12C-DE34F-GH56J-KL78M-NP90Q", "XXXXX-XXXXX-XXXXX-XXXXX-NP90Q")]
    [InlineData("ABCDEF", "XBCDEF")]
    [InlineData("ABC", "ABC")]
    public void LicenseKey_IsMaskedExceptLastFive(string key, string expected) =>
        Assert.Equal(expected, VMwareInventoryMapper.MaskLicenseKey(key));

    [Fact]
    public void Wwn_IsFormattedAsHexPairs() =>
        Assert.Equal("21:00:00:24:ff:7e:5a:01", VMwareInventoryMapper.FormatWwn(0x210000_24ff7e5a01));

    // ------------------------------------------------------------------ model mapping

    [Fact]
    public async Task Maps_Source_Hosts_Clusters_Datastores()
    {
        var (snap, _, _) = await CollectAsync();

        var src = snap.Source;
        Assert.Equal(Platform.VMware, src.Platform);
        Assert.Equal("VMware vCenter Server", src.ProductName);
        Assert.Equal("8.0.2.0", src.ApiVersion);
        Assert.Equal("22617221", src.Build);
        Assert.Equal("VirtualCenter", src.Extra["API type"]);
        Assert.Equal(VSphereFixtures.InstanceUuid, src.Extra["VI SDK UUID"]);

        var esx01 = snap.Hosts.Single(x => x.Name == "esx01.lab.local");
        Assert.Equal("DC1", esx01.Datacenter);
        Assert.Equal("Cluster01", esx01.Cluster);
        Assert.Equal(2, esx01.CpuSockets);
        Assert.Equal(24, esx01.CoresPerSocket);
        Assert.Equal(10, Math.Round(esx01.CpuUsagePercent!.Value));
        Assert.Equal("VMware ESXi 8.0.2 build-22380479", esx01.Version);
        Assert.Equal("5JK8XG3", esx01.SerialNumber);
        Assert.Equal(["10.0.0.10", "10.0.0.11"], esx01.DnsServers);
        Assert.Equal(["0.pool.ntp.org", "1.pool.ntp.org"], esx01.NtpServers);
        Assert.Equal(3, esx01.Nics.Count);
        Assert.Equal("vSwitch0", esx01.Nics.Single(n => n.Name == "vmnic0").Switch);
        Assert.Equal("DSwitch01", esx01.Nics.Single(n => n.Name == "vmnic1").Switch);
        Assert.Equal("Uplink 1", esx01.Nics.Single(n => n.Name == "vmnic1").Extra["Uplink port"]);
        Assert.Equal("Down", esx01.Nics.Single(n => n.Name == "vmnic2").Status);
        Assert.Equal(["vmnic0"], esx01.Switches.Single().Uplinks);
        Assert.Equal(10, esx01.PortGroups.Single(p => p.Name == "Management Network").Vlan);
        Assert.Equal("PG-App-100", esx01.IpInterfaces.Single(i => i.Name == "vmk1").PortGroup);
        Assert.Equal("10.0.10.1", esx01.IpInterfaces.Single(i => i.Name == "vmk0").Gateway);
        var fc = esx01.StorageAdapters.Single(a => a.Device == "vmhba2");
        Assert.Equal("Fibre Channel", fc.Type);
        Assert.Equal("20:00:00:24:ff:7e:5a:01 21:00:00:24:ff:7e:5a:01", fc.Wwn);
        Assert.Equal("vSphere 8 Enterprise Plus", esx01.Extra["Assigned License(s)"]);
        Assert.Equal(true, esx01.Extra["NTPD running"]);
        Assert.Equal("Balanced", esx01.Extra["Host Power Policy"]);

        var esx02 = snap.Hosts.Single(x => x.Name == "esx02.lab.local");
        Assert.True(esx02.InMaintenance);

        var cl = snap.Clusters.Single();
        Assert.Equal("Cluster01", cl.Name);
        Assert.Equal("DC1", cl.Datacenter);
        Assert.True(cl.HaEnabled);
        Assert.True(cl.DrsEnabled);
        Assert.Equal(["esx01.lab.local", "esx02.lab.local"], cl.Nodes.Select(n => n.Name));
        Assert.Equal("inMaintenance", cl.Nodes[1].DrainStatus);

        var vmfs = snap.Datastores.Single(d => d.Name == "ds-vmfs01");
        Assert.Equal("VMFS", vmfs.Type);
        Assert.True(vmfs.Shared);
        Assert.Equal(["esx01.lab.local", "esx02.lab.local"], vmfs.Hosts);
        Assert.Equal("Cluster01", vmfs.Cluster);
        Assert.Equal("naa.600a098038304437415d4b6a59684a52", vmfs.Address);
        var nfs = snap.Datastores.Single(d => d.Name == "nfs01");
        Assert.Equal("nas01.lab.local:/vol/vmware", nfs.Address);
    }

    [Fact]
    public async Task Maps_VirtualMachines()
    {
        var (snap, _, _) = await CollectAsync();

        var web = snap.VirtualMachines.Single(v => v.Name == "web01");
        Assert.Equal("vm-100", web.VmId);
        Assert.Equal(PowerState.PoweredOn, web.PowerState);
        Assert.Equal("DC1", web.Datacenter);
        Assert.Equal("Cluster01", web.Cluster);
        Assert.Equal("esx01.lab.local", web.Host);
        Assert.Equal("/DC1/vm/Linux", web.ResourcePool);
        Assert.Equal("/DC1/host/Cluster01/Resources/Prod", web.Extra["Resource pool"]);
        Assert.Equal("Web front end & API", web.Annotation);
        Assert.Equal(4, web.CpuCount);
        Assert.Equal(2, web.Sockets);
        Assert.Equal(8192, web.MemoryMiB);
        Assert.Equal(8250, web.MemoryAssignedMiB);
        Assert.Equal(1638, web.MemoryDemandMiB);
        Assert.Equal(256, web.MemoryBalloonedMiB);
        Assert.True(web.SecureBoot);
        Assert.Equal("vmx-19", web.HardwareVersion);
        Assert.Equal("toolsOk", web.ToolsStatus);
        Assert.True(web.ToolsRunning);
        Assert.Equal("green", web.Heartbeat);
        Assert.Equal("high", web.Extra["HA Restart Priority"]);
        Assert.Equal("none", web.Extra["HA Isolation Response"]);
        Assert.Equal("separate-web-db", web.Extra["Cluster rule name(s)"]);
        Assert.Equal("TOOLSDEPLOYPKG_IDLE", web.Extra["Customization Info"]);

        Assert.Equal(2, web.Disks.Count);
        var d1 = web.Disks[0];
        Assert.Equal("Hard disk 1", d1.Label);
        Assert.Equal("SCSI controller 0", d1.Controller);
        Assert.Equal("VMware Paravirtual", d1.ControllerType);
        Assert.Equal("ds-vmfs01", d1.Datastore);
        Assert.True(d1.Thin);
        Assert.Equal(53687091200, d1.CapacityBytes);
        Assert.Equal("[ds-vmfs01] web01/web01-000001.vmdk", d1.ParentPath);
        Assert.Equal(600 + 10737418240 + 400 + 1073741824 + 400 + 536870912, d1.UsedBytes);
        var d2 = web.Disks[1];
        Assert.Equal("nfs01", d2.Datastore);
        Assert.False(d2.Thin);
        Assert.True(d2.Shared);
        Assert.Equal("independent_persistent", d2.Format);
        Assert.Equal(true, d2.Extra["Eagerly Scrub"]);

        Assert.Equal(2, web.Nics.Count);
        var n1 = web.Nics[0];
        Assert.Equal("Network adapter 1", n1.Label);
        Assert.Equal("vmxnet3", n1.AdapterType);
        Assert.Equal("PG-App-100", n1.Network);
        Assert.Equal("DSwitch01", n1.Switch);
        Assert.Equal(["10.0.100.11"], n1.Ipv4);
        Assert.Equal(["fe80::250:56ff:feaa:bb01"], n1.Ipv6);
        var n2 = web.Nics[1];
        Assert.Equal("e1000e", n2.AdapterType);
        Assert.Equal("VM Network", n2.Network);
        Assert.Equal("vSwitch0", n2.Switch);
        Assert.False(n2.Connected);
        Assert.Equal("10.0.100.11", web.PrimaryIp);

        var cd = Assert.Single(web.CdDrives);
        Assert.Equal("CD/DVD drive 1", cd.DeviceNode);
        Assert.Equal("[ds-vmfs01] iso/ubuntu-22.04.iso", cd.Media);
        var usb = Assert.Single(web.UsbDevices);
        Assert.Equal(true, usb.Extra["EHCI enabled"]);

        Assert.Equal(2, web.Snapshots.Count);
        Assert.Equal("Before patch", web.Snapshots[0].Name);
        Assert.Null(web.Snapshots[0].Parent);
        Assert.Equal("After patch", web.Snapshots[1].Name);
        Assert.Equal("Before patch", web.Snapshots[1].Parent);
        Assert.True(web.Snapshots[1].IncludesMemory);
        Assert.Equal("[ds-vmfs01] web01/web01-Snapshot1.vmsn", web.Snapshots[0].Path);
        Assert.Equal(2, web.Partitions.Count);

        var db = snap.VirtualMachines.Single(v => v.Name == "db01");
        Assert.Equal(PowerState.PoweredOff, db.PowerState);
        Assert.Equal("/DC1/vm", db.ResourcePool);
        Assert.Equal(true, db.Extra["EnableUUID"]);
        Assert.Equal("high", db.Extra["Latency Sensitivity"]);
        Assert.Equal("LSI Logic SAS", db.Disks.Single().ControllerType);
        Assert.False(db.Disks.Single().Thin);
        Assert.Equal("medium", db.Extra["HA Restart Priority"]);

        var tmpl = snap.VirtualMachines.Single(v => v.Name == "tmpl-ubuntu");
        Assert.True(tmpl.IsTemplate);
        Assert.Equal("SATA controller 0", tmpl.Disks.Single().Controller);
        Assert.Null(tmpl.Extra["Resource pool"]);
    }

    [Fact]
    public async Task Health_AddsVmwareSpecificItems()
    {
        var (snap, _, _) = await CollectAsync();
        Assert.Contains(snap.Health, x => x.Name == "web01" && x.Message.Contains("consolidation"));
        Assert.Contains(snap.Health, x => x.Name == "nfs01" && x.Severity == HealthSeverity.Warning);
        Assert.Contains(snap.Health, x => x.Name == "esx02.lab.local" && x.Message.Contains("overall status is yellow"));
        Assert.DoesNotContain(snap.Health, x => x.Message.Contains("maintenance mode"));
    }

    [Fact]
    public async Task StandaloneEsxi_PrefersGroupOverHaDatacenter()
    {
        var h = new FakeVSphereHandler
        {
            Override = (op, type, _) => op == "RetrievePropertiesEx" && type == "ManagedEntity"
                ? ((HttpStatusCode, string)?)(HttpStatusCode.OK, VSphereFixtures.Entities().Replace(">DC1<", ">ha-datacenter<"))
                : null,
        };
        var snap = await new VMwareCollector(() => h).CollectAsync(Request(group: "London"), null, CancellationToken.None);
        Assert.All(snap.VirtualMachines, v => Assert.Equal("London", v.Datacenter));
        Assert.Equal("/ha-datacenter/vm/Linux", snap.VirtualMachines.Single(v => v.Name == "web01").ResourcePool);
    }

    // ------------------------------------------------------------------ extra sheets

    [Fact]
    public async Task SheetRows_ForPoolsSwitchesMultipathAndLicenses()
    {
        var (snap, _, _) = await CollectAsync();
        List<Dictionary<string, object?>> Sheet(string name) =>
            Assert.IsType<List<Dictionary<string, object?>>>(snap.Source.Extra[VMwareInventoryMapper.SheetPrefix + name]);

        var rp = Sheet("vRP");
        Assert.Equal(2, rp.Count);
        var prod = rp.Single(r => (string?)r["Resource Pool name"] == "Prod");
        Assert.Equal("/DC1/host/Cluster01/Resources/Prod", prod["Resource Pool path"]);
        Assert.Equal(1, prod["# VMs total"]);
        Assert.Equal(4, prod["# vCPUs"]);
        Assert.Equal(32768L, prod["Mem limit"]);
        Assert.Equal(true, prod["CPU expandableReservation"]);

        var dvs = Assert.Single(Sheet("dvSwitch"));
        Assert.Equal("DSwitch01", dvs["Switch"]);
        Assert.Equal("DC1", dvs["Datacenter"]);
        Assert.Equal("8.0.0", dvs["Version"]);
        Assert.Equal("esx01.lab.local, esx02.lab.local", dvs["Host members"]);
        Assert.Equal(9000, dvs["Max MTU"]);
        Assert.Equal("lag1", dvs["LACP Name"]);
        Assert.Equal("cdp", dvs["CDP Type"]);

        var ports = Sheet("dvPort");
        var app = ports.Single(p => (string?)p["Port"] == "PG-App-100");
        Assert.Equal("DSwitch01", app["Switch"]);
        Assert.Equal("100", app["VLAN"]);
        Assert.Equal("Uplink 1", app["Active Uplink"]);
        Assert.Equal("loadbalance_srcid", app["Policy"]);
        Assert.Equal(true, app["Block Override"]);
        Assert.Equal("0-4094", ports.Single(p => (string?)p["Port"] == "DSwitch01-DVUplinks-60")["VLAN"]);

        var mp = Assert.Single(Sheet("vMultiPath"));
        Assert.Equal("esx01.lab.local", mp["Host"]);
        Assert.Equal("ds-vmfs01", mp["Datastore"]);
        Assert.Equal("VMW_PSP_RR", mp["Policy"]);
        Assert.Equal("vmhba2:C0:T0:L1", mp["Path 1"]);
        Assert.Equal("standby", mp["Path 2 state"]);
        Assert.Equal("NETAPP", mp["Vendor"]);

        var lic = Sheet("vLicense");
        Assert.Equal(2, lic.Count);
        Assert.Equal("XXXXX-XXXXX-XXXXX-XXXXX-NP90Q", lic[0]["Key"]);
        Assert.Equal("XXXXX-XXXXX-XXXXX-XXXXX-NM10L", lic[1]["Key"]);
        Assert.Equal("vMotion, vSphere DRS", lic[1]["Features"]);
        Assert.Equal(48L, lic[1]["Used"]);
        Assert.Equal("Never", lic[0]["Expiration Date"]);
        Assert.All(lic, r => Assert.Equal(VSphereFixtures.InstanceUuid, r["VI SDK UUID"]));
        Assert.DoesNotContain(snap.Source.Extra.Values.OfType<List<Dictionary<string, object?>>>()
            .SelectMany(rows => rows).SelectMany(r => r.Values).OfType<string>(), s => s.Contains("AB12C"));
    }

    // ------------------------------------------------------------------ RVTools tables

    [Fact]
    public async Task RvToolsTables_ShowVmwareValues()
    {
        var (snap, _, _) = await CollectAsync();
        var store = new InventoryStore();
        store.Upsert(snap);
        var inv = store.Current;

        var vInfo = RvToolsTables.Get("vInfo");
        var web = RowFor(vInfo, inv, "VM", "web01");
        Assert.Equal(19, Cell(vInfo, web, "HW version"));
        Assert.Equal("green", Cell(vInfo, web, "Config status"));
        Assert.Equal("/DC1/vm/Linux", Cell(vInfo, web, "Folder"));
        Assert.Equal("group-v100", Cell(vInfo, web, "Folder ID"));
        Assert.Equal("PG-App-100", Cell(vInfo, web, "Network #1"));
        Assert.Equal("VM Network", Cell(vInfo, web, "Network #2"));
        Assert.Equal("connected", Cell(vInfo, web, "Connection state"));
        Assert.Equal("running", Cell(vInfo, web, "Guest state"));
        Assert.Equal("True", Cell(vInfo, web, "Consolidation Needed"));
        Assert.Equal("True", Cell(vInfo, web, "CBT"));
        Assert.Equal(2, Cell(vInfo, web, "Num Monitors"));
        Assert.Equal(8192L, Cell(vInfo, web, "Video Ram KiB"));
        Assert.Equal(5000L, Cell(vInfo, web, "Boot delay"));
        Assert.Equal("intel-sandybridge", Cell(vInfo, web, "min Required EVC Mode Key"));
        Assert.Equal(154112L, Cell(vInfo, web, "Provisioned MiB"));
        Assert.Equal(31232L, Cell(vInfo, web, "In Use MiB"));
        Assert.Equal("DC1", Cell(vInfo, web, "Datacenter"));
        Assert.Equal("VMware vCenter Server", Cell(vInfo, web, "VI SDK Server type"));
        Assert.Equal(VSphereFixtures.InstanceUuid, Cell(vInfo, web, "VI SDK UUID"));
        Assert.Equal("vm-100", Cell(vInfo, web, "VM ID"));
        Assert.Equal("4231a1b2-c3d4-e5f6-0718-293a4b5c6d7e", Cell(vInfo, web, "SMBIOS UUID"));
        Assert.Equal("True", Cell(vInfo, RowFor(vInfo, inv, "VM", "tmpl-ubuntu"), "Template"));

        var vCpu = RvToolsTables.Get("vCPU");
        var webCpu = RowFor(vCpu, inv, "VM", "web01");
        Assert.Equal(4000, Cell(vCpu, webCpu, "Shares"));
        Assert.Equal(840L, Cell(vCpu, webCpu, "Overall"));

        var vDisk = RvToolsTables.Get("vDisk");
        var disks = vDisk.Rows(inv).Where(r => Equals(Cell(vDisk, r, "VM"), "web01")).ToList();
        Assert.Equal("True", Cell(vDisk, disks[0], "Thin"));
        Assert.Equal("False", Cell(vDisk, disks[1], "Thin"));
        Assert.Equal("sharingMultiWriter", Cell(vDisk, disks[1], "Sharing mode"));
        Assert.Equal("persistent", Cell(vDisk, disks[0], "Disk Mode"));
        Assert.Equal(51200L, Cell(vDisk, disks[0], "Capacity MiB"));
        Assert.Equal("SCSI controller 0", Cell(vDisk, disks[0], "Controller"));
        Assert.Equal(1, Cell(vDisk, disks[1], "SCSI Unit #"));
        Assert.Equal(2000, Cell(vDisk, disks[0], "Disk Key"));
        Assert.Equal("6000C29a-1b2c-3d4e-5f60-718293a4b5c6", Cell(vDisk, disks[0], "Disk UUID"));

        var vNetwork = RvToolsTables.Get("vNetwork");
        var nics = vNetwork.Rows(inv).Where(r => Equals(Cell(vNetwork, r, "VM"), "web01")).ToList();
        Assert.Equal("DSwitch01", Cell(vNetwork, nics[0], "Switch"));
        Assert.Equal("vmxnet3", Cell(vNetwork, nics[0], "Adapter"));
        Assert.Equal("assigned", Cell(vNetwork, nics[0], "Type"));
        Assert.Equal("10.0.100.11", Cell(vNetwork, nics[0], "IPv4 Address"));
        Assert.Equal("vSwitch0", Cell(vNetwork, nics[1], "Switch"));

        var vSnap = RvToolsTables.Get("vSnapshot");
        var snaps = vSnap.Rows(inv).ToList();
        Assert.Equal(2, snaps.Count);
        Assert.Equal(1026L, Cell(vSnap, snaps[0], "Size MiB (total)"));
        Assert.Equal(2L, Cell(vSnap, snaps[0], "Size MiB (vmsn)"));
        Assert.Equal("poweredOn", Cell(vSnap, snaps[1], "State"));
        Assert.Equal("True", Cell(vSnap, snaps[1], "Quiesced"));

        var vHost = RvToolsTables.Get("vHost");
        var esx01 = RowFor(vHost, inv, "Host", "esx01.lab.local");
        Assert.Equal("intel-icelake", Cell(vHost, esx01, "Current EVC"));
        Assert.Equal("DC1", Cell(vHost, esx01, "Datacenter"));
        Assert.Equal("Cluster01", Cell(vHost, esx01, "Cluster"));
        Assert.Equal(1, Cell(vHost, esx01, "# VMs total"));
        Assert.Equal("5JK8XG3", Cell(vHost, esx01, "Service tag"));
        Assert.Equal("True", Cell(vHost, esx01, "HT Available"));
        Assert.Equal("1.10.2", Cell(vHost, esx01, "BIOS Version"));

        var vCluster = RvToolsTables.Get("vCluster");
        var cl = vCluster.Rows(inv).Single();
        Assert.Equal("True", Cell(vCluster, cl, "HA enabled"));
        Assert.Equal("True", Cell(vCluster, cl, "DRS enabled"));
        Assert.Equal("fullyAutomated", Cell(vCluster, cl, "DRS default VM behavior"));
        Assert.Equal("medium", Cell(vCluster, cl, "Restart Priority"));
        Assert.Equal(42, Cell(vCluster, cl, "Num VMotions"));
        Assert.Equal("domain-c10", Cell(vCluster, cl, "Object ID"));

        var vDatastore = RvToolsTables.Get("vDatastore");
        var ds = RowFor(vDatastore, inv, "Name", "ds-vmfs01");
        Assert.Equal(1, Cell(vDatastore, ds, "Block size"));
        Assert.Equal("True", Cell(vDatastore, ds, "SIOC enabled"));
        Assert.Equal("True", Cell(vDatastore, ds, "MHA"));
        Assert.Equal(2, Cell(vDatastore, ds, "# VMs total"));
        Assert.Equal("6.82", Cell(vDatastore, ds, "Version"));

        var vSwitch = RvToolsTables.Get("vSwitch");
        var sw = vSwitch.Rows(inv).First(r => Equals(Cell(vSwitch, r, "Host"), "esx01.lab.local"));
        Assert.Equal("loadbalance_srcid", Cell(vSwitch, sw, "Policy"));
        Assert.Equal(2546, Cell(vSwitch, sw, "Free Ports"));

        var vSource = RvToolsTables.Get("vSource");
        var src = vSource.Rows(inv).Single();
        Assert.Equal("VirtualCenter", Cell(vSource, src, "API type"));
        Assert.Equal("VMware vCenter Server 8.0.2 build-22617221", Cell(vSource, src, "Fullname"));
        Assert.Equal("vpx", Cell(vSource, src, "Product line"));

        var vTools = RvToolsTables.Get("vTools");
        var dbTools = RowFor(vTools, inv, "VM", "db01");
        Assert.Equal("True", Cell(vTools, dbTools, "Upgradeable"));
        Assert.Equal(15, Cell(vTools, dbTools, "VM Version"));
    }
}
