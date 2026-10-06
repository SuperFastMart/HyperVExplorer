using System.Net;
using System.Net.Sockets;
using System.Security.Authentication;
using HypervisorExplorer.Collectors.Proxmox;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Tests.Proxmox;

public class ProxmoxCollectorTests
{
    private static readonly DateTimeOffset Now = new(2024, 6, 4, 12, 0, 0, TimeSpan.Zero);

    private static ConnectionRequest TokenRequest(string address = "pve1.lab.local") => new()
    {
        Platform = Platform.Proxmox,
        Address = address,
        CredentialKind = CredentialKind.ApiToken,
        Username = "root@pam!explorer",
        Secret = "11111111-2222-3333-4444-555555555555",
        Group = "London",
    };

    private static ConnectionRequest PasswordRequest() => new()
    {
        Platform = Platform.Proxmox,
        Address = "pve1.lab.local",
        CredentialKind = CredentialKind.UsernamePassword,
        Username = "root",
        Secret = "p@ss word",
    };

    private static ProxmoxCollector Collector(FakePveHandler h) => new(h)
    {
        Clock = () => Now,
        AgentTimeout = TimeSpan.FromMilliseconds(300),
        RequestTimeout = TimeSpan.FromSeconds(5),
    };

    private static Task<InventorySnapshot> Collect(FakePveHandler h, ConnectionRequest? req = null) =>
        Collector(h).CollectAsync(req ?? TokenRequest(), null, CancellationToken.None);

    [Fact]
    public async Task Token_AuthHeaderFormat_SentOnEveryRequest()
    {
        var h = PveFixtures.Cluster();
        await Collect(h);

        Assert.NotEmpty(h.Requests);
        Assert.All(h.Requests, r =>
            Assert.Equal("PVEAPIToken=root@pam!explorer=11111111-2222-3333-4444-555555555555", r.Authorization));
        Assert.DoesNotContain(h.Requests, r => r.Path == "/access/ticket");
        Assert.All(h.Requests, r => Assert.Equal("GET", r.Method));
    }

    [Fact]
    public async Task Password_TicketFlow_AppendsPamAndSendsCookie()
    {
        var h = PveFixtures.Cluster();
        var snap = await Collect(h, PasswordRequest());

        var first = h.Requests.First();
        Assert.Equal("POST", first.Method);
        Assert.Equal("/access/ticket", first.Path);
        Assert.Contains("username=root%40pam", first.Body);
        Assert.Contains("password=p%40ss+word", first.Body);

        var gets = h.Requests.Skip(1).ToList();
        Assert.NotEmpty(gets);
        Assert.All(gets, r =>
        {
            Assert.Equal("PVEAuthCookie=PVE:root@pam:66A1B2C3::c2lnbmF0dXJl+/=", r.Cookie);
            Assert.Null(r.Authorization);
            Assert.Null(r.Csrf); // CSRF token only accompanies write requests
        });
        Assert.Equal(3, snap.VirtualMachines.Count);
    }

    [Fact]
    public async Task Collects_SourceClusterAndHosts()
    {
        var snap = await Collect(PveFixtures.Cluster());

        Assert.Equal("Proxmox VE", snap.Source.ProductName);
        Assert.Equal("8.2.4", snap.Source.Version);
        Assert.Equal("8.2/faa83925c9641325", snap.Source.ApiVersion);
        Assert.Equal("Proxmox Server Solutions GmbH", snap.Source.Vendor);
        Assert.Equal("Linux", snap.Source.OsType);
        Assert.Equal("London", snap.Source.Group);

        var cluster = Assert.Single(snap.Clusters);
        Assert.Equal("lab-cluster", cluster.Name);
        Assert.True(cluster.Quorate);
        Assert.Equal(2, cluster.NumHosts);
        Assert.Equal(2, cluster.NumEffectiveHosts);
        Assert.Equal(40, cluster.NumCpuCores);
        Assert.Equal(80, cluster.NumCpuThreads);
        Assert.True(cluster.HaEnabled);
        Assert.Equal(["pve1", "pve2"], cluster.Nodes.Select(n => n.Name));
        Assert.Equal("10.0.0.12", cluster.Nodes[1].Address);
        Assert.Equal("2", cluster.Nodes[1].Id);
        Assert.All(cluster.Nodes, n => Assert.Equal("online", n.State));

        var pve1 = snap.Hosts.Single(h => h.Name == "pve1");
        Assert.Equal("lab-cluster", pve1.Cluster);
        Assert.Equal("green", pve1.Status);
        Assert.Equal("Intel(R) Xeon(R) Silver 4314 CPU @ 2.40GHz", pve1.CpuModel);
        Assert.Equal(2400, pve1.CpuMhz);
        Assert.Equal(2, pve1.CpuSockets);
        Assert.Equal(16, pve1.CoresPerSocket);
        Assert.Equal(32, pve1.CpuCores);
        Assert.Equal(64, pve1.CpuThreads);
        Assert.True(pve1.HyperThreadingActive);
        Assert.Equal(270000000000, pve1.MemoryBytes);
        Assert.Equal(54000000000, pve1.MemoryUsedBytes);
        Assert.Equal(5.3, pve1.CpuUsagePercent);
        Assert.Equal("Proxmox VE 8.2.4", pve1.Version);
        Assert.Equal("6.8.8-2-pve", pve1.KernelVersion);
        Assert.Equal(Now - TimeSpan.FromDays(14), pve1.BootTime);
        Assert.Equal(["10.0.0.2", "1.1.1.1"], pve1.DnsServers);
        Assert.Equal("lab.local", pve1.Domain);
        Assert.Equal("Europe/London", pve1.TimeZone);
        Assert.Null(pve1.Vendor); // DMI data is not exposed by the API

        // Network: physical NICs via bond into a VLAN-aware bridge, a VLAN interface, two IP interfaces.
        Assert.Equal(["eno1", "eno2"], pve1.Nics.Select(n => n.Name));
        Assert.All(pve1.Nics, n => Assert.Equal("vmbr0", n.Switch));
        Assert.Contains("bond0", pve1.Nics[0].Description);
        var vmbr0 = Assert.Single(pve1.Switches);
        Assert.Equal("vmbr0", vmbr0.Name);
        Assert.Equal(["bond0"], vmbr0.Uplinks);
        Assert.Equal(9000, vmbr0.Mtu);
        Assert.Equal("Linux bridge (VLAN-aware)", vmbr0.Type);
        Assert.Contains("eno1, eno2", vmbr0.Notes);
        var vlan = Assert.Single(pve1.PortGroups);
        Assert.Equal(("vmbr0.20", "vmbr0", 20), (vlan.Name, vlan.Switch, vlan.Vlan));
        Assert.Equal(2, pve1.IpInterfaces.Count);
        var mgmt = pve1.IpInterfaces.Single(i => i.Name == "vmbr0");
        Assert.Equal(("10.0.0.11", "255.255.255.0", "10.0.0.1"), (mgmt.Ipv4, mgmt.SubnetMask, mgmt.Gateway));
        Assert.False(mgmt.Dhcp);

        // hardware/pci: storage controllers only.
        Assert.Equal(["SATA (AHCI)", "NVMe"], pve1.StorageAdapters.Select(a => a.Type));

        var pve2 = snap.Hosts.Single(h => h.Name == "pve2");
        Assert.Empty(pve2.StorageAdapters); // 403 on hardware/pci is tolerated
        Assert.Equal("255.255.255.0", pve2.IpInterfaces.Single().SubnetMask);
        Assert.Equal("vmbr0", Assert.Single(pve2.Nics).Switch);
    }

    [Fact]
    public async Task Collects_QemuVm_WithAgent()
    {
        var snap = await Collect(PveFixtures.Cluster());
        var vm = snap.VirtualMachines.Single(v => v.VmId == "100");

        Assert.Equal("web01", vm.Name);
        Assert.Equal(GuestKind.VirtualMachine, vm.Kind);
        Assert.False(vm.IsTemplate);
        Assert.Equal(Platform.Proxmox, vm.Platform);
        Assert.Equal("pve1", vm.Host);
        Assert.Equal("lab-cluster", vm.Cluster);
        Assert.Equal("London", vm.Datacenter);
        Assert.Equal("pve1.lab.local", vm.SourceAddress);
        Assert.Equal("Proxmox VE 8.2.4", vm.SourceProduct);
        Assert.Equal("prod", vm.ResourcePool);
        Assert.Equal(["web", "prod"], vm.Tags);
        Assert.Equal(PowerState.PoweredOn, vm.PowerState);
        Assert.Equal("running", vm.RawState);
        Assert.Equal("/etc/pve/nodes/pve1/qemu-server/100.conf", vm.ConfigPath);

        Assert.Equal(4, vm.CpuCount);
        Assert.Equal(2, vm.Sockets);
        Assert.Equal(2, vm.CoresPerSocket);
        Assert.Equal("x86-64-v2-AES", vm.CpuType);
        Assert.Equal(8192, vm.MemoryMiB);
        Assert.Equal(2048, vm.MemoryMinMiB);
        Assert.True(vm.DynamicMemory);
        Assert.Equal(6144, vm.MemoryAssignedMiB);
        Assert.Equal(2048, vm.MemoryBalloonedMiB);
        Assert.Equal(3072, vm.MemoryDemandMiB);

        Assert.Equal("efi", vm.Firmware);
        Assert.True(vm.SecureBoot);
        Assert.Equal("pc-q35-8.1", vm.HardwareVersion);
        Assert.Equal("Linux 2.6+ kernel", vm.OsConfigured);
        Assert.Equal("Web front end\nOwner: ops", vm.Annotation);
        Assert.Equal("7b0c3a62-1d4e-4c1a-9a55-6e1f2a3b4c5d", vm.BiosUuid);
        Assert.Equal("0f6a6d0e-2c1b-4a5e-8f3d-1a2b3c4d5e6f", vm.Uuid);
        Assert.True(vm.AutoStart);
        Assert.Equal(30, vm.StartDelaySeconds);
        Assert.Equal("scsi0, ide2, net0", vm.BootOrder);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(1704067200), vm.CreationDate);
        Assert.Equal(TimeSpan.FromDays(1), vm.Uptime);
        Assert.Equal(Now.AddDays(-1), vm.PowerOnTime);
        Assert.Contains("hostpci0", (string)vm.Extra["Fixed Passthru HotPlug"]!);

        // Disks: scsi0, scsi1, efidisk0 (the ISO is a CD drive, not a disk).
        Assert.Equal(["scsi0", "scsi1", "efidisk0"], vm.Disks.Select(d => d.Label));
        var scsi0 = vm.Disks[0];
        Assert.Equal(34359738368, scsi0.CapacityBytes);
        Assert.Equal("local-lvm", scsi0.Datastore);
        Assert.Equal("local-lvm:vm-100-disk-0", scsi0.Path);
        Assert.Equal("raw", scsi0.Format);
        Assert.True(scsi0.Thin);
        Assert.Equal(12884901888, scsi0.UsedBytes);
        Assert.Equal("SCSI (virtio-scsi-single)", scsi0.ControllerType);
        Assert.Contains("discard=on", scsi0.Options);
        Assert.Contains("iothread=1", scsi0.Options);
        Assert.Contains("ssd=1", scsi0.Options);
        var scsi1 = vm.Disks[1];
        Assert.Equal(107374182400, scsi1.CapacityBytes);
        Assert.Equal("nfs-shared", scsi1.Datastore);
        Assert.Equal("qcow2", scsi1.Format);
        Assert.True(scsi1.Thin);
        Assert.Equal("writeback", scsi1.Cache);
        Assert.Equal(21474836480, scsi1.UsedBytes);
        Assert.Equal(4194304, vm.Disks[2].CapacityBytes);

        var cd = Assert.Single(vm.CdDrives);
        Assert.Equal("ide2", cd.DeviceNode);
        Assert.Equal("local:iso/debian-12.5.0-amd64-netinst.iso", cd.Media);
        Assert.True(cd.Connected);

        Assert.Equal("host=046d:c52b", Assert.Single(vm.UsbDevices).DeviceType);

        // NICs with agent-reported IPs matched by MAC.
        Assert.Equal(2, vm.Nics.Count);
        var net0 = vm.Nics[0];
        Assert.Equal(("net0", "vmbr0", "BC:24:11:AA:BB:01"), (net0.Label, net0.Network, net0.Mac));
        Assert.Null(net0.Vlan);
        Assert.True(net0.Firewall);
        Assert.Equal("VirtIO (paravirtualized)", net0.AdapterType);
        Assert.Equal(["10.0.0.50"], net0.Ipv4);
        Assert.Equal(["fe80::be24:11ff:feaa:bb01"], net0.Ipv6);
        var net1 = vm.Nics[1];
        Assert.Equal(20, net1.Vlan);
        Assert.Equal(["10.20.0.50"], net1.Ipv4);
        Assert.Equal("10.0.0.50", vm.PrimaryIp);

        // Agent.
        Assert.True(vm.ToolsRunning);
        Assert.Equal("guestToolsRunning", vm.ToolsStatus);
        Assert.Equal("7.2.11", vm.ToolsVersion);
        Assert.Equal("green", vm.Heartbeat);
        Assert.Equal("Debian GNU/Linux 12 (bookworm)", vm.OsGuest);
        Assert.Equal("web01.lab.local", vm.DnsName);
        Assert.Equal(["/", "/boot/efi", "/srv"], vm.Partitions.Select(p => p.Name));
        var root = vm.Partitions[0];
        Assert.Equal(("ext4", 33501757440L, 33501757440L - 5368709120L), (root.FileSystem, root.CapacityBytes!.Value, root.FreeBytes!.Value));

        // Snapshot ("current" excluded).
        var snapItem = Assert.Single(vm.Snapshots);
        Assert.Equal("pre-upgrade", snapItem.Name);
        Assert.Equal("Before apt upgrade", snapItem.Description);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(1717200000), snapItem.Created);
        Assert.True(snapItem.IncludesMemory);

        // HA.
        Assert.True(vm.HaProtected);
        Assert.Equal("started", vm.HaState);
        Assert.Equal("prefer-pve1", vm.FailoverPriority);
        Assert.Equal("pve1", vm.OwnerNode);
        Assert.Equal(["pve1", "pve2"], vm.PreferredOwners);
    }

    [Fact]
    public async Task Collects_ContainerAndTemplate()
    {
        var snap = await Collect(PveFixtures.Cluster());

        var ct = snap.VirtualMachines.Single(v => v.VmId == "101");
        Assert.Equal(GuestKind.Container, ct.Kind);
        Assert.Equal("ct01", ct.Name);
        Assert.Equal("pve2", ct.Host);
        Assert.Equal(PowerState.PoweredOn, ct.PowerState);
        Assert.Equal(2, ct.CpuCount);
        Assert.Equal(1024, ct.MemoryMiB);
        Assert.Equal("LXC: debian", ct.OsConfigured);
        Assert.Equal("ct01", ct.DnsName);
        Assert.Equal("/etc/pve/nodes/pve2/lxc/101.conf", ct.ConfigPath);
        Assert.Contains("unprivileged='1'", (string)ct.Extra["Guest Detailed Data"]!);
        Assert.False(ct.AutoStart);
        Assert.False(ct.HaProtected);
        Assert.Empty(ct.Snapshots);

        Assert.Equal(["rootfs", "mp0"], ct.Disks.Select(d => d.Label));
        var rootfs = ct.Disks[0];
        Assert.Equal(8589934592, rootfs.CapacityBytes);
        Assert.Equal("local-lvm", rootfs.Datastore);
        Assert.Equal(1610612736, rootfs.UsedBytes); // from status "disk"
        var bind = ct.Disks[1];
        Assert.True(bind.Passthrough);
        Assert.Null(bind.Datastore);
        Assert.Contains("mp=/shared", bind.Options);

        var eth0 = Assert.Single(ct.Nics);
        Assert.Equal(("eth0", "vmbr0", 30, "BC:24:11:CC:DD:01"), (eth0.Label, eth0.Network, eth0.Vlan, eth0.Mac));
        Assert.Equal(["10.0.0.60"], eth0.Ipv4);
        Assert.Equal(["fd00::60"], eth0.Ipv6); // from /interfaces (ip6=dhcp in config)

        var tpl = snap.VirtualMachines.Single(v => v.VmId == "9000");
        Assert.True(tpl.IsTemplate);
        Assert.Equal(PowerState.PoweredOff, tpl.PowerState);
        Assert.Equal(10737418240, Assert.Single(tpl.Disks).CapacityBytes);
        Assert.Equal("SCSI (virtio-scsi-pci)", tpl.Disks[0].ControllerType);
        Assert.Equal("gray", tpl.Heartbeat);
        Assert.False(tpl.HaProtected);
    }

    [Fact]
    public async Task Collects_Datastores_SharedDedupedLocalPerNode()
    {
        var h = PveFixtures.Cluster();
        var snap = await Collect(h);

        var nfs = Assert.Single(snap.Datastores, d => d.Name == "nfs-shared");
        Assert.True(nfs.Shared);
        Assert.Equal(["pve1", "pve2"], nfs.Hosts);
        Assert.Equal("nfs", nfs.Type);
        Assert.Equal("10.0.0.5:/export/pve", nfs.Address);
        Assert.Equal(4000000000000, nfs.CapacityBytes);
        Assert.Equal(3000000000000, nfs.FreeBytes);
        Assert.True(nfs.Accessible);

        var lvm = snap.Datastores.Where(d => d.Name == "local-lvm").ToList();
        Assert.Equal(2, lvm.Count);
        Assert.Equal(["pve1"], lvm[0].Hosts);
        Assert.Equal(["pve2"], lvm[1].Hosts);
        Assert.Equal("pve/data", lvm[0].Address);
        Assert.False(lvm[0].Shared);
        Assert.Equal(5, snap.Datastores.Count);

        // One content listing per storage per node; shared storage listed once; non-image storage skipped.
        Assert.Single(h.Paths, p => p.Contains("/storage/nfs-shared/content"));
        Assert.DoesNotContain(h.Paths, p => p.Contains("/storage/local/content"));
        Assert.Equal(3, h.Paths.Count(p => p.EndsWith("/content")));
    }

    [Fact]
    public async Task AgentFailure_DoesNotFailCollection()
    {
        var h = PveFixtures.Cluster()
            .Status("/nodes/pve1/qemu/100/agent/info", HttpStatusCode.InternalServerError, "QEMU guest agent is not running");
        var snap = await Collect(h);

        var vm = snap.VirtualMachines.Single(v => v.VmId == "100");
        Assert.Equal(PowerState.PoweredOn, vm.PowerState);
        Assert.False(vm.ToolsRunning);
        Assert.Equal("guestToolsNotRunning", vm.ToolsStatus);
        Assert.Equal("gray", vm.Heartbeat);
        Assert.Empty(vm.Nics[0].Ipv4);
        Assert.Null(vm.OsGuest);
        Assert.Empty(vm.Partitions);
        Assert.DoesNotContain(h.Paths, p => p.EndsWith("/agent/get-osinfo")); // liveness probe short-circuits
        Assert.Equal(4, vm.Disks.Count + vm.CdDrives.Count);
    }

    [Fact]
    public async Task AgentTimeout_DoesNotFailCollection()
    {
        var h = PveFixtures.Cluster().Route("/nodes/pve1/qemu/100/agent/info", async (_, ct) =>
        {
            await Task.Delay(TimeSpan.FromSeconds(10), ct);
            return FakePveHandler.Json("{}");
        });
        var snap = await Collect(h);
        var vm = snap.VirtualMachines.Single(v => v.VmId == "100");
        Assert.False(vm.ToolsRunning);
        Assert.Equal("guestToolsNotRunning", vm.ToolsStatus);
    }

    [Fact]
    public async Task PerVmFailure_AddsWarningAndKeepsGoing()
    {
        var h = PveFixtures.Cluster()
            .Status("/nodes/pve2/lxc/101/config", HttpStatusCode.InternalServerError, "unable to parse config");
        var snap = await Collect(h);

        Assert.Equal(3, snap.VirtualMachines.Count);
        var ct = snap.VirtualMachines.Single(v => v.VmId == "101");
        Assert.Equal("ct01", ct.Name); // basic data from /cluster/resources
        Assert.Equal(2, ct.CpuCount);
        Assert.Contains(snap.Warnings, w => w.Contains("101") && w.Contains("unable to parse config"));
        Assert.Equal("guestToolsRunning", snap.VirtualMachines.Single(v => v.VmId == "100").ToolsStatus);
    }

    [Fact]
    public async Task OfflineNode_ListedRed_SkipsPerNodeCalls()
    {
        var h = PveFixtures.Cluster()
            .Data("/cluster/resources", PveFixtures.Resources.Replace(
                "\"node\":\"pve2\",\"status\":\"online\"", "\"node\":\"pve2\",\"status\":\"offline\""))
            .Data("/cluster/status", PveFixtures.ClusterStatus.Replace(
                "\"ip\":\"10.0.0.12\",\"online\":1", "\"ip\":\"10.0.0.12\",\"online\":0"));
        var snap = await Collect(h);

        var pve2 = snap.Hosts.Single(x => x.Name == "pve2");
        Assert.Equal("red", pve2.Status);
        Assert.Contains(snap.Warnings, w => w.Contains("pve2"));
        Assert.DoesNotContain(h.Paths, p => p.StartsWith("/nodes/pve2/"));
        var ct = snap.VirtualMachines.Single(v => v.VmId == "101");
        Assert.Equal(PowerState.Unknown, ct.PowerState);
        Assert.Equal(1, snap.Clusters[0].NumEffectiveHosts);
        Assert.Contains(snap.Health, i => i.Name == "pve2" && i.Severity == HealthSeverity.Error);
    }

    [Fact]
    public async Task StandaloneNode_HasNoCluster()
    {
        var h = PveFixtures.Cluster().Data("/cluster/status", PveFixtures.StandaloneStatus);
        var snap = await Collect(h);
        Assert.Empty(snap.Clusters);
        Assert.All(snap.Hosts, x => Assert.Null(x.Cluster));
        Assert.All(snap.VirtualMachines, v => Assert.Null(v.Cluster));
    }

    [Fact]
    public async Task Unauthorized_Token_MapsToAuthenticationFailedWithHint()
    {
        var h = PveFixtures.Cluster().Status("/version", HttpStatusCode.Unauthorized, "invalid token value!");
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h));
        Assert.Equal(CollectionFailure.AuthenticationFailed, ex.Kind);
        Assert.Contains("invalid token value!", ex.Message);
        Assert.Contains("user@realm!tokenname", ex.Hint);
        Assert.Contains("curl -k -H 'Authorization: PVEAPIToken=root@pam!explorer=SECRET'", ex.Hint);
    }

    [Fact]
    public async Task Unauthorized_Password_MapsToAuthenticationFailed()
    {
        var h = PveFixtures.Cluster().Status("/access/ticket", HttpStatusCode.Unauthorized, "authentication failure");
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h, PasswordRequest()));
        Assert.Equal(CollectionFailure.AuthenticationFailed, ex.Kind);
        Assert.Contains("root@pam", ex.Hint);
    }

    [Fact]
    public async Task Forbidden_MapsToPermissionDenied()
    {
        var h = PveFixtures.Cluster().Status("/cluster/resources", HttpStatusCode.Forbidden, "Permission check failed");
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h));
        Assert.Equal(CollectionFailure.PermissionDenied, ex.Kind);
        Assert.Contains("PVEAuditor", ex.Hint);
    }

    [Fact]
    public async Task NoVisibleNodes_MapsToPermissionDenied()
    {
        var h = PveFixtures.Cluster().Data("/cluster/resources", "[]");
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h));
        Assert.Equal(CollectionFailure.PermissionDenied, ex.Kind);
        Assert.Contains("Privilege Separation", ex.Hint);
    }

    [Fact]
    public async Task ConnectionRefused_MapsToUnreachable()
    {
        var h = PveFixtures.Cluster();
        h.ThrowOnSend = new HttpRequestException(HttpRequestError.ConnectionError, "Connection refused",
            new SocketException((int)SocketError.ConnectionRefused));
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h));
        Assert.Equal(CollectionFailure.Unreachable, ex.Kind);
        Assert.Contains("connection refused", ex.Message);
        Assert.Contains("8006", ex.Hint);
    }

    [Fact]
    public async Task TlsFailure_MapsToCertificate()
    {
        var h = PveFixtures.Cluster();
        h.ThrowOnSend = new HttpRequestException(HttpRequestError.SecureConnectionError, "SSL connection could not be established",
            new AuthenticationException("The remote certificate is invalid according to the validation procedure"));
        var ex = await Assert.ThrowsAsync<CollectionException>(() => Collect(h, TokenRequest() with { IgnoreCertificateErrors = false }));
        Assert.Equal(CollectionFailure.Certificate, ex.Kind);
        Assert.Contains("Ignore certificate errors", ex.Hint);
    }

    [Fact]
    public async Task Cancellation_MapsToCancelled()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            Collector(PveFixtures.Cluster()).CollectAsync(TokenRequest(), null, cts.Token));
        Assert.Equal(CollectionFailure.Cancelled, ex.Kind);
    }

    [Fact]
    public async Task ReportsProgress()
    {
        var messages = new List<string>();
        var progress = new SyncProgress(messages);
        await Collector(PveFixtures.Cluster()).CollectAsync(TokenRequest(), progress, CancellationToken.None);
        Assert.Contains(messages, m => m.StartsWith("Connecting to pve1.lab.local:8006"));
        Assert.Contains(messages, m => m.StartsWith("Collected 3/3 guests"));
    }

    private sealed class SyncProgress(List<string> sink) : IProgress<string>
    {
        public void Report(string value)
        {
            lock (sink) sink.Add(value);
        }
    }

    [Fact]
    public async Task Snapshot_FeedsRvToolsTables()
    {
        var snap = await Collect(PveFixtures.Cluster());
        var store = new InventoryStore();
        store.Upsert(snap);
        var inv = store.Current;

        // Every table renders without throwing.
        foreach (var table in RvToolsTables.All.Concat(ExtendedTables.All))
            foreach (var row in table.Rows(inv))
                foreach (var col in table.Columns)
                    _ = col.GetValue(row);

        var vInfo = RvToolsTables.Get("vInfo");
        object Row(string vm) => vInfo.Rows(inv).Single(r => vInfo.VmOf(r)!.Name == vm);
        object? Cell(TableDefinition t, object row, string header) => t.Columns.Single(c => c.Header == header).GetValue(row);

        var web = Row("web01");
        Assert.Equal("poweredOn", Cell(vInfo, web, "Powerstate"));
        Assert.Equal(4, Cell(vInfo, web, "CPUs"));
        Assert.Equal(8192L, Cell(vInfo, web, "Memory"));
        Assert.Equal(135172L, Cell(vInfo, web, "Provisioned MiB")); // 32 GiB + 100 GiB + 4 MiB efidisk
        Assert.Equal(2, Cell(vInfo, web, "NICs"));
        Assert.Equal(3, Cell(vInfo, web, "Disks"));
        Assert.Equal("10.0.0.50", Cell(vInfo, web, "Primary IP Address"));
        Assert.Equal("efi", Cell(vInfo, web, "Firmware"));
        Assert.Equal("True", Cell(vInfo, web, "EFI Secure boot"));
        Assert.Equal("green", Cell(vInfo, web, "Heartbeat"));
        Assert.Equal("lab-cluster", Cell(vInfo, web, "Cluster"));
        Assert.Equal("pve1", Cell(vInfo, web, "Host"));
        Assert.Equal("London", Cell(vInfo, web, "Datacenter"));
        Assert.Equal("Proxmox VE 8.2.4", Cell(vInfo, web, "VI SDK Server type"));
        Assert.Equal("Debian GNU/Linux 12 (bookworm)", Cell(vInfo, web, "OS according to the VMware Tools"));
        Assert.Equal("vmbr0", Cell(vInfo, web, "Network #1"));
        Assert.Equal(30000, Cell(vInfo, web, "Boot delay"));
        Assert.Equal("poweredOff", Cell(vInfo, Row("tpl-debian"), "Powerstate"));
        Assert.Equal("True", Cell(vInfo, Row("tpl-debian"), "Template"));

        var vDisk = RvToolsTables.Get("vDisk");
        Assert.Equal(3 + 2 + 1, vDisk.Rows(inv).Count);

        var vNetwork = RvToolsTables.Get("vNetwork");
        var tagged = vNetwork.Rows(inv).Single(r => Cell(vNetwork, r, "Mac Address") as string == "BC:24:11:AA:BB:02");
        Assert.Equal("vmbr0 (VLAN 20)", Cell(vNetwork, tagged, "Network"));

        var vDatastore = RvToolsTables.Get("vDatastore");
        var nfsRow = vDatastore.Rows(inv).Single(r => Cell(vDatastore, r, "Name") as string == "nfs-shared");
        Assert.Equal("pve1, pve2", Cell(vDatastore, nfsRow, "Hosts"));
        Assert.Equal(1, Cell(vDatastore, nfsRow, "# VMs total"));

        var vHost = RvToolsTables.Get("vHost");
        var hostRow = vHost.Rows(inv).Single(r => Cell(vHost, r, "Host") as string == "pve1");
        Assert.Equal(1, Cell(vHost, hostRow, "# VMs total"));
        Assert.Equal(32, Cell(vHost, hostRow, "# Cores"));

        Assert.Single(RvToolsTables.Get("vSnapshot").Rows(inv));
        Assert.Single(RvToolsTables.Get("vCluster").Rows(inv));
        Assert.Equal(3, RvToolsTables.Get("vPartition").Rows(inv).Count);
    }
}
