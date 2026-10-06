using HypervisorExplorer.Collectors.HyperV;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Tests.HyperV;

public class HyperVJsonMapperTests
{
    private static readonly ConnectionRequest ClusterRequest = new()
    {
        Platform = Platform.HyperV,
        Address = "hvclu01.contoso.local",
        CredentialKind = CredentialKind.CurrentUser,
        Group = "London DC",
    };

    private static InventorySnapshot MapCluster() => HyperVJsonMapper.Map(HyperVFixtures.ClusterData, ClusterRequest);

    [Fact]
    public void Source_IsHyperVWithOsVersion()
    {
        var s = MapCluster().Source;
        Assert.Equal("hvclu01.contoso.local", s.Address);
        Assert.Equal(Platform.HyperV, s.Platform);
        Assert.Equal("London DC", s.Group);
        Assert.Equal("Microsoft Hyper-V", s.ProductName);
        Assert.Equal("10.0.20348", s.Version);
        Assert.Equal("20348", s.Build);
        Assert.Equal("WinRM", s.ApiVersion);
        Assert.Equal("Microsoft Corporation", s.Vendor);
        Assert.Equal("Windows", s.OsType);
        Assert.Contains("Windows Server 2022 Datacenter", (string)s.Extra["Fullname"]!);
    }

    [Fact]
    public void Hosts_ComeFromEveryCollectedNode()
    {
        var snap = MapCluster();
        Assert.Equal(["HV01", "HV02"], snap.Hosts.Select(h => h.Name));
        var hv01 = snap.Hosts[0];
        Assert.Equal("HVCLU01", hv01.Cluster);
        Assert.Equal("London DC", hv01.Datacenter);
        Assert.Equal("Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz", hv01.CpuModel);
        Assert.Equal(2000, hv01.CpuMhz);
        Assert.Equal(2, hv01.CpuSockets);
        Assert.Equal(16, hv01.CoresPerSocket);
        Assert.Equal(32, hv01.CpuCores);
        Assert.Equal(64, hv01.CpuThreads);
        Assert.True(hv01.HyperThreadingActive);
        Assert.Equal(15, hv01.CpuUsagePercent);
        Assert.Equal(549755813888, hv01.MemoryBytes);
        Assert.Equal((536870912L - 268435456L) * 1024, hv01.MemoryUsedBytes);
        Assert.Equal("Microsoft Windows Server 2022 Datacenter 10.0.20348", hv01.Version);
        Assert.Equal(new DateTimeOffset(2026, 9, 1, 3, 15, 0, TimeSpan.FromHours(1)), hv01.BootTime);
        Assert.Equal("Dell Inc.", hv01.Vendor);
        Assert.Equal("PowerEdge R750", hv01.Model);
        Assert.Equal("ABC1234", hv01.SerialNumber);
        Assert.Equal("1.10.2", hv01.BiosVersion);
        Assert.Equal("4c4c4544-0042-3510-8051-b4c04f333332", hv01.Uuid);
        Assert.Equal(["10.0.0.10", "10.0.0.11"], hv01.DnsServers);
        Assert.Equal("contoso.local", hv01.Domain);
        Assert.Equal(["dc01.contoso.local"], hv01.NtpServers);
        Assert.Equal("GMT Standard Time", hv01.TimeZone);
        Assert.Equal(true, hv01.Extra["VMotion support"]);

        var hv02 = snap.Hosts[1];
        Assert.False(hv02.HyperThreadingActive);
        Assert.Equal(["dc01.contoso.local"], hv02.NtpServers); // ",0x9" flag stripped
    }

    [Fact]
    public void HostNetworking_SwitchesNicsPortGroupsAndIps()
    {
        var hv01 = MapCluster().Hosts[0];

        var sw = Assert.Single(hv01.Switches);
        Assert.Equal("SETswitch", sw.Name);
        Assert.Equal("External (SET)", sw.Type);
        Assert.Equal(["SLOT 3 Port 1", "SLOT 3 Port 2"], sw.Uplinks);

        Assert.Equal(3, hv01.Nics.Count);
        var p1 = hv01.Nics[0];
        Assert.Equal(25000, p1.SpeedMbps);
        Assert.Equal("b8:59:9f:00:00:01", p1.Mac);
        Assert.Equal("SETswitch", p1.Switch);
        Assert.Equal("mlx5.sys", p1.Driver);
        Assert.Equal("0000:3b:00.0", p1.Pci);
        Assert.Equal(0, hv01.Nics[2].SpeedMbps); // disconnected
        Assert.Null(hv01.Nics[2].Switch);

        Assert.Equal([("Management", 10), ("LiveMigration", 30)], hv01.PortGroups.Select(p => (p.Name, p.Vlan ?? -1)));

        var mgmt = hv01.IpInterfaces.First(i => i.PortGroup == "Management");
        Assert.Equal("vEthernet (Management)", mgmt.Name);
        Assert.Equal("10.0.10.21", mgmt.Ipv4);
        Assert.Equal("255.255.255.0", mgmt.SubnetMask);
        Assert.Equal("10.0.10.1", mgmt.Gateway);
        Assert.False(mgmt.Dhcp);
        Assert.Equal(9000, hv01.IpInterfaces.First(i => i.PortGroup == "LiveMigration").Mtu);

        var fc = hv01.StorageAdapters.First(a => a.Type == "Fibre Channel");
        Assert.Equal("20:00:00:90:fa:00:00:01 10:00:00:90:fa:00:00:01", fc.Wwn);
        Assert.Contains(hv01.StorageAdapters, a => a.Type == "SCSI" && a.Driver == "SmartPqi");
    }

    [Fact]
    public void Cluster_NodesNetworksQuorumAndTotals()
    {
        var snap = MapCluster();
        var c = Assert.Single(snap.Clusters);
        Assert.Equal("HVCLU01", c.Name);
        Assert.Equal(Platform.HyperV, c.Platform);
        Assert.Equal("contoso.local", c.Domain);
        Assert.Equal("a1b2c3d4-1111-2222-3333-444455556666", c.Id);
        Assert.Equal("NodeAndFileShareMajority", c.QuorumType);
        Assert.Equal("File Share Witness", c.QuorumWitness);
        Assert.Equal("11 (Windows Server 2022)", c.FunctionalLevel);
        Assert.Equal(3, c.NumHosts);
        Assert.Equal(2, c.NumEffectiveHosts);
        Assert.Equal(32 + 24, c.NumCpuCores);
        Assert.Equal(64 + 24, c.NumCpuThreads);
        Assert.Equal(549755813888 + 412316860416, c.TotalMemoryBytes);
        Assert.Equal(32L * 2000 + 24L * 2100, c.TotalCpuMhz);
        Assert.Equal("yellow", c.OverallStatus);
        Assert.True(c.HaEnabled);
        Assert.Equal("2 of 3 nodes collected", c.Extra["Config status"]);

        Assert.Equal(["HV01", "HV02", "HV03"], c.Nodes.Select(n => n.Name));
        Assert.Equal("10.0.10.22", c.Nodes[1].Address); // client-facing network preferred
        Assert.Equal("Down", c.Nodes[2].State);
        Assert.Equal(1, c.Nodes[0].Votes);
        Assert.Equal("NotInitiated", c.Nodes[0].DrainStatus);

        Assert.Equal(["Cluster and client", "Cluster only", "None"], c.Networks.Select(n => n.Role));
        Assert.Equal("10.0.10.0/24", c.Networks[0].Address);
        Assert.Equal("10.0.40.0 / 255.255.255.0", c.Networks[2].Address);
        Assert.Equal("70384 (auto)", c.Networks[0].Metric);
        Assert.Equal("79999", c.Networks[2].Metric);

        Assert.Contains(snap.Health, h => h.Name == "HV03" && h.Message.Contains("Down"));
        Assert.Contains(snap.Health, h => h.Name == "OLDVM" && h.Severity == HealthSeverity.Error);
        Assert.Contains(snap.Health, h => h.Name == "Cluster Disk 2" && h.Message.Contains("redirected"));
        Assert.Contains("Cluster node HV03 is Down; not collected.", snap.Warnings);
        Assert.Contains(snap.Warnings, w => w.Contains("Sample warning"));
    }

    [Fact]
    public void Datastores_LocalVolumesCsvsAndSmbShares()
    {
        var snap = MapCluster();
        var names = snap.Datastores.Select(d => d.Name).ToList();
        Assert.Contains("HV01 C:", names);
        Assert.Contains("HV01 D: (Data)", names);
        Assert.Contains("HV02 C:", names);
        Assert.Contains(@"\\fs01\vms", names);

        var csv1 = snap.Datastores.Single(d => d.Name == "Cluster Disk 1");
        Assert.Equal("CSV (NTFS)", csv1.Type);
        Assert.True(csv1.Shared);
        Assert.Equal(["HV01", "HV02", "HV03"], csv1.Hosts);
        Assert.Equal(4398046511104, csv1.CapacityBytes);
        Assert.Equal(2199023255552, csv1.FreeBytes);
        Assert.Equal(@"C:\ClusterStorage\Volume1", csv1.Content);
        Assert.Equal("HVCLU01", csv1.Cluster);

        var local = snap.Datastores.Single(d => d.Name == "HV01 D: (Data)");
        Assert.False(local.Shared);
        Assert.Equal(["HV01"], local.Hosts);
        Assert.Equal("ReFS", local.Type);

        var smb = snap.Datastores.Single(d => d.Name == @"\\fs01\vms");
        Assert.Equal("SMB", smb.Type);
        Assert.Equal(["HV02"], smb.Hosts);
    }

    [Fact]
    public void Gen2ClusteredVm_IsFullyMapped()
    {
        var snap = MapCluster();
        var vm = snap.VirtualMachines.Single(v => v.Name == "SQL01");

        Assert.Equal("6f1d3a2e-8b4c-4d5e-9f00-112233445566", vm.VmId);
        Assert.Equal(vm.VmId, vm.Uuid);
        Assert.Equal("2b9f1e4a-6c7d-4e8f-9a0b-1c2d3e4f5a6b", vm.BiosUuid);
        Assert.Equal("HV01", vm.Host);
        Assert.Equal("HVCLU01", vm.Cluster);
        Assert.Equal("London DC", vm.Datacenter);
        Assert.Equal("hvclu01.contoso.local", vm.SourceAddress);
        Assert.Equal("Microsoft Hyper-V Microsoft Windows Server 2022 Datacenter", vm.SourceProduct);
        Assert.Equal(PowerState.PoweredOn, vm.PowerState);
        Assert.Equal("Running", vm.RawState);
        Assert.Equal(TimeSpan.FromDays(1), vm.Uptime);
        Assert.Equal(new DateTimeOffset(2026, 10, 5, 10, 0, 0, TimeSpan.FromHours(1)), vm.PowerOnTime);
        Assert.Equal(8, vm.CpuCount);
        Assert.Equal(100, vm.CpuShares);
        Assert.Equal(16384, vm.MemoryMiB);
        Assert.Equal(20480, vm.MemoryAssignedMiB);
        Assert.Equal(18432, vm.MemoryDemandMiB);
        Assert.Equal(8192, vm.MemoryMinMiB);
        Assert.Equal(32768, vm.MemoryMaxMiB);
        Assert.True(vm.DynamicMemory);
        Assert.Equal("efi", vm.Firmware);
        Assert.True(vm.SecureBoot);
        Assert.Equal("MicrosoftWindows", vm.Extra["Secure boot template"]);
        Assert.Equal("Gen 2", vm.Generation);
        Assert.Equal("10.0", vm.HardwareVersion);
        Assert.Equal("Hard drive, DVD, Network, File (Windows Boot Manager)", vm.BootOrder);
        Assert.Equal("Windows Server 2022 Datacenter", vm.OsGuest);
        Assert.Equal("sql01.contoso.local", vm.DnsName);
        Assert.Equal("Production SQL", vm.Annotation);
        Assert.Equal(@"C:\ClusterStorage\Volume1\SQL01", vm.ConfigPath);
        Assert.Equal(@"C:\ClusterStorage\Volume1\SQL01\Snapshots", vm.SnapshotDirectory);
        Assert.True(vm.AutoStart);
        Assert.Equal(120, vm.StartDelaySeconds);

        // Failover cluster role
        Assert.True(vm.HaProtected);
        Assert.Equal("Online", vm.HaState);
        Assert.Equal("HV01", vm.OwnerNode);
        Assert.Equal(["HV01", "HV02"], vm.PreferredOwners);
        Assert.Equal("High", vm.FailoverPriority);
        Assert.Equal("high", vm.Extra["HA Restart Priority"]);

        // Integration services
        Assert.Equal("green", vm.Heartbeat);
        Assert.True(vm.ToolsRunning);
        Assert.Equal("toolsOk", vm.ToolsStatus);
        Assert.Equal("10.0.20348.2849", vm.ToolsVersion);
        Assert.Equal(6, vm.IntegrationComponents.Count);
        Assert.Contains(vm.IntegrationComponents, ic => ic.Name == "Guest Service Interface" && !ic.Enabled);

        // Disks
        Assert.Equal(2, vm.Disks.Count);
        var os = vm.Disks[0];
        Assert.Equal("Hard disk 1", os.Label);
        Assert.Equal("SCSI 0:0", os.Controller);
        Assert.Equal("Cluster Disk 1", os.Datastore);
        Assert.Equal("vhdx", os.Format);
        Assert.True(os.Thin);
        Assert.Equal(136365211648, os.CapacityBytes);
        Assert.Equal(42949672960, os.UsedBytes);
        Assert.False(os.Passthrough);
        var data = vm.Disks[1];
        Assert.Equal("Cluster Disk 2", data.Datastore); // Volume10, not Volume1, and case-insensitive
        Assert.True(data.Thin);
        Assert.Equal("Differencing", data.Options);
        Assert.Equal(@"C:\ClusterStorage\Volume10\SQL01\SQL01_Data.vhdx", data.ParentPath);
        Assert.Equal(1, data.Unit);

        // DVD
        var cd = Assert.Single(vm.CdDrives);
        Assert.Equal("SCSI 0:2", cd.DeviceNode);
        Assert.True(cd.Connected);
        Assert.Equal(@"C:\ClusterStorage\Volume1\ISO\SQLServer2022.iso", cd.Media);
        Assert.StartsWith("ISO ", cd.DeviceType);

        // NIC
        var nic = Assert.Single(vm.Nics);
        Assert.Equal("Network Adapter", nic.Label);
        Assert.Equal("SETswitch", nic.Network);
        Assert.Equal(20, nic.Vlan);
        Assert.Equal("00:15:5d:0a:10:01", nic.Mac);
        Assert.Equal(["10.0.20.50"], nic.Ipv4);
        Assert.Equal(["fe80::1c2d:3e4f:5a6b:7c8d", "2001:db8:20::50"], nic.Ipv6);
        Assert.True(nic.Connected);
        Assert.Equal("Synthetic", nic.AdapterType);
        Assert.Equal("Static MAC", nic.Options);
        Assert.Equal("10.0.20.50", vm.PrimaryIp);

        // Checkpoint
        var snapShot = Assert.Single(vm.Snapshots);
        Assert.Equal("Before CU12", snapShot.Name);
        Assert.Equal("Production", snapShot.Type);
        Assert.Equal(new DateTimeOffset(2026, 9, 20, 22, 0, 0, TimeSpan.FromHours(1)), snapShot.Created);
    }

    [Fact]
    public void Gen1VmWithPassthroughAndLegacyNic()
    {
        var vm = MapCluster().VirtualMachines.Single(v => v.Name == "WEB01");
        Assert.Equal(PowerState.PoweredOff, vm.PowerState);
        Assert.Equal("bios", vm.Firmware);
        Assert.False(vm.SecureBoot);
        Assert.Equal("Gen 1", vm.Generation);
        Assert.Equal("CD, IDE, LegacyNetworkAdapter, Floppy", vm.BootOrder);
        Assert.False(vm.AutoStart);
        Assert.Null(vm.PowerOnTime);
        Assert.Equal("gray", vm.Heartbeat);
        Assert.Equal("toolsNotRunning", vm.ToolsStatus);
        Assert.Null(vm.ToolsVersion); // "0.0" means not installed
        Assert.False(vm.HaProtected);
        Assert.Null(vm.OwnerNode);
        Assert.Equal(10, vm.CpuReservation);
        Assert.Equal(75, vm.CpuLimit);
        Assert.Equal(200, vm.CpuShares);

        Assert.Equal("HV01 D: (Data)", vm.Disks[0].Datastore);
        Assert.Equal("vhd", vm.Disks[0].Format);
        Assert.False(vm.Disks[0].Thin);
        Assert.Equal("IDE 0:0", vm.Disks[0].Controller);
        Assert.True(vm.Disks[1].Passthrough);
        Assert.Null(vm.Disks[1].Datastore);
        Assert.Equal("Disk 4", vm.Disks[1].Extra["Raw LUN ID"]);

        var cd = Assert.Single(vm.CdDrives);
        Assert.False(cd.Connected);
        Assert.Equal("Empty", cd.DeviceType);

        var nic = Assert.Single(vm.Nics);
        Assert.Equal("Legacy Network Adapter", nic.AdapterType);
        Assert.Equal("Not connected", nic.Network);
        Assert.False(nic.Connected);
        Assert.Null(nic.Vlan);
    }

    [Fact]
    public void ExpandedNodeVm_MatchesRoleByVmIdAndToleratesUnrolledArrays()
    {
        var vm = MapCluster().VirtualMachines.Single(v => v.Name == "APP01");
        Assert.Equal("HV02", vm.Host);
        Assert.Equal("HVCLU01", vm.Cluster);
        Assert.Equal("Low", vm.FailoverPriority);    // role "SCVMM APP01 Resources" matched by VmId
        Assert.Equal("HV02", vm.OwnerNode);
        Assert.Empty(vm.PreferredOwners);
        Assert.Equal("red", vm.Heartbeat);           // LostCommunication
        Assert.False(vm.ToolsRunning);
        Assert.Equal("toolsNotRunning", vm.ToolsStatus);
        Assert.Equal(@"\\fs01\vms", vm.Disks[0].Datastore);
        Assert.False(vm.Disks[0].Thin);
        Assert.True(vm.Disks[1].Shared);             // VHD Set
        Assert.Equal("vhdset", vm.Disks[1].Format);

        var nic = Assert.Single(vm.Nics);            // single object instead of array
        Assert.Equal(["10.0.20.60"], nic.Ipv4);      // single string instead of array
        Assert.Null(nic.Vlan);                       // trunk mode: no access VLAN
        Assert.Contains("Trunk 20-25 native 1", nic.Options);
    }

    [Fact]
    public void StandaloneHost_ManualRunDocument()
    {
        var req = new ConnectionRequest { Platform = Platform.HyperV, Address = "HVSTD01" };
        var snap = HyperVJsonMapper.Map(HyperVFixtures.StandaloneHost, req);

        Assert.Empty(snap.Clusters);
        var host = Assert.Single(snap.Hosts);
        Assert.Equal("HVSTD01", host.Name);
        Assert.Null(host.Cluster);
        Assert.Equal(1, host.CpuSockets);            // processors was a lone object
        Assert.Equal(8, host.CpuCores);
        Assert.Equal(["Local CMOS Clock"], host.NtpServers);
        Assert.Equal(1.0, host.Extra["GMT Offset"]);
        Assert.Equal("(UTC+01:00) Amsterdam, Berlin", host.Extra["Time Zone Name"]);
        Assert.Equal(["Embedded LOM 1 Port 1"], Assert.Single(host.Switches).Uplinks);
        Assert.Equal(0, Assert.Single(host.PortGroups).Vlan);
        var ip = Assert.Single(host.IpInterfaces);
        Assert.Equal("192.168.1.20", ip.Ipv4);
        Assert.True(ip.Dhcp);
        Assert.Equal("External", ip.PortGroup);

        var ds = Assert.Single(snap.Datastores);
        Assert.Equal("HVSTD01 C: (OS)", ds.Name);

        var vm = Assert.Single(snap.VirtualMachines);
        Assert.Equal(PowerState.Suspended, vm.PowerState);
        Assert.Null(vm.Cluster);
        Assert.Null(vm.HaProtected);                 // standalone: no HA concept
        Assert.Equal("HVSTD01 C: (OS)", vm.Disks[0].Datastore);
        Assert.Equal(2, vm.Snapshots.Count);
        Assert.Equal("DC01 - (01/09/2026 - 10:00:00)", vm.Snapshots[1].Parent);
        Assert.Equal(30, vm.StartDelaySeconds);
        Assert.Equal(4096, vm.MemoryMiB);
        Assert.Equal("gray", vm.Heartbeat);
    }

    [Fact]
    public void Map_AcceptsEnvelopeAndUtf16File()
    {
        var req = new ConnectionRequest { Platform = Platform.HyperV, Address = "hvclu01" };
        var fromEnvelope = HyperVJsonMapper.Map(HyperVFixtures.Envelope(HyperVFixtures.ClusterData), req);
        Assert.Equal(3, fromEnvelope.VirtualMachines.Count);

        // Windows PowerShell 5.1 "> file.json" writes UTF-16LE with a BOM.
        var path = Path.Combine(Path.GetTempPath(), $"hvx-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, HyperVFixtures.StandaloneHost, new System.Text.UnicodeEncoding(false, true));
            Assert.Single(HyperVJsonMapper.MapFile(path, req).VirtualMachines);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Map_RejectsForeignJson()
    {
        var req = new ConnectionRequest { Platform = Platform.HyperV, Address = "x" };
        var ex = Assert.Throws<CollectionException>(() => HyperVJsonMapper.Map("""{"foo": 1}""", req));
        Assert.Equal(CollectionFailure.Protocol, ex.Kind);
        ex = Assert.Throws<CollectionException>(() => HyperVJsonMapper.Map("not json", req));
        Assert.Equal(CollectionFailure.Protocol, ex.Kind);
    }

    [Theory]
    [InlineData(@"C:\ClusterStorage\Volume1\VMs\a.vhdx", "CSV1")]
    [InlineData(@"C:\ClusterStorage\Volume10\VMs\a.vhdx", "CSV10")]
    [InlineData(@"c:\clusterstorage\VOLUME10\a.vhdx", "CSV10")]
    [InlineData(@"C:\ClusterStorage\Volume1", "CSV1")]
    [InlineData(@"C:\ClusterStorage\Volume100\a.vhdx", "Local C")]
    [InlineData(@"C:/Hyper-V/a.vhdx", "Local C")]
    [InlineData(@"\\?\C:\ClusterStorage\Volume1\a.vhdx", "CSV1")]
    [InlineData(@"D:\Mounts\Fast\vm\a.vhdx", "Fast")]
    [InlineData(@"D:\Other\a.vhdx", "Local D")]
    [InlineData(@"E:\x.vhdx", null)]
    [InlineData(@"\\fs01\vms\a.vhdx", null)]
    [InlineData("", null)]
    public void MatchDatastore_UsesLongestMountPrefix(string path, string? expected)
    {
        var mounts = new Dictionary<string, string>
        {
            [@"C:\"] = "Local C",
            [@"D:\"] = "Local D",
            [@"D:\Mounts\Fast\"] = "Fast",
            [@"C:\ClusterStorage\Volume1"] = "CSV1",
            [@"C:\ClusterStorage\Volume10"] = "CSV10",
        };
        Assert.Equal(expected, HyperVJsonMapper.MatchDatastore(path, mounts));
    }

    [Theory]
    [InlineData(@"\\fs01\vms\APP01\a.vhdx", @"\\fs01\vms")]
    [InlineData(@"\\?\UNC\fs01\vms\a.vhdx", @"\\fs01\vms")]
    [InlineData(@"\\fs01", null)]
    [InlineData(@"C:\vms\a.vhdx", null)]
    public void SmbShareRoot(string path, string? expected) =>
        Assert.Equal(expected, HyperVJsonMapper.SmbShareRoot(path));

    [Theory]
    [InlineData("Running", PowerState.PoweredOn)]
    [InlineData("RunningCritical", PowerState.PoweredOn)]
    [InlineData("Off", PowerState.PoweredOff)]
    [InlineData("OffCritical", PowerState.PoweredOff)]
    [InlineData("Saved", PowerState.Suspended)]
    [InlineData("Paused", PowerState.Suspended)]
    [InlineData("Starting", PowerState.PoweredOn)]
    [InlineData("Other", PowerState.Unknown)]
    [InlineData(null, PowerState.Unknown)]
    public void MapPowerState(string? state, PowerState expected) =>
        Assert.Equal(expected, HyperVJsonMapper.MapPowerState(state));

    [Theory]
    [InlineData("OkApplicationsHealthy", null, "green")]
    [InlineData("OkApplicationsCritical", null, "yellow")]
    [InlineData(null, "Ok", "green")]
    [InlineData("NoContact", null, "gray")]
    [InlineData("LostCommunication", null, "red")]
    [InlineData(null, "Lost Communication", "red")]
    [InlineData(null, null, "gray")]
    public void MapHeartbeat(string? vmHeartbeat, string? icStatus, string expected) =>
        Assert.Equal(expected, HyperVJsonMapper.MapHeartbeat(vmHeartbeat, icStatus, PowerState.PoweredOn));

    [Theory]
    [InlineData(3000, "High")]
    [InlineData(2000, "Medium")]
    [InlineData(1000, "Low")]
    [InlineData(0, "No auto start")]
    public void PriorityText(long value, string expected) => Assert.Equal(expected, HyperVJsonMapper.PriorityText(value));

    [Theory]
    [InlineData(24, "255.255.255.0")]
    [InlineData(16, "255.255.0.0")]
    [InlineData(26, "255.255.255.192")]
    [InlineData(32, "255.255.255.255")]
    [InlineData(0, "0.0.0.0")]
    public void PrefixToMask(long prefix, string expected) => Assert.Equal(expected, HyperVJsonMapper.PrefixToMask(prefix));
}
