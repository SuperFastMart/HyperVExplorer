namespace HypervisorExplorer.Core.Model;

/// <summary>
/// Deterministic fictional estate (Hyper-V failover cluster, Proxmox cluster, VMware cluster) for demo mode and tests.
/// All names and addresses are made up (RFC 5737 / example.com).
/// </summary>
public static class SampleInventory
{
    public static IReadOnlyList<InventorySnapshot> Create(int seed = 42)
    {
        var rng = new Random(seed);
        return [HyperV(rng), Proxmox(rng), VMware(rng)];
    }

    private static readonly string[] Roles = ["web", "app", "sql", "dc", "file", "print", "rds", "mon", "build", "proxy", "cache", "mq"];

    private static InventorySnapshot HyperV(Random rng)
    {
        const string src = "hv-clu01.example.com";
        var source = new Source
        {
            Address = src, Platform = Platform.HyperV, Group = "London", ProductName = "Microsoft Hyper-V",
            Version = "10.0.20348", Build = "20348", ApiVersion = "WinRM", Vendor = "Microsoft Corporation", OsType = "Windows",
        };
        var cluster = new ClusterInfo
        {
            Name = "HV-CLU01", Platform = Platform.HyperV, SourceAddress = src, Datacenter = "London",
            QuorumType = "NodeAndFileShareMajority", QuorumWitness = @"\\witness.example.com\hvclu01", FunctionalLevel = "11",
            Quorate = true, HaEnabled = true, Domain = "example.com",
            Networks =
            [
                new ClusterNetwork { Name = "Management", Role = "ClusterAndClient", Address = "192.0.2.0/24", State = "Up", Metric = "70384" },
                new ClusterNetwork { Name = "Live Migration", Role = "Cluster", Address = "198.51.100.0/24", State = "Up", Metric = "39840" },
            ],
        };
        var snap = new InventorySnapshot { Source = source, Clusters = { cluster } };
        var csv = new Datastore
        {
            Name = "Volume1", Platform = Platform.HyperV, SourceAddress = src, Cluster = cluster.Name, Type = "CSVFS",
            Address = @"C:\ClusterStorage\Volume1", Shared = true, CapacityBytes = 8L << 40, FreeBytes = (long)(1.1 * (1L << 40)),
        };
        snap.Datastores.Add(csv);

        for (var n = 1; n <= 2; n++)
        {
            var host = new HostSystem
            {
                Name = $"HV-NODE0{n}", Platform = Platform.HyperV, SourceAddress = src, Datacenter = "London", Cluster = cluster.Name,
                Status = "green", CpuModel = "Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz", CpuMhz = 2000, CpuSockets = 2,
                CoresPerSocket = 32, CpuCores = 64, CpuThreads = 128, HyperThreadingActive = true, CpuUsagePercent = 18 + n * 4,
                MemoryBytes = 768L << 30, MemoryUsedBytes = (long)((0.55 + n * 0.05) * (768L << 30)),
                Version = "Windows Server 2022 Datacenter 10.0.20348", BootTime = DateTimeOffset.Now.AddDays(-23 - n),
                Vendor = "Dell Inc.", Model = "PowerEdge R750", SerialNumber = $"HVSN00{n}", BiosVendor = "Dell Inc.", BiosVersion = "1.10.2",
                DnsServers = ["192.0.2.10", "192.0.2.11"], Domain = "example.com", NtpServers = ["dc01.example.com"], TimeZone = "GMT Standard Time",
                Nics =
                [
                    new HostNic { Name = "NIC1", Description = "Intel(R) Ethernet 25G 2P E810-XXV", SpeedMbps = 25000, FullDuplex = true, Mac = $"00-15-5D-00-0{n}-01", Switch = "SET-Switch", Status = "Up" },
                    new HostNic { Name = "NIC2", Description = "Intel(R) Ethernet 25G 2P E810-XXV", SpeedMbps = 25000, FullDuplex = true, Mac = $"00-15-5D-00-0{n}-02", Switch = "SET-Switch", Status = "Up" },
                ],
                Switches = [new VirtualSwitch { Name = "SET-Switch", Type = "External (SET)", Uplinks = ["NIC1", "NIC2"] }],
                PortGroups = [new PortGroup { Name = "vEthernet (Management)", Switch = "SET-Switch", Vlan = 10 }],
                IpInterfaces = [new HostIpInterface { Name = "vEthernet (Management)", PortGroup = "SET-Switch", Ipv4 = $"192.0.2.2{n}", SubnetMask = "255.255.255.0", Gateway = "192.0.2.1", Dhcp = false }],
            };
            snap.Hosts.Add(host);
            csv.Hosts.Add(host.Name);
            cluster.Nodes.Add(new ClusterNode { Name = host.Name, State = "Up", DrainStatus = "NotInitiated", Id = n.ToString(), Votes = 1 });
            for (var i = 0; i < 7; i++) snap.VirtualMachines.Add(Vm(rng, Platform.HyperV, host, cluster.Name, src, "London", csv.Name, $"LON-{Roles[(i + n * 3) % Roles.Length].ToUpperInvariant()}{n}{i:00}"));
        }
        cluster.NumHosts = cluster.NumEffectiveHosts = 2;
        cluster.NumCpuCores = snap.Hosts.Sum(h => h.CpuCores ?? 0);
        cluster.NumCpuThreads = snap.Hosts.Sum(h => h.CpuThreads ?? 0);
        cluster.TotalMemoryBytes = snap.Hosts.Sum(h => h.MemoryBytes ?? 0);
        cluster.TotalCpuMhz = snap.Hosts.Sum(h => (long)(h.CpuCores ?? 0) * (h.CpuMhz ?? 0));
        return snap;
    }

    private static InventorySnapshot Proxmox(Random rng)
    {
        const string src = "192.0.2.50";
        var source = new Source
        {
            Address = src, Platform = Platform.Proxmox, Group = "Manchester", ProductName = "Proxmox VE", Version = "8.2.4",
            ApiVersion = "8.2", Vendor = "Proxmox Server Solutions GmbH", OsType = "Linux",
        };
        var cluster = new ClusterInfo { Name = "pve-man", Platform = Platform.Proxmox, SourceAddress = src, Datacenter = "Manchester", Quorate = true, HaEnabled = true };
        var snap = new InventorySnapshot { Source = source, Clusters = { cluster } };
        var ceph = new Datastore { Name = "ceph-vm", Platform = Platform.Proxmox, SourceAddress = src, Cluster = cluster.Name, Type = "rbd", Shared = true, CapacityBytes = 20L << 40, FreeBytes = 9L << 40, Content = "images,rootdir" };
        snap.Datastores.Add(ceph);
        for (var n = 1; n <= 3; n++)
        {
            var host = new HostSystem
            {
                Name = $"pve{n}", Platform = Platform.Proxmox, SourceAddress = src, Datacenter = "Manchester", Cluster = cluster.Name, Status = "green",
                CpuModel = "AMD EPYC 7443P 24-Core Processor", CpuMhz = 2850, CpuSockets = 1, CoresPerSocket = 24, CpuCores = 24, CpuThreads = 48,
                HyperThreadingActive = true, CpuUsagePercent = 9 + n * 6, MemoryBytes = 512L << 30, MemoryUsedBytes = (long)(0.4 * (512L << 30)) + n * (20L << 30),
                Version = "pve-manager/8.2.4", KernelVersion = "6.8.12-1-pve", BootTime = DateTimeOffset.Now.AddDays(-60 + n),
                DnsServers = ["192.0.2.10"], Domain = "example.com", TimeZone = "Europe/London",
                Nics = [new HostNic { Name = "enp65s0f0", SpeedMbps = 10000, Mac = $"bc:24:11:00:00:0{n}", Switch = "vmbr0", Status = "active" }],
                Switches = [new VirtualSwitch { Name = "vmbr0", Type = "Linux bridge (VLAN aware)", Uplinks = ["enp65s0f0"], Mtu = 1500 }],
                IpInterfaces = [new HostIpInterface { Name = "vmbr0", PortGroup = "vmbr0", Ipv4 = $"192.0.2.5{n}", SubnetMask = "255.255.255.0", Gateway = "192.0.2.1" }],
            };
            snap.Hosts.Add(host);
            ceph.Hosts.Add(host.Name);
            snap.Datastores.Add(new Datastore { Name = "local-lvm", Platform = Platform.Proxmox, SourceAddress = src, Cluster = cluster.Name, Type = "lvmthin", Shared = false, Hosts = [host.Name], CapacityBytes = 900L << 30, FreeBytes = (700L - n * 60) << 30, Content = "images,rootdir" });
            cluster.Nodes.Add(new ClusterNode { Name = host.Name, State = "online", Id = n.ToString(), Address = $"192.0.2.5{n}", Votes = 1 });
            for (var i = 0; i < 6; i++)
            {
                var vm = Vm(rng, Platform.Proxmox, host, cluster.Name, src, "Manchester", i % 3 == 0 ? "local-lvm" : "ceph-vm", $"man-{Roles[(i + n) % Roles.Length]}-{n}{i:00}");
                vm.VmId = (100 + n * 100 + i).ToString();
                if (i == 5) { vm.Kind = GuestKind.Container; vm.OsConfigured = "LXC: debian"; vm.Firmware = null; }
                snap.VirtualMachines.Add(vm);
            }
        }
        cluster.NumHosts = cluster.NumEffectiveHosts = 3;
        cluster.NumCpuCores = 72;
        cluster.NumCpuThreads = 144;
        cluster.TotalMemoryBytes = 3 * (512L << 30);
        return snap;
    }

    private static InventorySnapshot VMware(Random rng)
    {
        const string src = "vcenter.example.com";
        var source = new Source
        {
            Address = src, Platform = Platform.VMware, Group = "Hong Kong", ProductName = "VMware vCenter Server", Version = "8.0.3",
            Build = "24022515", ApiVersion = "8.0.3.0", Vendor = "VMware, Inc.", OsType = "linux-x64",
        };
        var cluster = new ClusterInfo { Name = "HK-Prod", Platform = Platform.VMware, SourceAddress = src, Datacenter = "HKDC1", HaEnabled = true, DrsEnabled = true, OverallStatus = "green" };
        var snap = new InventorySnapshot { Source = source, Clusters = { cluster } };
        var ds = new Datastore { Name = "vsanDatastore", Platform = Platform.VMware, SourceAddress = src, Cluster = cluster.Name, Type = "vsan", Shared = true, CapacityBytes = 30L << 40, FreeBytes = 3L << 40 };
        snap.Datastores.Add(ds);
        for (var n = 1; n <= 2; n++)
        {
            var host = new HostSystem
            {
                Name = $"hk-esxi0{n}.example.com", Platform = Platform.VMware, SourceAddress = src, Datacenter = "HKDC1", Cluster = cluster.Name,
                Status = n == 2 ? "yellow" : "green", CpuModel = "Intel(R) Xeon(R) Gold 6430", CpuMhz = 2100, CpuSockets = 2, CoresPerSocket = 32,
                CpuCores = 64, CpuThreads = 128, HyperThreadingActive = true, CpuUsagePercent = 31, MemoryBytes = 1024L << 30, MemoryUsedBytes = 700L << 30,
                Version = "VMware ESXi 8.0.3 build-24022510", BootTime = DateTimeOffset.Now.AddDays(-90), Vendor = "HPE", Model = "ProLiant DL380 Gen11",
                SerialNumber = $"HKSN0{n}", InMaintenance = false, TimeZone = "UTC", NtpServers = ["192.0.2.1"],
            };
            host.Extra["Current EVC"] = "intel-sapphirerapids";
            snap.Hosts.Add(host);
            ds.Hosts.Add(host.Name);
            cluster.Nodes.Add(new ClusterNode { Name = host.Name, State = "connected" });
            for (var i = 0; i < 8; i++) snap.VirtualMachines.Add(Vm(rng, Platform.VMware, host, cluster.Name, src, "HKDC1", ds.Name, $"HK-{Roles[(i + n * 5) % Roles.Length].ToUpperInvariant()}{n}{i:00}"));
        }
        cluster.NumHosts = cluster.NumEffectiveHosts = 2;
        cluster.NumCpuCores = 128;
        cluster.TotalMemoryBytes = 2048L << 30;
        source.Extra["__sheet:vLicense"] = new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "VMware vSphere 8 Enterprise Plus", ["Key"] = "*****-*****-*****-*****-AB12C", ["Total"] = 8L, ["Used"] = 4L, ["Cost Unit"] = "cpuPackage", ["VI SDK Server"] = src },
        };
        return snap;
    }

    private static VirtualMachine Vm(Random rng, Platform platform, HostSystem host, string cluster, string src, string dc, string datastore, string name)
    {
        var cpus = new[] { 2, 2, 4, 4, 8, 16 }[rng.Next(6)];
        var mem = new[] { 2048, 4096, 8192, 16384, 32768 }[rng.Next(5)];
        var on = rng.NextDouble() > 0.15;
        var windows = platform == Platform.HyperV || rng.NextDouble() > 0.6;
        var vm = new VirtualMachine
        {
            Name = name, VmId = Guid.NewGuid().ToString(), Uuid = Guid.NewGuid().ToString(), Platform = platform, Host = host.Name,
            Cluster = cluster, Datacenter = dc, SourceAddress = src, PowerState = on ? PowerState.PoweredOn : PowerState.PoweredOff,
            RawState = on ? "Running" : "Off", CpuCount = cpus, Sockets = 1, CoresPerSocket = cpus, MemoryMiB = mem,
            MemoryAssignedMiB = on ? mem : 0, MemoryDemandMiB = on ? mem * rng.Next(20, 80) / 100 : null, DynamicMemory = platform == Platform.HyperV && rng.NextDouble() > 0.5,
            Firmware = rng.NextDouble() > 0.3 ? "efi" : "bios", SecureBoot = rng.NextDouble() > 0.5,
            HardwareVersion = platform switch { Platform.HyperV => "10.0", Platform.Proxmox => "pc-q35-8.1", _ => "vmx-21" },
            Generation = platform == Platform.HyperV ? "Gen 2" : null,
            OsConfigured = windows ? "Windows Server 2022" : "Ubuntu Linux (64-bit)", OsGuest = on ? (windows ? "Microsoft Windows Server 2022 Datacenter" : "Ubuntu 24.04 LTS") : null,
            DnsName = on ? name.ToLowerInvariant() + ".example.com" : null, Heartbeat = on ? "green" : "gray",
            ToolsRunning = on ? rng.NextDouble() > 0.1 : null, ToolsStatus = on ? "toolsOk" : "toolsNotRunning",
            ToolsVersion = platform == Platform.VMware ? "12389" : null, CreationDate = DateTimeOffset.Now.AddDays(-rng.Next(30, 1500)),
            Uptime = on ? TimeSpan.FromHours(rng.Next(1, 2000)) : null, HaProtected = true, OwnerNode = host.Name, AutoStart = true,
            Annotation = rng.NextDouble() > 0.7 ? "Owner: Platform team" : null,
            ConfigPath = platform == Platform.HyperV ? $@"C:\ClusterStorage\Volume1\{name}" : null,
        };
        if (vm.PowerState == PowerState.PoweredOn) vm.PowerOnTime = DateTimeOffset.Now - vm.Uptime;
        var disks = rng.Next(1, 4);
        for (var d = 0; d < disks; d++)
        {
            var cap = (long)new[] { 40, 60, 100, 250, 500 }[rng.Next(5)] << 30;
            vm.Disks.Add(new VmDisk
            {
                Index = d + 1, Label = platform == Platform.Proxmox ? $"scsi{d}" : $"Hard disk {d + 1}", Controller = platform == Platform.VMware ? "SCSI controller 0" : $"SCSI 0:{d}",
                Datastore = datastore, CapacityBytes = cap, UsedBytes = (long)(cap * rng.Next(15, 90) / 100.0), Thin = true, Unit = d,
                Format = platform switch { Platform.HyperV => "vhdx", Platform.Proxmox => "raw", _ => "persistent" },
                Path = platform switch
                {
                    Platform.HyperV => $@"C:\ClusterStorage\Volume1\{name}\{name}_disk{d}.vhdx",
                    Platform.Proxmox => $"{datastore}:vm-{vm.VmId}-disk-{d}",
                    _ => $"[{datastore}] {name}/{name}_{d}.vmdk",
                },
            });
        }
        var vlan = new[] { 10, 20, 30 }[rng.Next(3)];
        vm.Nics.Add(new VmNic
        {
            Index = 1, Label = "Network adapter 1", AdapterType = platform switch { Platform.HyperV => "Synthetic", Platform.Proxmox => "virtio", _ => "vmxnet3" },
            Network = platform == Platform.Proxmox ? "vmbr0" : $"VLAN{vlan}", Vlan = vlan, Connected = on, StartsConnected = true,
            Mac = string.Join(":", Enumerable.Range(0, 6).Select(_ => rng.Next(256).ToString("x2"))),
            Ipv4 = on ? [$"198.51.100.{rng.Next(2, 250)}"] : [],
        });
        if (rng.NextDouble() > 0.8)
            vm.Snapshots.Add(new VmSnapshot { Name = "Before patching", Created = DateTimeOffset.Now.AddDays(-rng.Next(1, 40)), Type = "Standard" });
        if (rng.NextDouble() > 0.85)
            vm.CdDrives.Add(new VmCdDrive { DeviceNode = "DVD 0:1", Media = "SW_DVD9_Win_Server_2022.iso", Connected = true, DeviceType = "ISO" });
        if (on && !windows)
            vm.Partitions.Add(new VmPartition { Name = "/", FileSystem = "ext4", CapacityBytes = 40L << 30, FreeBytes = (long)rng.Next(2, 30) << 30 });
        return vm;
    }
}
