using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Tables;

// Row shapes: each pairs the inventory object with its parents so columns can reach context.
public sealed record VmRow(VirtualMachine Vm, Inventory Inv);
public sealed record VmDiskRow(VirtualMachine Vm, VmDisk Disk, Inventory Inv);
public sealed record VmNicRow(VirtualMachine Vm, VmNic Nic, Inventory Inv);
public sealed record VmCdRow(VirtualMachine Vm, VmCdDrive Cd, Inventory Inv);
public sealed record VmUsbRow(VirtualMachine Vm, VmUsbDevice Usb, Inventory Inv);
public sealed record VmSnapshotRow(VirtualMachine Vm, VmSnapshot Snapshot, Inventory Inv);
public sealed record VmPartitionRow(VirtualMachine Vm, VmPartition Partition, Inventory Inv);
public sealed record HostRow(HostSystem Host, Inventory Inv);
public sealed record HostNicRow(HostSystem Host, HostNic Nic);
public sealed record HostSwitchRow(HostSystem Host, VirtualSwitch Switch);
public sealed record HostPortRow(HostSystem Host, PortGroup Port);
public sealed record HostIpRow(HostSystem Host, HostIpInterface Ip);
public sealed record HostHbaRow(HostSystem Host, StorageAdapter Hba);
public sealed record DatastoreRow(Datastore Ds, Inventory Inv);
public sealed record ClusterRow(ClusterInfo Cluster);
public sealed record SourceRow(Source Source);
public sealed record HealthRow(HealthItem Item);
public sealed record RawSheetRow(Source Source, Dictionary<string, object?> Values);
public sealed record MetaRow(string Server, DateTime Created);

/// <summary>Table definitions for every RVTools worksheet.</summary>
public static class RvToolsTables
{
    public const string AppVersion = "3.0.0";

    private const double MiB = 1024 * 1024;

    private static long? ToMiB(long? bytes) => bytes is null ? null : (long)Math.Round(bytes.Value / MiB);

    public static string PowerStateText(PowerState s) => s switch
    {
        PowerState.PoweredOn => "poweredOn",
        PowerState.PoweredOff => "poweredOff",
        PowerState.Suspended => "suspended",
        _ => "unknown",
    };

    public static string PlatformName(Platform p) => p switch
    {
        Platform.HyperV => "Hyper-V",
        Platform.Proxmox => "Proxmox VE",
        Platform.VMware => "VMware",
        _ => p.ToString(),
    };

    private static string Datacenter(VirtualMachine vm, Inventory inv) =>
        vm.Datacenter ?? inv.FindSource(vm.SourceAddress)?.Group ?? PlatformName(vm.Platform);

    private static string Datacenter(HostSystem h, Inventory? inv) =>
        h.Datacenter ?? inv?.FindSource(h.SourceAddress)?.Group ?? PlatformName(h.Platform);

    private static string? Join(IEnumerable<string>? items) =>
        items is null ? null : string.Join(", ", items.Where(i => !string.IsNullOrWhiteSpace(i)));

    private static (string, string?, string?) Scope(VirtualMachine vm) => (vm.SourceAddress, vm.Host, vm.Cluster);
    private static (string, string?, string?) Scope(HostSystem h) => (h.SourceAddress, h.Name, h.Cluster);

    /// <summary>Columns shared by every per-VM sheet (Annotation .. VI SDK UUID).</summary>
    private static RvSheetBuilder<T> VmTrailer<T>(this RvSheetBuilder<T> b, Func<T, VirtualMachine> vm, Func<T, Inventory> inv)
        where T : class
    {
        var headers = b.Headers;
        if (headers.Contains("Annotation")) b.Text("Annotation", r => vm(r).Annotation);
        if (headers.Contains("Datacenter")) b.Text("Datacenter", r => Datacenter(vm(r), inv(r)));
        if (headers.Contains("Cluster")) b.Text("Cluster", r => vm(r).Cluster);
        if (headers.Contains("Host")) b.Text("Host", r => vm(r).Host);
        if (headers.Contains("Folder")) b.Text("Folder", r => vm(r).ResourcePool);
        if (headers.Contains("OS according to the configuration file")) b.Text("OS according to the configuration file", r => vm(r).OsConfigured);
        if (headers.Contains("OS according to the VMware Tools")) b.Text("OS according to the VMware Tools", r => vm(r).OsGuest);
        if (headers.Contains("OS according to the VMware tools")) b.Text("OS according to the VMware tools", r => vm(r).OsGuest);
        if (headers.Contains("VMRef")) b.Text("VMRef", r => vm(r).VmId);
        if (headers.Contains("VM ID")) b.Text("VM ID", r => vm(r).VmId);
        if (headers.Contains("VM UUID")) b.Text("VM UUID", r => vm(r).Uuid);
        if (headers.Contains("VI SDK Server")) b.Text("VI SDK Server", r => vm(r).SourceAddress);
        if (headers.Contains("VI SDK UUID")) b.Text("VI SDK UUID", r => vm(r).SourceAddress);
        if (headers.Contains("Powerstate")) b.Text("Powerstate", r => PowerStateText(vm(r).PowerState));
        if (headers.Contains("Template")) b.Bool("Template", r => vm(r).IsTemplate);
        if (headers.Contains("SRM Placeholder")) b.Bool("SRM Placeholder", _ => false);
        return b.Text("VM", r => vm(r).Name);
    }

    private static RvSheetBuilder<T> HostTrailer<T>(this RvSheetBuilder<T> b, Func<T, HostSystem> host)
        where T : class
    {
        var headers = b.Headers;
        if (headers.Contains("Host")) b.Text("Host", r => host(r).Name);
        if (headers.Contains("Datacenter")) b.Text("Datacenter", r => Datacenter(host(r), null));
        if (headers.Contains("Cluster")) b.Text("Cluster", r => host(r).Cluster);
        if (headers.Contains("VI SDK Server")) b.Text("VI SDK Server", r => host(r).SourceAddress);
        if (headers.Contains("VI SDK UUID")) b.Text("VI SDK UUID", r => host(r).SourceAddress);
        return b;
    }

    /// <summary>
    /// Key under <see cref="InventoryObject.Extra"/> of a <see cref="Source"/> holding pre-built rows for a whole
    /// sheet (<c>List&lt;Dictionary&lt;string, object?&gt;&gt;</c> keyed by RVTools header). Used by collectors for
    /// sheets with no model equivalent, e.g. vSphere resource pools or distributed switches.
    /// </summary>
    public static string RawSheetKey(string sheet) => "__sheet:" + sheet;

    private static TableDefinition RawSheet(string sheet, string description)
    {
        var key = RawSheetKey(sheet);
        var columns = RvToolsSchema.HeadersFor(sheet)
            .Select(h => new TableColumn(h, CellKind.Text, row =>
                ((RawSheetRow)row).Values.TryGetValue(h, out var v) ? v : null))
            .ToList();
        return new TableDefinition
        {
            Name = sheet,
            Description = description,
            Columns = columns,
            Rows = inv => inv.Sources
                .SelectMany(src => src.Extra.TryGetValue(key, out var raw) && raw is IEnumerable<Dictionary<string, object?>> rows
                    ? rows.Select(r => (object)new RawSheetRow(src, r))
                    : [])
                .ToList(),
            ScopeOf = row => (((RawSheetRow)row).Source.Address, null, null),
        };
    }

    private static IEnumerable<VmRow> Vms(Inventory inv) => inv.VirtualMachines.Select(v => new VmRow(v, inv));

    public static IReadOnlyList<TableDefinition> All { get; } = Build();

    public static TableDefinition Get(string name) => All.First(t => t.Name == name);

    private static List<TableDefinition> Build() =>
    [
        VInfo(), VCpu(), VMemory(), VDisk(), VPartition(), VNetwork(), VCd(), VUsb(), VSnapshot(), VTools(),
        VSource(), VRp(), VCluster(), VHost(), VHba(), VNic(), VSwitch(), VPort(), DvSwitch(), DvPort(),
        VScVmk(), VDatastore(), VMultiPath(), VLicense(), VFileInfo(), VHealth(), VMetaData(),
    ];

    private static TableDefinition VInfo()
    {
        var b = new RvSheetBuilder<VmRow>("vInfo").ExtrasFrom(r => r.Vm)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Config status", r => r.Vm.Status)
            .Text("DNS Name", r => r.Vm.DnsName)
            .Text("Connection state", r => r.Vm.PowerState == PowerState.Unknown ? "disconnected" : "connected")
            .Text("Guest state", r => r.Vm.PowerState switch
            {
                PowerState.PoweredOn => r.Vm.ToolsRunning == false ? "notRunning" : "running",
                PowerState.Suspended => "standby",
                _ => "notRunning",
            })
            .Text("Heartbeat", r => r.Vm.Heartbeat)
            .Bool("Consolidation Needed", _ => false)
            .Date("PowerOn", r => r.Vm.PowerOnTime)
            .Date("Creation date", r => r.Vm.CreationDate)
            .Num("CPUs", r => r.Vm.CpuCount)
            .Num("Memory", r => r.Vm.MemoryMiB)
            .Num("Active Memory", r => r.Vm.MemoryDemandMiB)
            .Num("NICs", r => r.Vm.Nics.Count)
            .Num("Disks", r => r.Vm.Disks.Count)
            .Num("Total disk capacity MiB", r => ToMiB(r.Vm.ProvisionedBytes))
            .Text("Primary IP Address", r => r.Vm.PrimaryIp)
            .Text("Resource pool", r => r.Vm.ResourcePool)
            .Num("Provisioned MiB", r => ToMiB(r.Vm.ProvisionedBytes))
            .Num("In Use MiB", r => r.Vm.UsedBytes > 0 ? ToMiB(r.Vm.UsedBytes) : null)
            .Num("Unshared MiB", r => r.Vm.UsedBytes > 0 ? ToMiB(r.Vm.UsedBytes) : null)
            .Text("HA Restart Priority", r => r.Vm.HaProtected == true ? (r.Vm.HaState ?? "enabled") : null)
            .Num("Boot delay", r => r.Vm.StartDelaySeconds is { } s ? s * 1000 : null)
            .Bool("EFI Secure boot", r => r.Vm.SecureBoot)
            .Text("Firmware", r => r.Vm.Firmware)
            .Text("HW version", r => r.Vm.HardwareVersion)
            .Text("Path", r => r.Vm.ConfigPath)
            .Text("Snapshot directory", r => r.Vm.SnapshotDirectory)
            .Text("SMBIOS UUID", r => r.Vm.BiosUuid)
            .Text("VI SDK Server type", r => r.Vm.SourceProduct)
            .Text("VI SDK API Version", r => r.Vm.SourceApiVersion);
        for (var i = 1; i <= 8; i++)
        {
            var idx = i - 1;
            b.Text($"Network #{i}", r => idx < r.Vm.Nics.Count ? r.Vm.Nics[idx].Network : null);
        }
        return b.Build(Vms, "One row per VM / container: power, sizing, network, storage and placement.",
            r => r.Vm, r => Scope(r.Vm));
    }

    private static TableDefinition VCpu() =>
        new RvSheetBuilder<VmRow>("vCPU").ExtrasFrom(r => r.Vm)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Num("CPUs", r => r.Vm.CpuCount)
            .Num("Sockets", r => r.Vm.Sockets)
            .Num("Cores p/s", r => r.Vm.CoresPerSocket)
            .Num("Shares", r => r.Vm.CpuShares)
            .Num("Reservation", r => r.Vm.CpuReservation)
            .Num("Limit", r => r.Vm.CpuLimit)
            .Bool("Hot Add", r => r.Vm.CpuHotAdd)
            .Build(Vms, "Virtual CPU configuration per VM.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VMemory() =>
        new RvSheetBuilder<VmRow>("vMemory").ExtrasFrom(r => r.Vm)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Num("Size MiB", r => r.Vm.MemoryMiB)
            .Num("Max", r => r.Vm.MemoryMaxMiB)
            .Num("Consumed", r => r.Vm.MemoryAssignedMiB)
            .Num("Ballooned", r => r.Vm.MemoryBalloonedMiB)
            .Num("Active", r => r.Vm.MemoryDemandMiB)
            .Num("Reservation", r => r.Vm.MemoryMinMiB)
            .Bool("Hot Add", r => r.Vm.DynamicMemory)
            .Build(Vms, "Memory configuration and usage per VM.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VDisk() =>
        new RvSheetBuilder<VmDiskRow>("vDisk").ExtrasFrom(r => r.Disk)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Disk", r => r.Disk.Label)
            .Text("Disk Path", r => r.Disk.Path)
            .Num("Capacity MiB", r => ToMiB(r.Disk.CapacityBytes))
            .Bool("Raw", r => r.Disk.Passthrough ?? false)
            .Text("Disk Mode", r => r.Disk.Format)
            .Bool("Sharing mode", r => r.Disk.Shared)
            .Bool("Thin", r => r.Disk.Thin)
            .Text("Write Through", r => r.Disk.Cache)
            .Text("Controller", r => r.Disk.Controller ?? r.Disk.ControllerType)
            .Text("Label", r => r.Disk.Label)
            .Num("Unit #", r => r.Disk.Unit)
            .Text("Path", r => r.Disk.ParentPath)
            .Num("Internal Sort Column", r => r.Disk.Index)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.Disks.Select(d => new VmDiskRow(v, d, inv))),
                "One row per virtual disk.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VPartition() =>
        new RvSheetBuilder<VmPartitionRow>("vPartition").ExtrasFrom(r => r.Partition)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Disk", r => r.Partition.Name)
            .Num("Capacity MiB", r => ToMiB(r.Partition.CapacityBytes))
            .Num("Consumed MiB", r => r.Partition.CapacityBytes is { } c && r.Partition.FreeBytes is { } f ? ToMiB(c - f) : null)
            .Num("Free MiB", r => ToMiB(r.Partition.FreeBytes))
            .Num("Free %", r => r.Partition.CapacityBytes is > 0 && r.Partition.FreeBytes is { } f
                ? (long)Math.Round(f * 100.0 / r.Partition.CapacityBytes.Value) : null)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.Partitions.Select(p => new VmPartitionRow(v, p, inv))),
                "Guest file systems (requires guest tools/agent).", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VNetwork() =>
        new RvSheetBuilder<VmNicRow>("vNetwork").ExtrasFrom(r => r.Nic)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("NIC label", r => r.Nic.Label)
            .Text("Adapter", r => r.Nic.AdapterType)
            .Text("Network", r => r.Nic.Vlan is { } v && v > 0 ? $"{r.Nic.Network} (VLAN {v})" : r.Nic.Network)
            .Text("Switch", r => r.Nic.Switch)
            .Bool("Connected", r => r.Nic.Connected)
            .Bool("Starts Connected", r => r.Nic.StartsConnected)
            .Text("Mac Address", r => r.Nic.Mac)
            .Text("Type", r => r.Nic.AdapterType)
            .Text("IPv4 Address", r => Join(r.Nic.Ipv4))
            .Text("IPv6 Address", r => Join(r.Nic.Ipv6))
            .Num("Internal Sort Column", r => r.Nic.Index)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.Nics.Select(n => new VmNicRow(v, n, inv))),
                "One row per virtual network adapter.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VCd() =>
        new RvSheetBuilder<VmCdRow>("vCD").ExtrasFrom(r => r.Cd)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Device Node", r => r.Cd.DeviceNode)
            .Bool("Connected", r => r.Cd.Connected)
            .Text("Device Type", r => r.Cd.DeviceType ?? r.Cd.Media)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.CdDrives.Select(c => new VmCdRow(v, c, inv))),
                "Virtual CD/DVD drives and mounted media.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VUsb() =>
        new RvSheetBuilder<VmUsbRow>("vUSB").ExtrasFrom(r => r.Usb)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Device Node", r => r.Usb.DeviceNode)
            .Text("Device Type", r => r.Usb.DeviceType)
            .Bool("Connected", r => r.Usb.Connected)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.UsbDevices.Select(u => new VmUsbRow(v, u, inv))),
                "USB devices passed through to VMs.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VSnapshot() =>
        new RvSheetBuilder<VmSnapshotRow>("vSnapshot").ExtrasFrom(r => r.Snapshot)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("Name", r => r.Snapshot.Name)
            .Text("Description", r => r.Snapshot.Description)
            .Date("Date / time", r => r.Snapshot.Created)
            .Text("Filename", r => r.Snapshot.Path)
            .Num("Size MiB (total)", r => ToMiB(r.Snapshot.SizeBytes))
            .Text("State", r => r.Snapshot.IncludesMemory == true ? "poweredOn" : r.Snapshot.Type)
            .Build(inv => inv.VirtualMachines.SelectMany(v => v.Snapshots.Select(s => new VmSnapshotRow(v, s, inv))),
                "Snapshots / checkpoints.", r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VTools() =>
        new RvSheetBuilder<VmRow>("vTools").ExtrasFrom(r => r.Vm)
            .VmTrailer(r => r.Vm, r => r.Inv)
            .Text("VM Version", r => r.Vm.HardwareVersion)
            .Text("Tools", r => r.Vm.ToolsStatus)
            .Text("Tools Version", r => r.Vm.ToolsVersion)
            .Text("Heartbeat status", r => r.Vm.Heartbeat)
            .Build(Vms, "Guest tools: Hyper-V integration services, QEMU guest agent or VMware Tools.",
                r => r.Vm, r => Scope(r.Vm));

    private static TableDefinition VSource() =>
        new RvSheetBuilder<SourceRow>("vSource").ExtrasFrom(r => r.Source)
            .Text("Name", r => r.Source.ProductName)
            .Text("OS type", r => r.Source.OsType)
            .Text("API type", r => r.Source.Platform switch
            {
                Platform.HyperV => "WinRM",
                Platform.Proxmox => "PVE REST",
                _ => "VI SDK",
            })
            .Text("API version", r => r.Source.ApiVersion)
            .Text("Version", r => r.Source.Version)
            .Text("Build", r => r.Source.Build)
            .Text("Fullname", r => r.Source.FullName)
            .Text("Product name", r => r.Source.ProductName)
            .Text("Product version", r => r.Source.Version)
            .Text("Vendor", r => r.Source.Vendor)
            .Text("VI SDK Server", r => r.Source.Address)
            .Text("VI SDK UUID", r => r.Source.Address)
            .Build(inv => inv.Sources.Select(s => new SourceRow(s)), "Connected sources.",
                scopeOf: r => (r.Source.Address, null, null));

    private static TableDefinition VRp() =>
        RawSheet("vRP", "Resource pools (vCenter / ESXi).");

    private static TableDefinition VCluster() =>
        new RvSheetBuilder<ClusterRow>("vCluster").ExtrasFrom(r => r.Cluster)
            .Text("Name", r => r.Cluster.Name)
            .Text("OverallStatus", r => r.Cluster.OverallStatus ?? (r.Cluster.Quorate switch
            {
                true => "green",
                false => "red",
                null => null,
            }))
            .Num("NumHosts", r => r.Cluster.NumHosts)
            .Num("numEffectiveHosts", r => r.Cluster.NumEffectiveHosts)
            .Num("TotalCpu", r => r.Cluster.TotalCpuMhz)
            .Num("NumCpuCores", r => r.Cluster.NumCpuCores)
            .Num("NumCpuThreads", r => r.Cluster.NumCpuThreads)
            .Num("TotalMemory", r => r.Cluster.TotalMemoryBytes)
            .Bool("HA enabled", r => r.Cluster.HaEnabled)
            .Text("Object ID", r => r.Cluster.Id)
            .Text("VI SDK Server", r => r.Cluster.SourceAddress)
            .Text("VI SDK UUID", r => r.Cluster.SourceAddress)
            .Build(inv => inv.Clusters.Select(c => new ClusterRow(c)), "Clusters (PVE, Failover Cluster, vSphere).",
                scopeOf: r => (r.Cluster.SourceAddress, null, r.Cluster.Name));

    private static TableDefinition VHost() =>
        new RvSheetBuilder<HostRow>("vHost").ExtrasFrom(r => r.Host)
            .HostTrailer(r => r.Host)
            .Text("Datacenter", r => Datacenter(r.Host, r.Inv))
            .Text("Config status", r => r.Host.Status)
            .Bool("in Maintenance Mode", r => r.Host.InMaintenance)
            .Text("CPU Model", r => r.Host.CpuModel)
            .Num("Speed", r => r.Host.CpuMhz)
            .Bool("HT Available", r => r.Host.CpuThreads is { } t && r.Host.CpuCores is { } c ? t > c : null)
            .Bool("HT Active", r => r.Host.HyperThreadingActive)
            .Num("# CPU", r => r.Host.CpuSockets)
            .Num("Cores per CPU", r => r.Host.CoresPerSocket)
            .Num("# Cores", r => r.Host.CpuCores)
            .Num("CPU usage %", r => r.Host.CpuUsagePercent is { } p ? Math.Round(p) : null)
            .Num("# Memory", r => ToMiB(r.Host.MemoryBytes))
            .Num("Memory usage %", r => r.Host.MemoryBytes is > 0 && r.Host.MemoryUsedBytes is { } u
                ? Math.Round(u * 100.0 / r.Host.MemoryBytes.Value) : null)
            .Num("# NICs", r => r.Host.Nics.Count)
            .Num("# HBAs", r => r.Host.StorageAdapters.Count)
            .Num("# VMs total", r => HostVms(r).Count())
            .Num("# VMs", r => HostVms(r).Count(v => v.PowerState == PowerState.PoweredOn))
            .Num("VMs per Core", r => r.Host.CpuCores is > 0 ? Math.Round(HostVms(r).Count(v => v.PowerState == PowerState.PoweredOn) / (double)r.Host.CpuCores.Value, 2) : null)
            .Num("# vCPUs", r => HostVms(r).Where(v => v.PowerState == PowerState.PoweredOn).Sum(v => v.CpuCount))
            .Num("vCPUs per Core", r => r.Host.CpuCores is > 0 ? Math.Round(HostVms(r).Where(v => v.PowerState == PowerState.PoweredOn).Sum(v => v.CpuCount) / (double)r.Host.CpuCores.Value, 2) : null)
            .Num("vRAM", r => HostVms(r).Where(v => v.PowerState == PowerState.PoweredOn).Sum(v => v.MemoryMiB))
            .Num("VM Used memory", r => HostVms(r).Sum(v => v.MemoryAssignedMiB ?? 0) is var s && s > 0 ? s : null)
            .Num("VM Memory Ballooned", r => HostVms(r).Sum(v => v.MemoryBalloonedMiB ?? 0))
            .Text("ESX Version", r => r.Host.Version)
            .Date("Boot time", r => r.Host.BootTime)
            .Text("DNS Servers", r => Join(r.Host.DnsServers))
            .Text("Domain", r => r.Host.Domain)
            .Text("DNS Search Order", r => Join(r.Host.DnsSearch))
            .Text("NTP Server(s)", r => Join(r.Host.NtpServers))
            .Text("Time Zone", r => r.Host.TimeZone)
            .Text("Time Zone Name", r => r.Host.TimeZone)
            .Text("Vendor", r => r.Host.Vendor)
            .Text("Model", r => r.Host.Model)
            .Text("Serial number", r => r.Host.SerialNumber)
            .Text("Service tag", r => r.Host.SerialNumber)
            .Text("BIOS Vendor", r => r.Host.BiosVendor)
            .Text("BIOS Version", r => r.Host.BiosVersion)
            .Date("BIOS Date", r => r.Host.BiosDate)
            .Text("UUID", r => r.Host.Uuid)
            .Build(inv => inv.Hosts.Select(h => new HostRow(h, inv)), "Physical hosts / nodes.",
                scopeOf: r => Scope(r.Host));

    private static IEnumerable<VirtualMachine> HostVms(HostRow r) =>
        r.Inv.VirtualMachines.Where(v => v.SourceAddress == r.Host.SourceAddress && v.Host == r.Host.Name);

    private static TableDefinition VHba() =>
        new RvSheetBuilder<HostHbaRow>("vHBA").ExtrasFrom(r => r.Hba)
            .HostTrailer(r => r.Host)
            .Text("Device", r => r.Hba.Device)
            .Text("Type", r => r.Hba.Type)
            .Text("Status", r => r.Hba.Status)
            .Text("Pci", r => r.Hba.Pci)
            .Text("Driver", r => r.Hba.Driver)
            .Text("Model", r => r.Hba.Model)
            .Text("WWN", r => r.Hba.Wwn)
            .Build(inv => inv.Hosts.SelectMany(h => h.StorageAdapters.Select(a => new HostHbaRow(h, a))),
                "Host storage adapters.", scopeOf: r => Scope(r.Host));

    private static TableDefinition VNic() =>
        new RvSheetBuilder<HostNicRow>("vNIC").ExtrasFrom(r => r.Nic)
            .HostTrailer(r => r.Host)
            .Text("Network Device", r => r.Nic.Name)
            .Text("Driver", r => r.Nic.Driver ?? r.Nic.Description)
            .Num("Speed", r => r.Nic.SpeedMbps)
            .Text("Duplex", r => r.Nic.FullDuplex switch { true => "Full Duplex", false => "Half Duplex", _ => null })
            .Text("MAC", r => r.Nic.Mac)
            .Text("Switch", r => r.Nic.Switch)
            .Text("PCI", r => r.Nic.Pci)
            .Build(inv => inv.Hosts.SelectMany(h => h.Nics.Select(n => new HostNicRow(h, n))),
                "Physical host network adapters.", scopeOf: r => Scope(r.Host));

    private static TableDefinition VSwitch() =>
        new RvSheetBuilder<HostSwitchRow>("vSwitch").ExtrasFrom(r => r.Switch)
            .HostTrailer(r => r.Host)
            .Text("Switch", r => r.Switch.Name)
            .Num("# Ports", r => r.Switch.Ports)
            .Text("Policy", r => r.Switch.Type)
            .Text("Rolling Order", r => Join(r.Switch.Uplinks))
            .Num("MTU", r => r.Switch.Mtu)
            .Build(inv => inv.Hosts.SelectMany(h => h.Switches.Select(s => new HostSwitchRow(h, s))),
                "Virtual switches / Linux bridges. Uplinks are shown in 'Rolling Order'.", scopeOf: r => Scope(r.Host));

    private static TableDefinition VPort() =>
        new RvSheetBuilder<HostPortRow>("vPort").ExtrasFrom(r => r.Port)
            .HostTrailer(r => r.Host)
            .Text("Port Group", r => r.Port.Name)
            .Text("Switch", r => r.Port.Switch)
            .Num("VLAN", r => r.Port.Vlan)
            .Build(inv => inv.Hosts.SelectMany(h => h.PortGroups.Select(p => new HostPortRow(h, p))),
                "Port groups / VLAN interfaces.", scopeOf: r => Scope(r.Host));

    private static TableDefinition DvSwitch() =>
        RawSheet("dvSwitch", "Distributed switches (vCenter).");

    private static TableDefinition DvPort() =>
        RawSheet("dvPort", "Distributed port groups (vCenter).");

    private static TableDefinition VScVmk() =>
        new RvSheetBuilder<HostIpRow>("vSC_VMK").ExtrasFrom(r => r.Ip)
            .HostTrailer(r => r.Host)
            .Text("Port Group", r => r.Ip.PortGroup)
            .Text("Device", r => r.Ip.Name)
            .Text("Mac Address", r => r.Ip.Mac)
            .Bool("DHCP", r => r.Ip.Dhcp)
            .Text("IP Address", r => r.Ip.Ipv4)
            .Text("IP 6 Address", r => r.Ip.Ipv6)
            .Text("Subnet mask", r => r.Ip.SubnetMask)
            .Text("Gateway", r => r.Ip.Gateway)
            .Num("MTU", r => r.Ip.Mtu)
            .Build(inv => inv.Hosts.SelectMany(h => h.IpInterfaces.Select(i => new HostIpRow(h, i))),
                "Host management IP interfaces.", scopeOf: r => Scope(r.Host));

    private static TableDefinition VDatastore() =>
        new RvSheetBuilder<DatastoreRow>("vDatastore").ExtrasFrom(r => r.Ds)
            .Text("Name", r => r.Ds.Name)
            .Text("Address", r => r.Ds.Address)
            .Bool("Accessible", r => r.Ds.Accessible)
            .Text("Type", r => r.Ds.Type)
            .Num("# VMs total", r => DsVms(r).Count())
            .Num("# VMs", r => DsVms(r).Count(v => v.PowerState == PowerState.PoweredOn))
            .Num("Capacity MiB", r => ToMiB(r.Ds.CapacityBytes))
            .Num("Provisioned MiB", r => ToMiB(DsVms(r).SelectMany(v => v.Disks)
                .Where(d => DiskOn(d, r.Ds)).Sum(d => d.CapacityBytes ?? 0)))
            .Num("In Use MiB", r => r.Ds.CapacityBytes is { } c && r.Ds.FreeBytes is { } f ? ToMiB(c - f) : null)
            .Num("Free MiB", r => ToMiB(r.Ds.FreeBytes))
            .Num("Free %", r => r.Ds.CapacityBytes is > 0 && r.Ds.FreeBytes is { } f
                ? Math.Round(f * 100.0 / r.Ds.CapacityBytes.Value) : null)
            .Num("# Hosts", r => r.Ds.Hosts.Count)
            .Text("Hosts", r => Join(r.Ds.Hosts))
            .Text("Cluster name", r => r.Ds.Cluster)
            .Text("URL", r => r.Ds.Content)
            .Text("VI SDK Server", r => r.Ds.SourceAddress)
            .Text("VI SDK UUID", r => r.Ds.SourceAddress)
            .Build(inv => inv.Datastores.Select(d => new DatastoreRow(d, inv)),
                "Datastores / storage pools / volumes.", scopeOf: r => (r.Ds.SourceAddress, null, r.Ds.Cluster));

    private static bool DiskOn(VmDisk d, Datastore ds) =>
        string.Equals(d.Datastore, ds.Name, StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<VirtualMachine> DsVms(DatastoreRow r) =>
        r.Inv.VirtualMachines.Where(v => v.SourceAddress == r.Ds.SourceAddress
            && (r.Ds.Hosts.Count == 0 || r.Ds.Hosts.Contains(v.Host, StringComparer.OrdinalIgnoreCase))
            && v.Disks.Any(d => DiskOn(d, r.Ds)));

    private static TableDefinition VMultiPath() =>
        RawSheet("vMultiPath", "Storage multipathing (ESXi).");

    private static TableDefinition VLicense() =>
        RawSheet("vLicense", "Licenses (vCenter / ESXi).");

    private static TableDefinition VFileInfo() =>
        RawSheet("vFileInfo", "Datastore file listing.");

    private static TableDefinition VHealth() =>
        new RvSheetBuilder<HealthRow>("vHealth")
            .Text("Name", r => r.Item.Name)
            .Text("Message", r => r.Item.Message)
            .Text("Message type", r => r.Item.Severity.ToString())
            .Text("VI SDK Server", r => r.Item.SourceAddress)
            .Text("VI SDK UUID", r => r.Item.SourceAddress)
            .Build(inv => inv.Health.Select(h => new HealthRow(h)), "Health checks and collection warnings.",
                scopeOf: r => (r.Item.SourceAddress, null, null));

    private static TableDefinition VMetaData()
    {
        var t = new RvSheetBuilder<MetaRow>("vMetaData")
            .Text("RVTools major version", _ => "Hypervisor Explorer " + AppVersion)
            .Text("RVTools version", _ => AppVersion)
            .Date("xlsx creation datetime", r => r.Created)
            .Text("Server", r => r.Server)
            .Build(inv => [new MetaRow(string.Join(", ", inv.Sources.Select(s => s.Address)), DateTime.Now)],
                "Export metadata.");
        return new TableDefinition
        {
            Name = t.Name, Description = t.Description, Columns = t.Columns, Rows = t.Rows,
            FrozenColumns = 0, AutoFilter = false,
        };
    }
}
