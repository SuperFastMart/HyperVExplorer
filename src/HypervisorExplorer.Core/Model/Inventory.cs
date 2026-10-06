namespace HypervisorExplorer.Core.Model;

/// <summary>Hypervisor platforms the collectors understand.</summary>
public enum Platform
{
    HyperV,
    Proxmox,
    /// <summary>VMware ESXi host or vCenter, via the vSphere Web Services (SOAP) API.</summary>
    VMware,
}

public enum PowerState
{
    Unknown,
    PoweredOn,
    PoweredOff,
    Suspended,
}

public enum GuestKind
{
    VirtualMachine,
    Container,
}

public enum HealthSeverity
{
    Info,
    Warning,
    Error,
}

/// <summary>
/// Base for inventory objects. <see cref="Extra"/> carries platform-specific values keyed by the exact
/// RVTools column header (e.g. "EVC Mode" from vSphere); table definitions prefer these over derived values.
/// </summary>
public abstract class InventoryObject
{
    public Dictionary<string, object?> Extra { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>
/// A connection endpoint that inventory was collected from (a Hyper-V host or a PVE node/cluster address).
/// Maps to the RVTools "VI SDK Server" / vSource concepts.
/// </summary>
public sealed class Source : InventoryObject
{
    public string Address { get; set; } = "";
    public Platform Platform { get; set; }
    /// <summary>Optional site/group label (from host groups). Used as the RVTools "Datacenter".</summary>
    public string? Group { get; set; }
    public string ProductName { get; set; } = "";
    public string Version { get; set; } = "";
    public string? Build { get; set; }
    public string? ApiVersion { get; set; }
    public string Vendor { get; set; } = "";
    public string? OsType { get; set; }
    public DateTimeOffset CollectedAt { get; set; } = DateTimeOffset.Now;

    public string FullName => string.IsNullOrEmpty(Build) ? $"{ProductName} {Version}".Trim() : $"{ProductName} {Version} build-{Build}".Trim();
}

public sealed class ClusterInfo : InventoryObject
{
    public string Name { get; set; } = "";
    public Platform Platform { get; set; }
    public string SourceAddress { get; set; } = "";
    public string? Datacenter { get; set; }
    public string? OverallStatus { get; set; }
    public bool? Quorate { get; set; }
    public bool? HaEnabled { get; set; }
    public int NumHosts { get; set; }
    public int NumEffectiveHosts { get; set; }
    public int NumCpuCores { get; set; }
    public int NumCpuThreads { get; set; }
    public long TotalCpuMhz { get; set; }
    public long TotalMemoryBytes { get; set; }
    public string? Id { get; set; }
    public string? Domain { get; set; }
    public string? QuorumType { get; set; }
    public string? QuorumWitness { get; set; }
    public string? FunctionalLevel { get; set; }
    public bool? DrsEnabled { get; set; }
    public List<ClusterNode> Nodes { get; set; } = [];
    public List<ClusterNetwork> Networks { get; set; } = [];
}

public sealed class ClusterNode : InventoryObject
{
    public string Name { get; set; } = "";
    public string? State { get; set; }
    /// <summary>Failover Cluster drain status / PVE online flag, as reported.</summary>
    public string? DrainStatus { get; set; }
    public string? Id { get; set; }
    public string? Address { get; set; }
    public int? Votes { get; set; }
}

public sealed class ClusterNetwork : InventoryObject
{
    public string Name { get; set; } = "";
    public string? Role { get; set; }
    public string? Address { get; set; }
    public string? State { get; set; }
    public string? Metric { get; set; }
}

public sealed class HostSystem : InventoryObject
{
    public string Name { get; set; } = "";
    public Platform Platform { get; set; }
    public string SourceAddress { get; set; } = "";
    public string? Datacenter { get; set; }
    public string? Cluster { get; set; }
    public string? Status { get; set; }
    public bool InMaintenance { get; set; }

    public string? CpuModel { get; set; }
    public int? CpuMhz { get; set; }
    public int? CpuSockets { get; set; }
    public int? CoresPerSocket { get; set; }
    public int? CpuCores { get; set; }
    public int? CpuThreads { get; set; }
    public bool? HyperThreadingActive { get; set; }
    public double? CpuUsagePercent { get; set; }

    public long? MemoryBytes { get; set; }
    public long? MemoryUsedBytes { get; set; }

    public string? Version { get; set; }
    public string? KernelVersion { get; set; }
    public DateTimeOffset? BootTime { get; set; }

    public string? Vendor { get; set; }
    public string? Model { get; set; }
    public string? SerialNumber { get; set; }
    public string? BiosVendor { get; set; }
    public string? BiosVersion { get; set; }
    public DateTimeOffset? BiosDate { get; set; }
    public string? Uuid { get; set; }

    public List<string> DnsServers { get; set; } = [];
    public string? Domain { get; set; }
    public List<string> DnsSearch { get; set; } = [];
    public List<string> NtpServers { get; set; } = [];
    public string? TimeZone { get; set; }

    public List<HostNic> Nics { get; set; } = [];
    public List<VirtualSwitch> Switches { get; set; } = [];
    public List<PortGroup> PortGroups { get; set; } = [];
    public List<HostIpInterface> IpInterfaces { get; set; } = [];
    public List<StorageAdapter> StorageAdapters { get; set; } = [];
}

public sealed class HostNic : InventoryObject
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public string? Driver { get; set; }
    public long? SpeedMbps { get; set; }
    public bool? FullDuplex { get; set; }
    public string? Mac { get; set; }
    public string? Switch { get; set; }
    public string? Pci { get; set; }
    public string? Status { get; set; }
}

public sealed class VirtualSwitch : InventoryObject
{
    public string Name { get; set; } = "";
    public string? Type { get; set; }
    public List<string> Uplinks { get; set; } = [];
    public int? Mtu { get; set; }
    public int? Ports { get; set; }
    public string? Notes { get; set; }
}

public sealed class PortGroup : InventoryObject
{
    public string Name { get; set; } = "";
    public string? Switch { get; set; }
    public int? Vlan { get; set; }
}

/// <summary>Host-level IP interface (management OS vNIC, Linux bridge IP). RVTools vSC_VMK.</summary>
public sealed class HostIpInterface : InventoryObject
{
    public string Name { get; set; } = "";
    public string? PortGroup { get; set; }
    public string? Mac { get; set; }
    public bool? Dhcp { get; set; }
    public string? Ipv4 { get; set; }
    public string? SubnetMask { get; set; }
    public string? Gateway { get; set; }
    public string? Ipv6 { get; set; }
    public int? Mtu { get; set; }
}

public sealed class StorageAdapter : InventoryObject
{
    public string Device { get; set; } = "";
    public string? Type { get; set; }
    public string? Status { get; set; }
    public string? Driver { get; set; }
    public string? Model { get; set; }
    public string? Wwn { get; set; }
    public string? Pci { get; set; }
}

public sealed class VirtualMachine : InventoryObject
{
    public string Name { get; set; } = "";
    /// <summary>Platform VM identifier: Hyper-V VM GUID or PVE VMID.</summary>
    public string? VmId { get; set; }
    public string? Uuid { get; set; }
    public string? BiosUuid { get; set; }
    public GuestKind Kind { get; set; }
    public bool IsTemplate { get; set; }
    public Platform Platform { get; set; }
    public string Host { get; set; } = "";
    public string? Cluster { get; set; }
    public string? Datacenter { get; set; }
    public string SourceAddress { get; set; } = "";
    public string? SourceProduct { get; set; }
    public string? SourceApiVersion { get; set; }
    public string? ResourcePool { get; set; }
    public List<string> Tags { get; set; } = [];

    public PowerState PowerState { get; set; }
    /// <summary>Raw platform state (e.g. "Running", "stopped") for display.</summary>
    public string? RawState { get; set; }
    public string? Status { get; set; }
    public string? Heartbeat { get; set; }
    public string? DnsName { get; set; }
    public DateTimeOffset? CreationDate { get; set; }
    public DateTimeOffset? PowerOnTime { get; set; }
    public TimeSpan? Uptime { get; set; }

    public int CpuCount { get; set; }
    public int? Sockets { get; set; }
    public int? CoresPerSocket { get; set; }
    public int? CpuLimit { get; set; }
    public int? CpuReservation { get; set; }
    public int? CpuShares { get; set; }
    public bool? CpuHotAdd { get; set; }
    public string? CpuType { get; set; }

    /// <summary>Configured memory (startup / max).</summary>
    public long MemoryMiB { get; set; }
    public long? MemoryAssignedMiB { get; set; }
    public long? MemoryDemandMiB { get; set; }
    public long? MemoryMinMiB { get; set; }
    public long? MemoryMaxMiB { get; set; }
    public bool? DynamicMemory { get; set; }
    public long? MemoryBalloonedMiB { get; set; }

    public string? Firmware { get; set; }
    public bool? SecureBoot { get; set; }
    public string? HardwareVersion { get; set; }
    public string? Generation { get; set; }
    public string? OsConfigured { get; set; }
    public string? OsGuest { get; set; }
    public string? Annotation { get; set; }
    public string? ConfigPath { get; set; }
    public string? SnapshotDirectory { get; set; }
    public string? BootOrder { get; set; }
    public bool? AutoStart { get; set; }
    public int? StartDelaySeconds { get; set; }
    public bool? HaProtected { get; set; }
    public string? HaState { get; set; }
    /// <summary>Failover Cluster owner node / HA group.</summary>
    public string? OwnerNode { get; set; }
    public List<string> PreferredOwners { get; set; } = [];
    public string? FailoverPriority { get; set; }

    public string? ToolsStatus { get; set; }
    public string? ToolsVersion { get; set; }
    public bool? ToolsRunning { get; set; }
    public List<IntegrationComponent> IntegrationComponents { get; set; } = [];

    public List<VmDisk> Disks { get; set; } = [];
    public List<VmNic> Nics { get; set; } = [];
    public List<VmCdDrive> CdDrives { get; set; } = [];
    public List<VmSnapshot> Snapshots { get; set; } = [];
    public List<VmPartition> Partitions { get; set; } = [];
    public List<VmUsbDevice> UsbDevices { get; set; } = [];

    public long ProvisionedBytes => Disks.Sum(d => d.CapacityBytes ?? 0);
    public long UsedBytes => Disks.Sum(d => d.UsedBytes ?? 0);

    public string? PrimaryIp =>
        Nics.SelectMany(n => n.Ipv4).FirstOrDefault(ip => !string.IsNullOrWhiteSpace(ip));
}

public sealed class IntegrationComponent
{
    public string Name { get; set; } = "";
    public bool Enabled { get; set; }
    public string? Status { get; set; }
}

public sealed class VmDisk : InventoryObject
{
    public int Index { get; set; }
    public string Label { get; set; } = "";
    /// <summary>Controller slot, e.g. "SCSI 0:1" or "scsi0".</summary>
    public string? Controller { get; set; }
    public string? ControllerType { get; set; }
    public int? ControllerNumber { get; set; }
    public int? Unit { get; set; }
    public string? Path { get; set; }
    public string? Datastore { get; set; }
    public long? CapacityBytes { get; set; }
    public long? UsedBytes { get; set; }
    public string? Format { get; set; }
    public bool? Thin { get; set; }
    public bool? Shared { get; set; }
    public string? ParentPath { get; set; }
    public string? Cache { get; set; }
    public bool? Passthrough { get; set; }
    public string? Options { get; set; }
}

public sealed class VmNic : InventoryObject
{
    public int Index { get; set; }
    public string Label { get; set; } = "";
    public string? AdapterType { get; set; }
    public string? Network { get; set; }
    public string? Switch { get; set; }
    public int? Vlan { get; set; }
    public string? Mac { get; set; }
    public bool? Connected { get; set; }
    public bool? StartsConnected { get; set; }
    public List<string> Ipv4 { get; set; } = [];
    public List<string> Ipv6 { get; set; } = [];
    public bool? Firewall { get; set; }
    public string? Options { get; set; }
}

public sealed class VmCdDrive : InventoryObject
{
    public string DeviceNode { get; set; } = "";
    public string? Media { get; set; }
    public bool? Connected { get; set; }
    public string? DeviceType { get; set; }
}

public sealed class VmUsbDevice : InventoryObject
{
    public string DeviceNode { get; set; } = "";
    public string? DeviceType { get; set; }
    public bool? Connected { get; set; }
}

public sealed class VmSnapshot : InventoryObject
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public DateTimeOffset? Created { get; set; }
    public string? Parent { get; set; }
    public string? Type { get; set; }
    public bool? IncludesMemory { get; set; }
    public long? SizeBytes { get; set; }
    public string? Path { get; set; }
}

public sealed class VmPartition : InventoryObject
{
    public string Name { get; set; } = "";
    public string? FileSystem { get; set; }
    public long? CapacityBytes { get; set; }
    public long? FreeBytes { get; set; }
}

public sealed class Datastore : InventoryObject
{
    public string Name { get; set; } = "";
    public Platform Platform { get; set; }
    public string SourceAddress { get; set; } = "";
    public string? Cluster { get; set; }
    public string? Type { get; set; }
    public string? Address { get; set; }
    public bool Accessible { get; set; } = true;
    public bool? Shared { get; set; }
    public long? CapacityBytes { get; set; }
    public long? FreeBytes { get; set; }
    public string? Content { get; set; }
    public List<string> Hosts { get; set; } = [];
}

public sealed class HealthItem
{
    public string Name { get; set; } = "";
    public string Message { get; set; } = "";
    public HealthSeverity Severity { get; set; }
    public string SourceAddress { get; set; } = "";
}

/// <summary>Result of collecting one source; the unit merged into an <see cref="Inventory"/>.</summary>
public sealed class InventorySnapshot
{
    public required Source Source { get; init; }
    public List<ClusterInfo> Clusters { get; init; } = [];
    public List<HostSystem> Hosts { get; init; } = [];
    public List<VirtualMachine> VirtualMachines { get; init; } = [];
    public List<Datastore> Datastores { get; init; } = [];
    public List<HealthItem> Health { get; init; } = [];
    /// <summary>Non-fatal collection warnings (a VM whose config could not be read, etc.).</summary>
    public List<string> Warnings { get; init; } = [];

    /// <summary>
    /// Renames the source everywhere it is referenced. Used to key sources by "address:port" when a
    /// non-default port is used, so two endpoints on one address do not overwrite each other.
    /// </summary>
    public void RekeySource(string address)
    {
        var old = Source.Address;
        if (string.Equals(old, address, StringComparison.OrdinalIgnoreCase)) return;
        Source.Address = address;
        foreach (var c in Clusters) if (c.SourceAddress == old) c.SourceAddress = address;
        foreach (var h in Hosts) if (h.SourceAddress == old) h.SourceAddress = address;
        foreach (var v in VirtualMachines) if (v.SourceAddress == old) v.SourceAddress = address;
        foreach (var d in Datastores) if (d.SourceAddress == old) d.SourceAddress = address;
        foreach (var i in Health) if (i.SourceAddress == old) i.SourceAddress = address;

        // Collector-supplied Extra values (e.g. vSphere "VI SDK Server") and raw sheet rows.
        foreach (var obj in AllObjects())
        {
            foreach (var key in obj.Extra.Keys.ToList())
            {
                switch (obj.Extra[key])
                {
                    case string str when string.Equals(str, old, StringComparison.OrdinalIgnoreCase):
                        obj.Extra[key] = address;
                        break;
                    case IEnumerable<Dictionary<string, object?>> rows:
                        foreach (var row in rows)
                        foreach (var k in row.Keys.ToList())
                            if (row[k] is string v && string.Equals(v, old, StringComparison.OrdinalIgnoreCase)) row[k] = address;
                        break;
                }
            }
        }
    }

    /// <summary>Every inventory object in the snapshot (for bulk fix-ups such as rekeying or normalising Extra).</summary>
    public IEnumerable<InventoryObject> AllObjects()
    {
        yield return Source;
        foreach (var c in Clusters)
        {
            yield return c;
            foreach (var n in c.Nodes) yield return n;
            foreach (var n in c.Networks) yield return n;
        }
        foreach (var h in Hosts)
        {
            yield return h;
            foreach (var o in h.Nics) yield return o;
            foreach (var o in h.Switches) yield return o;
            foreach (var o in h.PortGroups) yield return o;
            foreach (var o in h.IpInterfaces) yield return o;
            foreach (var o in h.StorageAdapters) yield return o;
        }
        foreach (var v in VirtualMachines)
        {
            yield return v;
            foreach (var o in v.Disks) yield return o;
            foreach (var o in v.Nics) yield return o;
            foreach (var o in v.CdDrives) yield return o;
            foreach (var o in v.UsbDevices) yield return o;
            foreach (var o in v.Snapshots) yield return o;
            foreach (var o in v.Partitions) yield return o;
        }
        foreach (var d in Datastores) yield return d;
    }
}
