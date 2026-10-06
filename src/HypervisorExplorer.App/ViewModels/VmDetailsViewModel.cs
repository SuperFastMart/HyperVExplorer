using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.ViewModels;

public sealed record DetailField(string Label, string Value);

public sealed record DetailSection(string Title, IReadOnlyList<DetailField> Fields);

/// <summary>Read-only summary of one VM for the details panel.</summary>
public sealed class VmDetailsViewModel
{
    public VmDetailsViewModel(VirtualMachine vm)
    {
        Vm = vm;
        Title = vm.Name;
        Subtitle = $"{RvToolsTables.PlatformName(vm.Platform)} · {(vm.Kind == GuestKind.Container ? "Container" : vm.IsTemplate ? "Template" : "VM")} · {vm.RawState ?? RvToolsTables.PowerStateText(vm.PowerState)}";

        var sections = new List<DetailSection>
        {
            Section("General",
                ("Host", vm.Host), ("Cluster", vm.Cluster), ("Source", vm.SourceAddress), ("VM ID", vm.VmId),
                ("Guest OS", vm.OsGuest ?? vm.OsConfigured), ("DNS name", vm.DnsName), ("Primary IP", vm.PrimaryIp),
                ("Uptime", vm.Uptime is { } u ? $"{(int)u.TotalDays}d {u.Hours}h {u.Minutes}m" : null),
                ("Created", vm.CreationDate?.LocalDateTime.ToString("yyyy-MM-dd")),
                ("Firmware", Join(vm.Firmware, vm.SecureBoot == true ? "Secure Boot" : null)),
                ("Generation / HW", vm.Generation ?? vm.HardwareVersion),
                ("Tools", Join(vm.ToolsStatus, vm.ToolsVersion)), ("Heartbeat", vm.Heartbeat),
                ("HA", vm.HaProtected == true ? Join("Protected", vm.HaState, vm.OwnerNode is null ? null : "owner " + vm.OwnerNode) : null),
                ("Pool / folder", vm.ResourcePool), ("Tags", vm.Tags.Count > 0 ? string.Join(", ", vm.Tags) : null),
                ("Notes", vm.Annotation)),
            Section("Compute",
                ("vCPU", Join(vm.CpuCount.ToString(), vm.Sockets is { } s && vm.CoresPerSocket is { } c ? $"{s} socket(s) × {c} core(s)" : null)),
                ("CPU type", vm.CpuType),
                ("Memory", $"{vm.MemoryMiB:N0} MiB" + (vm.DynamicMemory == true ? $" (dynamic {vm.MemoryMinMiB:N0}–{vm.MemoryMaxMiB:N0})" : "")),
                ("Assigned / demand", vm.MemoryAssignedMiB is null ? null : $"{vm.MemoryAssignedMiB:N0} / {vm.MemoryDemandMiB:N0} MiB")),
        };

        sections.Add(new DetailSection($"Disks ({vm.Disks.Count})", vm.Disks.Select(d => new DetailField(
            d.Label,
            Join(d.CapacityBytes is { } cap ? Gib(cap) : null,
                 d.UsedBytes is { } used ? $"{Gib(used)} used" : null,
                 d.Format, d.Thin == true ? "thin" : null, d.Datastore, d.Path) ?? "")).ToList()));

        sections.Add(new DetailSection($"Network ({vm.Nics.Count})", vm.Nics.Select(n => new DetailField(
            n.Label,
            Join(n.AdapterType, n.Network, n.Vlan is > 0 ? $"VLAN {n.Vlan}" : null, n.Mac,
                 n.Ipv4.Count + n.Ipv6.Count > 0 ? string.Join(", ", n.Ipv4.Concat(n.Ipv6)) : null,
                 n.Connected == false ? "disconnected" : null) ?? "")).ToList()));

        if (vm.Snapshots.Count > 0)
            sections.Add(new DetailSection($"Snapshots ({vm.Snapshots.Count})", vm.Snapshots.Select(s => new DetailField(
                s.Name, Join(s.Created?.LocalDateTime.ToString("yyyy-MM-dd HH:mm"), s.Description) ?? "")).ToList()));

        if (vm.Partitions.Count > 0)
            sections.Add(new DetailSection("Guest file systems", vm.Partitions.Select(p => new DetailField(
                p.Name, Join(p.FileSystem, p.CapacityBytes is { } c ? Gib(c) : null,
                    p.FreeBytes is { } f ? $"{Gib(f)} free" : null) ?? "")).ToList()));

        if (vm.CdDrives.Count > 0)
            sections.Add(new DetailSection("CD/DVD", vm.CdDrives.Select(c => new DetailField(c.DeviceNode, c.Media ?? "(empty)")).ToList()));

        Sections = sections.Where(s => s.Fields.Count > 0).ToList();
    }

    public VirtualMachine Vm { get; }
    public string Title { get; }
    public string Subtitle { get; }
    public IReadOnlyList<DetailSection> Sections { get; }

    private static DetailSection Section(string title, params (string Label, string? Value)[] fields) =>
        new(title, fields.Where(f => !string.IsNullOrWhiteSpace(f.Value)).Select(f => new DetailField(f.Label, f.Value!)).ToList());

    private static string? Join(params string?[] parts)
    {
        var p = parts.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        return p.Count == 0 ? null : string.Join(" · ", p);
    }

    private static string Gib(long bytes) => bytes >= 1L << 40 ? $"{bytes / 1099511627776.0:0.##} TiB" : $"{bytes / 1073741824.0:0.#} GiB";
}
