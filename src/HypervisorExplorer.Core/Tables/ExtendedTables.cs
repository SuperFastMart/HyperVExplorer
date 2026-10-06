using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Tables;

public sealed record ClusterNodeRow(ClusterInfo Cluster, ClusterNode Node);
public sealed record ClusterNetworkRow(ClusterInfo Cluster, ClusterNetwork Network);

/// <summary>Builds a table with free-form columns (for data RVTools has no sheet for).</summary>
public sealed class TableBuilder<T> where T : class
{
    private readonly string _name;
    private readonly List<TableColumn> _columns = [];

    public TableBuilder(string name) => _name = name;

    public TableBuilder<T> Text(string header, Func<T, object?> get) => Add(header, CellKind.Text, get);
    public TableBuilder<T> Num(string header, Func<T, object?> get) => Add(header, CellKind.Number, get);
    public TableBuilder<T> Date(string header, Func<T, object?> get) => Add(header, CellKind.Date, get);
    public TableBuilder<T> Bool(string header, Func<T, bool?> get) => Add(header, CellKind.Bool, r => get(r));

    private TableBuilder<T> Add(string header, CellKind kind, Func<T, object?> get)
    {
        _columns.Add(new TableColumn(header, kind, o => get((T)o)));
        return this;
    }

    public TableDefinition Build(Func<Inventory, IEnumerable<T>> rows, string description,
        Func<T, VirtualMachine?>? vmOf = null, Func<T, (string, string?, string?)>? scopeOf = null) => new()
    {
        Name = _name,
        Description = description,
        Columns = _columns,
        Rows = inv => rows(inv).Cast<object>().ToList(),
        VmOf = vmOf is null ? _ => null : o => vmOf((T)o),
        ScopeOf = scopeOf is null ? _ => ("", null, null) : o => scopeOf((T)o),
    };
}

/// <summary>Tables beyond the RVTools schema: cluster membership, cluster networks, VM HA roles.</summary>
public static class ExtendedTables
{
    public static IReadOnlyList<TableDefinition> All { get; } =
    [
        new TableBuilder<VmRow>("hvOverview")
            .Text("VM", r => r.Vm.Name)
            .Text("Platform", r => RvToolsTables.PlatformName(r.Vm.Platform))
            .Text("Kind", r => r.Vm.Kind == GuestKind.Container ? "Container" : r.Vm.IsTemplate ? "Template" : "VM")
            .Text("State", r => r.Vm.RawState ?? RvToolsTables.PowerStateText(r.Vm.PowerState))
            .Text("Host", r => r.Vm.Host)
            .Text("Cluster", r => r.Vm.Cluster)
            .Num("vCPU", r => r.Vm.CpuCount)
            .Num("Memory MiB", r => r.Vm.MemoryMiB)
            .Num("Provisioned GiB", r => Math.Round(r.Vm.ProvisionedBytes / 1073741824.0, 1))
            .Text("Primary IP", r => r.Vm.PrimaryIp)
            .Text("Guest OS", r => r.Vm.OsGuest ?? r.Vm.OsConfigured)
            .Text("Uptime", r => r.Vm.Uptime is { } u ? $"{(int)u.TotalDays}d {u.Hours}h {u.Minutes}m" : null)
            .Num("Snapshots", r => r.Vm.Snapshots.Count)
            .Text("Tools", r => r.Vm.ToolsStatus)
            .Text("Generation / HW", r => r.Vm.Generation ?? r.Vm.HardwareVersion)
            .Text("Tags", r => string.Join(", ", r.Vm.Tags))
            .Text("Notes", r => r.Vm.Annotation)
            .Text("Source", r => r.Vm.SourceAddress)
            .Build(inv => inv.VirtualMachines.Select(v => new VmRow(v, inv)),
                "Cross-platform summary of every VM and container.", r => r.Vm, r => (r.Vm.SourceAddress, r.Vm.Host, r.Vm.Cluster)),

        new TableBuilder<ClusterNodeRow>("hvClusterNodes")
            .Text("Cluster", r => r.Cluster.Name)
            .Text("Platform", r => RvToolsTables.PlatformName(r.Cluster.Platform))
            .Text("Node", r => r.Node.Name)
            .Text("State", r => r.Node.State)
            .Text("Drain status", r => r.Node.DrainStatus)
            .Text("Address", r => r.Node.Address)
            .Num("Votes", r => r.Node.Votes)
            .Text("Node ID", r => r.Node.Id)
            .Text("Quorum", r => r.Cluster.QuorumType)
            .Text("Witness", r => r.Cluster.QuorumWitness)
            .Text("Source", r => r.Cluster.SourceAddress)
            .Build(inv => inv.Clusters.SelectMany(c => c.Nodes.Select(n => new ClusterNodeRow(c, n))),
                "Cluster membership: Failover Cluster nodes, PVE cluster nodes, vSphere cluster hosts.",
                scopeOf: r => (r.Cluster.SourceAddress, r.Node.Name, r.Cluster.Name)),

        new TableBuilder<ClusterNetworkRow>("hvClusterNetworks")
            .Text("Cluster", r => r.Cluster.Name)
            .Text("Network", r => r.Network.Name)
            .Text("Role", r => r.Network.Role)
            .Text("Address", r => r.Network.Address)
            .Text("State", r => r.Network.State)
            .Text("Metric", r => r.Network.Metric)
            .Text("Source", r => r.Cluster.SourceAddress)
            .Build(inv => inv.Clusters.SelectMany(c => c.Networks.Select(n => new ClusterNetworkRow(c, n))),
                "Failover Cluster networks and roles.", scopeOf: r => (r.Cluster.SourceAddress, null, r.Cluster.Name)),

        new TableBuilder<VmRow>("hvHA")
            .Text("VM", r => r.Vm.Name)
            .Text("Cluster", r => r.Vm.Cluster)
            .Bool("HA protected", r => r.Vm.HaProtected)
            .Text("HA state", r => r.Vm.HaState)
            .Text("Owner node", r => r.Vm.OwnerNode ?? r.Vm.Host)
            .Text("Preferred owners", r => string.Join(", ", r.Vm.PreferredOwners))
            .Text("Priority", r => r.Vm.FailoverPriority)
            .Bool("Auto start", r => r.Vm.AutoStart)
            .Num("Start delay (s)", r => r.Vm.StartDelaySeconds)
            .Text("Source", r => r.Vm.SourceAddress)
            .Build(inv => inv.VirtualMachines.Where(v => v.Cluster is not null || v.HaProtected is not null)
                    .Select(v => new VmRow(v, inv)),
                "High availability: Failover Cluster roles, PVE HA resources, vSphere HA.",
                r => r.Vm, r => (r.Vm.SourceAddress, r.Vm.Host, r.Vm.Cluster)),
    ];
}
