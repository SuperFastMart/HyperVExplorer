namespace HypervisorExplorer.Core.Model;

/// <summary>
/// The merged inventory across all connected sources. Each source contributes one
/// <see cref="InventorySnapshot"/>; refreshing a source replaces its snapshot wholesale.
/// Thread-safe: collectors may complete on background threads.
/// </summary>
public sealed class InventoryStore
{
    private readonly object _gate = new();
    private readonly Dictionary<string, InventorySnapshot> _snapshots = new(StringComparer.OrdinalIgnoreCase);
    private Inventory _current = Inventory.Empty;

    public event EventHandler? Changed;

    public Inventory Current
    {
        get { lock (_gate) return _current; }
    }

    public IReadOnlyCollection<string> SourceAddresses
    {
        get { lock (_gate) return _snapshots.Keys.ToArray(); }
    }

    public bool Contains(string address)
    {
        lock (_gate) return _snapshots.ContainsKey(address);
    }

    public void Upsert(InventorySnapshot snapshot)
    {
        lock (_gate)
        {
            _snapshots[snapshot.Source.Address] = snapshot;
            _current = Inventory.FromSnapshots(_snapshots.Values);
        }
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public bool Remove(string address)
    {
        bool removed;
        lock (_gate)
        {
            removed = _snapshots.Remove(address);
            if (removed) _current = Inventory.FromSnapshots(_snapshots.Values);
        }
        if (removed) Changed?.Invoke(this, EventArgs.Empty);
        return removed;
    }

    public void Clear()
    {
        lock (_gate)
        {
            _snapshots.Clear();
            _current = Inventory.Empty;
        }
        Changed?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>Returns the raw snapshots, e.g. for JSON export.</summary>
    public IReadOnlyList<InventorySnapshot> Snapshots
    {
        get { lock (_gate) return _snapshots.Values.ToList(); }
    }
}

/// <summary>An immutable, flattened view over a set of snapshots.</summary>
public sealed class Inventory
{
    public static readonly Inventory Empty = new([], [], [], [], [], [], []);

    public Inventory(
        IReadOnlyList<InventorySnapshot> snapshots,
        IReadOnlyList<Source> sources,
        IReadOnlyList<ClusterInfo> clusters,
        IReadOnlyList<HostSystem> hosts,
        IReadOnlyList<VirtualMachine> vms,
        IReadOnlyList<Datastore> datastores,
        IReadOnlyList<HealthItem> health)
    {
        Snapshots = snapshots;
        Sources = sources;
        Clusters = clusters;
        Hosts = hosts;
        VirtualMachines = vms;
        Datastores = datastores;
        Health = health;
        _hostsByName = hosts
            .GroupBy(h => (h.SourceAddress, h.Name))
            .ToDictionary(g => g.Key, g => g.First());
    }

    private readonly Dictionary<(string Source, string Name), HostSystem> _hostsByName;

    public IReadOnlyList<InventorySnapshot> Snapshots { get; }
    public IReadOnlyList<Source> Sources { get; }
    public IReadOnlyList<ClusterInfo> Clusters { get; }
    public IReadOnlyList<HostSystem> Hosts { get; }
    public IReadOnlyList<VirtualMachine> VirtualMachines { get; }
    public IReadOnlyList<Datastore> Datastores { get; }
    public IReadOnlyList<HealthItem> Health { get; }

    public HostSystem? FindHost(string sourceAddress, string name) =>
        _hostsByName.GetValueOrDefault((sourceAddress, name));

    public Source? FindSource(string address) =>
        Sources.FirstOrDefault(s => string.Equals(s.Address, address, StringComparison.OrdinalIgnoreCase));

    public static Inventory FromSnapshots(IEnumerable<InventorySnapshot> snapshots)
    {
        var list = snapshots.OrderBy(s => s.Source.Address, StringComparer.OrdinalIgnoreCase).ToList();
        var vms = list.SelectMany(s => s.VirtualMachines)
            .OrderBy(v => v.Name, StringComparer.OrdinalIgnoreCase).ToList();
        var health = list.SelectMany(s => s.Health).ToList();
        var inv = new Inventory(
            list,
            list.Select(s => s.Source).ToList(),
            list.SelectMany(s => s.Clusters).ToList(),
            list.SelectMany(s => s.Hosts).OrderBy(h => h.Name, StringComparer.OrdinalIgnoreCase).ToList(),
            vms,
            list.SelectMany(s => s.Datastores).OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase).ToList(),
            health);
        // Derived health checks run over the merged view so cross-source rules see everything.
        var derived = HealthAnalyzer.Analyze(inv);
        return derived.Count == 0
            ? inv
            : new Inventory(inv.Snapshots, inv.Sources, inv.Clusters, inv.Hosts, inv.VirtualMachines,
                inv.Datastores, [.. health, .. derived]);
    }
}
