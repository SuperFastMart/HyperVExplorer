using System.Xml.Linq;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.VMware;

/// <summary>
/// Maps raw vSphere property data to the inventory model. Values without a model property go into
/// <see cref="InventoryObject.Extra"/> keyed by the exact RVTools header. Where the same header appears on two
/// sheets with different meanings (vCPU/vMemory "Level", "Shares", "Limit", "Max", "Entitlement"…), a sheet-qualified
/// key ("vMemory:Level") is used. Whole sheets the model cannot hold (vRP, dvSwitch, dvPort, vMultiPath, vLicense)
/// are stored as rows in <c>Source.Extra["__sheet:{name}"]</c>.
/// </summary>
public sealed partial class VMwareInventoryMapper
{
    /// <summary>Prefix of Source.Extra keys that hold whole RVTools sheets as <c>List&lt;Dictionary&lt;string, object?&gt;&gt;</c>.</summary>
    public const string SheetPrefix = "__sheet:";

    private const double MiB = 1024 * 1024;

    private readonly ConnectionRequest _req;
    private readonly VSphereData _d;
    private readonly ServiceContent _sc;
    private readonly string _addr;
    private readonly string? _uuid;
    private readonly Dictionary<string, Entity> _entities = new(StringComparer.Ordinal);
    private readonly Dictionary<string, VimObject> _hostObjs = new(StringComparer.Ordinal);
    private readonly Dictionary<string, VimObject> _dvsByUuid = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, VimObject> _pgByKey = new(StringComparer.Ordinal);
    private readonly Dictionary<string, ClusterHa> _clusterHa = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Dictionary<string, string>> _hostPortgroupSwitch = new(StringComparer.Ordinal);
    private InventorySnapshot _snap = null!;

    private sealed record Entity(MoRef Ref, string Name, MoRef? Parent);

    private sealed record ClusterRule(string Name, string Type, HashSet<string> Vms);

    private sealed class ClusterHa
    {
        public bool HaEnabled;
        public XElement? DasConfig;
        public Dictionary<string, XElement> VmOverrides { get; } = new(StringComparer.Ordinal);
        public List<ClusterRule> Rules { get; } = [];
    }

    public VMwareInventoryMapper(ConnectionRequest request, VSphereData data)
    {
        _req = request;
        _d = data;
        _sc = data.Content;
        _addr = request.Address;
        _uuid = string.IsNullOrEmpty(_sc.InstanceUuid) ? null : _sc.InstanceUuid;

        foreach (var e in data.Entities) AddEntity(e);
        // Typed retrievals also carry name/parent; use them in case the ManagedEntity pass was partial.
        foreach (var e in data.VirtualMachines.Concat(data.Hosts).Concat(data.Datastores).Concat(data.ComputeResources)
                     .Concat(data.ResourcePools).Concat(data.DistributedSwitches))
        {
            if (!_entities.ContainsKey(e.Ref.Value)) AddEntity(e);
        }
        foreach (var pg in data.DistributedPortgroups)
        {
            if (!_entities.ContainsKey(pg.Ref.Value)) AddEntity(pg);
            _pgByKey[pg.Ref.Value] = pg;
            if ((pg.Str("key") ?? pg.Str("config.key")) is { } key) _pgByKey[key] = pg;
        }
        foreach (var h in data.Hosts) _hostObjs[h.Ref.Value] = h;
        foreach (var s in data.DistributedSwitches)
        {
            if ((s.Str("uuid") ?? s.Str("summary.uuid")) is { } u) _dvsByUuid[u] = s;
        }
    }

    private void AddEntity(VimObject o) =>
        _entities[o.Ref.Value] = new Entity(o.Ref, o.Str("name") ?? o.Ref.Value, o.Ref1("parent"));

    /// <summary>Builds the snapshot.</summary>
    public InventorySnapshot Map()
    {
        _snap = new InventorySnapshot { Source = MapSource() };
        MapClusters();
        MapHosts();
        MapDatastores();
        MapVirtualMachines();
        FinishClusters();
        _snap.Source.Extra[SheetPrefix + "vRP"] = ResourcePoolRows();
        _snap.Source.Extra[SheetPrefix + "dvSwitch"] = DvSwitchRows();
        _snap.Source.Extra[SheetPrefix + "dvPort"] = DvPortRows();
        _snap.Source.Extra[SheetPrefix + "vMultiPath"] = MultiPathRows();
        _snap.Source.Extra[SheetPrefix + "vLicense"] = LicenseRows();
        AddHealth();
        StampSdkUuid();
        return _snap;
    }

    // ---------------------------------------------------------------- hierarchy helpers

    private string NameOf(MoRef? r) =>
        r is null ? "" : _entities.TryGetValue(r.Value, out var e) ? e.Name : r.Value;

    private MoRef? ParentOf(MoRef? r) => r is not null && _entities.TryGetValue(r.Value, out var e) ? e.Parent : null;

    private string? DatacenterOf(MoRef? r)
    {
        for (var i = 0; r is not null && i < 128; i++, r = ParentOf(r))
        {
            if (r.Type == "Datacenter") return DatacenterDisplay(NameOf(r));
        }
        return null;
    }

    /// <summary>Standalone ESXi always reports "ha-datacenter"; prefer the user's group label there.</summary>
    private string DatacenterDisplay(string name) =>
        name == "ha-datacenter" && !string.IsNullOrWhiteSpace(_req.Group) ? _req.Group! : name;

    /// <summary>Inventory path "/DC/vm/Folder" from the datacenter down to <paramref name="r"/> (inclusive).</summary>
    private string? InventoryPath(MoRef? r)
    {
        if (r is null) return null;
        var parts = new List<string>();
        for (var i = 0; r is not null && i < 128; i++, r = ParentOf(r))
        {
            parts.Add(NameOf(r));
            if (r.Type == "Datacenter") break;
        }
        parts.Reverse();
        return "/" + string.Join("/", parts);
    }

    private MoRef? ComputeResourceOfHost(MoRef? host) => ParentOf(host) is { } p && p.Type.EndsWith("ComputeResource", StringComparison.Ordinal) ? p : null;

    private string? ClusterOfHost(MoRef? host) =>
        ComputeResourceOfHost(host) is { Type: "ClusterComputeResource" } c ? NameOf(c) : null;

    private static long? Mib(long? bytes) => bytes is null ? null : (long)Math.Round(bytes.Value / MiB);

    /// <summary>"[datastore1] vm/vm.vmdk" → "datastore1".</summary>
    internal static string? DatastoreFromPath(string? path)
    {
        if (string.IsNullOrEmpty(path) || path[0] != '[') return null;
        var end = path.IndexOf(']');
        return end > 1 ? path[1..end] : null;
    }

    private static string? Join(IEnumerable<string?> items)
    {
        var s = string.Join(", ", items.Where(i => !string.IsNullOrWhiteSpace(i)));
        return s.Length == 0 ? null : s;
    }

    // ---------------------------------------------------------------- vSource

    private Source MapSource()
    {
        var a = _sc.About;
        var src = new Source
        {
            Address = _addr,
            Platform = Platform.VMware,
            Group = _req.Group,
            ProductName = a.Str("name") ?? "VMware",
            Version = a.Str("version") ?? "",
            Build = a.Str("build"),
            ApiVersion = a.Str("apiVersion"),
            Vendor = a.Str("vendor") ?? "VMware, Inc.",
            OsType = a.Str("osType"),
            CollectedAt = DateTimeOffset.Now,
        };
        var x = src.Extra;
        x["Name"] = a.Str("name");
        x["OS type"] = a.Str("osType");
        x["API type"] = a.Str("apiType");
        x["API version"] = a.Str("apiVersion");
        x["Version"] = a.Str("version");
        x["Patch level"] = a.Str("patchLevel");
        x["Build"] = a.Str("build");
        x["Fullname"] = a.Str("fullName");
        x["Product name"] = a.Str("licenseProductName");
        x["Product version"] = a.Str("licenseProductVersion");
        x["Product line"] = a.Str("productLineId");
        x["Vendor"] = a.Str("vendor");
        return src;
    }

    // ---------------------------------------------------------------- vHealth

    private void AddHealth()
    {
        void Add(string name, string message, HealthSeverity sev) =>
            _snap.Health.Add(new HealthItem { Name = name, Message = message, Severity = sev, SourceAddress = _addr });

        static HealthSeverity? StatusSeverity(string? s) => s switch
        {
            "red" => HealthSeverity.Error,
            "yellow" => HealthSeverity.Warning,
            _ => null,
        };

        foreach (var h in _d.Hosts)
        {
            var name = h.Str("name") ?? h.Ref.Value;
            var conn = h.Str("runtime.connectionState");
            if (conn is "disconnected" or "notResponding")
                Add(name, $"Host connection state is {conn}", HealthSeverity.Error);
            if (h.Bool("runtime.inQuarantineMode") == true)
                Add(name, "Host is in quarantine mode", HealthSeverity.Warning);
            if (StatusSeverity(h.Str("overallStatus")) is { } sev)
                Add(name, $"Host overall status is {h.Str("overallStatus")}", sev);
        }
        foreach (var c in _d.ComputeResources.Where(c => c.Ref.Type == "ClusterComputeResource"))
        {
            if (StatusSeverity(c.Str("overallStatus")) is { } sev)
                Add(c.Str("name") ?? c.Ref.Value, $"Cluster overall status is {c.Str("overallStatus")}", sev);
        }
        foreach (var ds in _d.Datastores)
        {
            if (StatusSeverity(ds.Str("overallStatus")) is { } sev && ds.Bool("summary.accessible") != false)
                Add(ds.Str("name") ?? ds.Ref.Value, $"Datastore overall status is {ds.Str("overallStatus")}", sev);
            if (ds.Str("summary.maintenanceMode") is "inMaintenance" or "enteringMaintenance")
                Add(ds.Str("name") ?? ds.Ref.Value, $"Datastore maintenance mode: {ds.Str("summary.maintenanceMode")}", HealthSeverity.Warning);
        }
        foreach (var o in _d.VirtualMachines)
        {
            var name = o.Str("name") ?? o.Ref.Value;
            var conn = o.Str("runtime.connectionState");
            if (conn is "orphaned" or "inaccessible" or "invalid" or "disconnected")
                Add(name, $"VM connection state is {conn}", HealthSeverity.Error);
            if (o.Bool("runtime.consolidationNeeded") == true)
                Add(name, "Virtual machine disks consolidation is needed", HealthSeverity.Warning);
            if (o.Bool("config.template") != true && o.Str("runtime.powerState") == "poweredOn")
            {
                var tv = o.Str("guest.toolsVersionStatus2");
                if (tv is "guestToolsNeedUpgrade" or "guestToolsSupportedOld" or "guestToolsTooOld")
                    Add(name, $"VMware Tools are out of date ({tv})", HealthSeverity.Info);
                if (tv is "guestToolsNotInstalled")
                    Add(name, "VMware Tools are not installed", HealthSeverity.Warning);
            }
            if (StatusSeverity(o.Str("overallStatus")) is { } sev)
                Add(name, $"VM overall status is {o.Str("overallStatus")}", sev);
        }
        foreach (var lic in _d.Licenses)
        {
            var exp = LicenseExpiration(lic);
            if (exp is { } e && e < DateTimeOffset.Now.AddDays(30) && lic.Str("licenseKey") != "00000-00000-00000-00000-00000")
                Add(lic.Str("name") ?? "License", e < DateTimeOffset.Now ? $"License expired on {e:yyyy-MM-dd}" : $"License expires on {e:yyyy-MM-dd}",
                    e < DateTimeOffset.Now ? HealthSeverity.Error : HealthSeverity.Warning);
        }
    }

    /// <summary>Sets "VI SDK UUID" (vCenter instanceUuid) on every object so per-row columns match RVTools.</summary>
    private void StampSdkUuid()
    {
        if (_uuid is null) return;
        void S(InventoryObject o) => o.Extra["VI SDK UUID"] = _uuid;
        S(_snap.Source);
        _snap.Clusters.ForEach(S);
        _snap.Datastores.ForEach(S);
        foreach (var h in _snap.Hosts)
        {
            S(h);
            h.Nics.ForEach(S);
            h.Switches.ForEach(S);
            h.PortGroups.ForEach(S);
            h.IpInterfaces.ForEach(S);
            h.StorageAdapters.ForEach(S);
        }
        foreach (var v in _snap.VirtualMachines)
        {
            S(v);
            v.Disks.ForEach(S);
            v.Nics.ForEach(S);
            v.CdDrives.ForEach(S);
            v.UsbDevices.ForEach(S);
            v.Snapshots.ForEach(S);
            v.Partitions.ForEach(S);
        }
    }

    private Dictionary<string, object?> SheetRow() => new(StringComparer.Ordinal)
    {
        ["VI SDK Server"] = _addr,
        ["VI SDK UUID"] = _uuid ?? _addr,
    };
}
