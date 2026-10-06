using System.Text.Json;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>
/// Collects inventory from a Proxmox VE node or cluster over the PVE REST API (https, port 8006).
/// One connection to any node yields the whole cluster: nodes, QEMU VMs, LXC containers, storage and HA.
/// Requires at least the PVEAuditor role on '/'.
/// </summary>
public sealed class ProxmoxCollector : IInventoryCollector
{
    private readonly HttpMessageHandler? _handler;

    public ProxmoxCollector()
    {
    }

    /// <summary>Test seam: all HTTP goes through <paramref name="handler"/> (not disposed by the collector).</summary>
    internal ProxmoxCollector(HttpMessageHandler handler) => _handler = handler;

    public Platform Platform => Platform.Proxmox;

    /// <summary>Maximum concurrent API requests per collection.</summary>
    internal int MaxConcurrency { get; init; } = 8;
    internal TimeSpan RequestTimeout { get; init; } = TimeSpan.FromSeconds(30);
    /// <summary>Guest-agent calls are short: an unresponsive agent must not stall the collection.</summary>
    internal TimeSpan AgentTimeout { get; init; } = TimeSpan.FromSeconds(6);
    internal Func<DateTimeOffset> Clock { get; init; } = () => DateTimeOffset.Now;

    public async Task<InventorySnapshot> CollectAsync(ConnectionRequest request, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (string.IsNullOrWhiteSpace(request.Address))
            throw new CollectionException(CollectionFailure.Other, "No Proxmox VE address was given.");
        if (request.CredentialKind != CredentialKind.CurrentUser && string.IsNullOrWhiteSpace(request.Username))
            throw new CollectionException(CollectionFailure.AuthenticationFailed, "No username or API token ID was given.",
                request.CredentialKind == CredentialKind.ApiToken
                    ? "Enter the token ID as user@realm!tokenname (e.g. root@pam!explorer) and the token secret."
                    : "Enter the username as user@realm (e.g. root@pam).");

        var baseUri = PveApiClient.BuildBaseUri(request.Address, request.Port);
        var handler = _handler ?? PveApiClient.CreateDefaultHandler(request.IgnoreCertificateErrors);
        using var api = new PveApiClient(handler, disposeHandler: _handler is null, baseUri, MaxConcurrency, RequestTimeout);

        JsonElement version, resources;
        try
        {
            progress?.Report($"Connecting to {baseUri.Host}:{baseUri.Port}...");
            await api.LoginAsync(request, cancellationToken).ConfigureAwait(false);
            version = await api.GetAsync("/version", cancellationToken).ConfigureAwait(false);
            progress?.Report($"Reading cluster resources from {baseUri.Host}...");
            resources = await api.GetAsync("/cluster/resources", cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw PveErrors.Map(ex, request, baseUri, cancellationToken);
        }

        if (!resources.Items().Any(r => r.Str("type") == "node"))
            throw new CollectionException(CollectionFailure.PermissionDenied,
                $"Signed in to {baseUri.Host}, but the account cannot see any nodes.", PveErrors.PermissionHint);

        try
        {
            var run = new Run(this, api, request, progress, cancellationToken);
            return await run.CollectAsync(version, resources).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            throw PveErrors.Map(ex, request, baseUri, cancellationToken);
        }
    }

    /// <summary>State for one collection.</summary>
    private sealed class Run
    {
        private readonly ProxmoxCollector _owner;
        private readonly PveApiClient _api;
        private readonly ConnectionRequest _req;
        private readonly IProgress<string>? _progress;
        private readonly CancellationToken _ct;
        private readonly DateTimeOffset _now;
        private readonly List<string> _warnings = [];
        private readonly Dictionary<(string Node, string Storage), string> _storageTypes = new();
        private readonly Dictionary<string, JsonElement> _storageConfig = new(StringComparer.Ordinal);
        private readonly Dictionary<string, (long? Size, long? Used, string? Format)> _content = new(StringComparer.Ordinal);
        private readonly HashSet<string> _onlineNodes = new(StringComparer.Ordinal);
        private string? _cluster;
        private string _product = "Proxmox VE";
        private string? _apiVersion;
        private int _guestsDone;
        private int _guestsTotal;

        public Run(ProxmoxCollector owner, PveApiClient api, ConnectionRequest req, IProgress<string>? progress, CancellationToken ct)
        {
            _owner = owner;
            _api = api;
            _req = req;
            _progress = progress;
            _ct = ct;
            _now = owner.Clock();
        }

        private void Warn(string message)
        {
            lock (_warnings) _warnings.Add(message);
        }

        private static string Seg(string s) => PveApiClient.Seg(s);

        public async Task<InventorySnapshot> CollectAsync(JsonElement version, JsonElement resources)
        {
            var ver = version.Str("version") ?? "";
            _product = string.IsNullOrEmpty(ver) ? "Proxmox VE" : $"Proxmox VE {ver}";
            var release = version.Str("release");
            var repoid = version.Str("repoid");
            _apiVersion = release is null ? repoid : repoid is null ? release : $"{release}/{repoid}";

            var source = new Source
            {
                Address = _req.Address,
                Platform = Platform.Proxmox,
                Group = _req.Group,
                ProductName = "Proxmox VE",
                Version = ver,
                ApiVersion = _apiVersion,
                Vendor = "Proxmox Server Solutions GmbH",
                OsType = "Linux",
                CollectedAt = _now,
            };

            // Cluster-wide, independent calls (all optional except resources, which we already have).
            var statusTask = _api.TryGetAsync("/cluster/status", _ct);
            var storageTask = _api.TryGetAsync("/storage", _ct);
            var haResTask = _api.TryGetAsync("/cluster/ha/resources", _ct);
            var haStatusTask = _api.TryGetAsync("/cluster/ha/status/current", _ct);
            var haGroupsTask = _api.TryGetAsync("/cluster/ha/groups", _ct);
            await Task.WhenAll(statusTask, storageTask, haResTask, haStatusTask, haGroupsTask).ConfigureAwait(false);
            var clusterStatus = statusTask.Result;
            var haRes = haResTask.Result;
            var haStatus = haStatusTask.Result;

            foreach (var s in storageTask.Result.Items())
                if (s.Str("storage") is { } id) _storageConfig[id] = s;

            var all = resources.Items().ToList();
            var nodeRes = all.Where(r => r.Str("type") == "node" && r.Str("node") is not null)
                .OrderBy(r => r.Str("node"), StringComparer.OrdinalIgnoreCase).ToList();
            var guestRes = all.Where(r => r.Str("type") is "qemu" or "lxc" && r.Str("node") is not null && r.Long("vmid") is not null).ToList();
            var storageRes = all.Where(r => r.Str("type") == "storage" && r.Str("storage") is not null && r.Str("node") is not null).ToList();

            // Cluster membership (absent on a standalone node).
            ClusterInfo? cluster = null;
            var statusOnline = new Dictionary<string, bool>(StringComparer.Ordinal);
            if (clusterStatus is { } cs)
            {
                foreach (var n in cs.Items().Where(e => e.Str("type") == "node" && e.Str("name") is not null))
                    statusOnline[n.Str("name")!] = n.Bool("online") ?? true;
                if (cs.Items().FirstOrDefault(e => e.Str("type") == "cluster") is { ValueKind: JsonValueKind.Object } ce)
                {
                    _cluster = ce.Str("name");
                    cluster = new ClusterInfo
                    {
                        Name = _cluster ?? "",
                        Platform = Platform.Proxmox,
                        SourceAddress = _req.Address,
                        Datacenter = _req.Group,
                        Quorate = ce.Bool("quorate"),
                        Id = _cluster,
                        QuorumType = "Corosync votequorum",
                        Nodes = cs.Items().Where(e => e.Str("type") == "node").Select(n => new ClusterNode
                        {
                            Name = n.Str("name") ?? "",
                            State = n.Bool("online") == false ? "offline" : "online",
                            Id = n.Str("nodeid"),
                            Address = n.Str("ip"),
                        }).OrderBy(n => n.Name, StringComparer.OrdinalIgnoreCase).ToList(),
                    };
                    if (ce.Str("version") is { } cv) cluster.Extra["Config version"] = cv;
                }
            }

            // Hosts.
            var hosts = new List<HostSystem>();
            foreach (var r in nodeRes)
            {
                var name = r.Str("node")!;
                var online = r.Str("status") == "online" && statusOnline.GetValueOrDefault(name, true);
                var host = new HostSystem
                {
                    Name = name,
                    Platform = Platform.Proxmox,
                    SourceAddress = _req.Address,
                    Datacenter = _req.Group,
                    Cluster = _cluster,
                    Status = online ? "green" : "red",
                    CpuThreads = r.Int("maxcpu"),
                    MemoryBytes = r.Long("maxmem"),
                    MemoryUsedBytes = online ? r.Long("mem") : null,
                    CpuUsagePercent = online && r.Dbl("cpu") is { } cpu ? Math.Round(cpu * 100, 1) : null,
                    Extra = { ["Object ID"] = r.Str("id") ?? $"node/{name}" },
                };
                if (online) _onlineNodes.Add(name);
                else Warn($"Node {name} is {r.Str("status") ?? "offline"}; its details and guests could not be collected.");
                hosts.Add(host);
            }

            foreach (var s in storageRes)
            {
                var type = s.Str("plugintype") ?? (_storageConfig.TryGetValue(s.Str("storage")!, out var c) ? c.Str("type") : null);
                if (type is not null) _storageTypes[(s.Str("node")!, s.Str("storage")!)] = type;
            }

            // Fan out: node details, storage content listings and guests run concurrently (bounded by the API gate).
            _progress?.Report($"Collecting {hosts.Count} node(s) and {guestRes.Count} guest(s)...");
            var nodeTasks = hosts.Where(h => _onlineNodes.Contains(h.Name)).Select(CollectNodeAsync).ToList();
            var contentTasks = PlanContentListings(storageRes).Select(p => ListContentAsync(p.Node, p.Storage, p.Shared)).ToList();
            _guestsTotal = guestRes.Count;
            var guestTasks = guestRes.Select(CollectGuestAsync).ToList();
            await Task.WhenAll(nodeTasks.Concat(contentTasks).Concat(guestTasks)).ConfigureAwait(false);

            var vms = guestTasks.Select(t => t.Result)
                .OrderBy(v => v.Name, StringComparer.OrdinalIgnoreCase).ThenBy(v => v.VmId).ToList();

            ApplyContent(vms);
            ApplyHa(vms, haRes, haStatus, haGroupsTask.Result);
            ApplyMaintenance(hosts, haStatus);
            var datastores = BuildDatastores(storageRes);

            if (cluster is not null)
            {
                cluster.NumHosts = cluster.Nodes.Count;
                cluster.NumEffectiveHosts = cluster.Nodes.Count(n => n.State == "online");
                cluster.NumCpuCores = hosts.Sum(h => h.CpuCores ?? 0);
                cluster.NumCpuThreads = hosts.Sum(h => h.CpuThreads ?? 0);
                cluster.TotalCpuMhz = hosts.Sum(h => (long)(h.CpuMhz ?? 0) * (h.CpuCores ?? 0));
                cluster.TotalMemoryBytes = hosts.Sum(h => h.MemoryBytes ?? 0);
                cluster.HaEnabled = haRes is null ? null : haRes.Value.Items().Any();
                cluster.Extra["Config status"] = cluster.Quorate == false ? "red" : "green";
            }

            var snapshot = new InventorySnapshot
            {
                Source = source,
                Clusters = cluster is null ? [] : [cluster],
                Hosts = hosts,
                VirtualMachines = vms,
                Datastores = datastores,
            };
            AddHealth(snapshot, cluster);
            lock (_warnings) snapshot.Warnings.AddRange(_warnings);
            _progress?.Report($"Collected {hosts.Count} node(s), {vms.Count} guest(s) from {_req.Address}.");
            return snapshot;
        }

        private string? StorageType(string node, string storage) =>
            _storageTypes.TryGetValue((node, storage), out var t) ? t
            : _storageConfig.TryGetValue(storage, out var c) ? c.Str("type")
            : null;

        private bool IsShared(JsonElement storageRes)
        {
            var id = storageRes.Str("storage")!;
            if (storageRes.Bool("shared") is { } s) return s;
            if (_storageConfig.TryGetValue(id, out var cfg) && cfg.Bool("shared") is { } cs) return cs;
            return PveParsing.IsInherentlyShared(storageRes.Str("plugintype") ?? cfg.Str("type"));
        }

        // ------------------------------------------------------------ nodes

        private async Task CollectNodeAsync(HostSystem host)
        {
            var n = Seg(host.Name);
            try
            {
                var status = _api.GetAsync($"/nodes/{n}/status", _ct);
                var dns = _api.TryGetAsync($"/nodes/{n}/dns", _ct);
                var time = _api.TryGetAsync($"/nodes/{n}/time", _ct);
                var network = _api.TryGetAsync($"/nodes/{n}/network", _ct);
                var pci = _api.TryGetAsync($"/nodes/{n}/hardware/pci", _ct);
                _progress?.Report($"Node {host.Name}: status, DNS, network...");

                try
                {
                    PveHostMapper.ApplyStatus(host, await status.ConfigureAwait(false), _now);
                }
                catch (Exception ex) when (PveApiClient.IsSoftFailure(ex, _ct))
                {
                    Warn($"Node {host.Name}: status not available ({ex.Message}).");
                }
                if (await dns.ConfigureAwait(false) is { } d) PveHostMapper.ApplyDns(host, d);
                if (await time.ConfigureAwait(false) is { } t) PveHostMapper.ApplyTime(host, t);
                if (await network.ConfigureAwait(false) is { } net) PveHostMapper.ApplyNetwork(host, net);
                else Warn($"Node {host.Name}: network configuration not readable (needs Sys.Audit).");
                if (await pci.ConfigureAwait(false) is { } p) PveHostMapper.ApplyPci(host, p);
            }
            catch (Exception ex) when (PveApiClient.IsSoftFailure(ex, _ct))
            {
                Warn($"Node {host.Name}: {ex.Message}");
            }
        }

        private static void ApplyMaintenance(List<HostSystem> hosts, JsonElement? haStatus)
        {
            foreach (var lrm in haStatus.Items().Where(e => e.Str("type") == "lrm"))
            {
                var node = lrm.Str("node");
                var text = lrm.Str("status") ?? "";
                if (node is not null && text.Contains("maintenance", StringComparison.OrdinalIgnoreCase)
                    && hosts.FirstOrDefault(h => h.Name == node) is { } h)
                    h.InMaintenance = true;
            }
        }

        // ------------------------------------------------------------ storage

        /// <summary>One content listing per local storage per node, and one per shared storage overall.</summary>
        private IEnumerable<(string Node, string Storage, bool Shared)> PlanContentListings(List<JsonElement> storageRes)
        {
            var sharedDone = new HashSet<string>(StringComparer.Ordinal);
            foreach (var s in storageRes)
            {
                var node = s.Str("node")!;
                var id = s.Str("storage")!;
                var content = s.Str("content") ?? (_storageConfig.TryGetValue(id, out var c) ? c.Str("content") : null) ?? "";
                if (!_onlineNodes.Contains(node) || s.Str("status") is { } st && st != "available") continue;
                if (!content.Split(',').Any(x => x.Trim() is "images" or "rootdir")) continue;
                var shared = IsShared(s);
                if (shared && !sharedDone.Add(id)) continue;
                yield return (node, id, shared);
            }
        }

        private async Task ListContentAsync(string node, string storage, bool shared)
        {
            var list = await _api.TryGetAsync($"/nodes/{Seg(node)}/storage/{Seg(storage)}/content", _ct).ConfigureAwait(false);
            if (list is null) return;
            var scope = shared ? "*" : node;
            lock (_content)
            {
                foreach (var v in list.Items())
                    if (v.Str("volid") is { } volid)
                        _content[$"{scope}|{volid}"] = (v.Long("size"), v.Long("used"), v.Str("format"));
            }
        }

        private void ApplyContent(List<VirtualMachine> vms)
        {
            foreach (var vm in vms)
            foreach (var d in vm.Disks.Where(d => d.Path is not null && d.Datastore is not null))
            {
                if (!_content.TryGetValue($"{vm.Host}|{d.Path}", out var info) && !_content.TryGetValue($"*|{d.Path}", out info))
                    continue;
                if (info.Used is > 0) d.UsedBytes ??= info.Used;
                d.CapacityBytes ??= info.Size;
                d.Format ??= info.Format;
                if (d.Thin is null && d.Format is not null) d.Thin = PveParsing.IsThin(StorageType(vm.Host, d.Datastore!), d.Format);
            }
        }

        private List<Datastore> BuildDatastores(List<JsonElement> storageRes)
        {
            var result = new List<Datastore>();
            foreach (var group in storageRes.GroupBy(s => s.Str("storage")!, StringComparer.Ordinal))
            {
                var entries = group.ToList();
                var shared = IsShared(entries[0]);
                List<List<JsonElement>> buckets = shared ? [entries] : entries.Select(e => new List<JsonElement> { e }).ToList();
                foreach (var bucket in buckets)
                {
                    var best = bucket.FirstOrDefault(e => e.Str("status") == "available" && e.Long("maxdisk") is > 0);
                    if (best.ValueKind == JsonValueKind.Undefined) best = bucket[0];
                    _storageConfig.TryGetValue(group.Key, out var cfg);
                    var capacity = best.Long("maxdisk");
                    var used = best.Long("disk");
                    var accessible = bucket.Any(e => e.Str("status") == "available");
                    var ds = new Datastore
                    {
                        Name = group.Key,
                        Platform = Platform.Proxmox,
                        SourceAddress = _req.Address,
                        Cluster = _cluster,
                        Type = best.Str("plugintype") ?? cfg.Str("type"),
                        Address = cfg.ValueKind == JsonValueKind.Object ? PveHostMapper.StorageAddress(cfg) : null,
                        Accessible = accessible,
                        Shared = shared,
                        CapacityBytes = capacity is > 0 ? capacity : null,
                        FreeBytes = capacity is > 0 && used is { } u ? Math.Max(0, capacity.Value - u) : null,
                        Content = best.Str("content") ?? cfg.Str("content"),
                        Hosts = bucket.Select(e => e.Str("node")!).Distinct().OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList(),
                    };
                    ds.Extra["Object ID"] = shared ? $"storage/{group.Key}" : best.Str("id") ?? $"storage/{bucket[0].Str("node")}/{group.Key}";
                    ds.Extra["Config status"] = accessible ? "green" : "red";
                    if (shared && bucket.Any(e => e.Str("status") != "available"))
                        ds.Extra["Config status"] = "yellow";
                    result.Add(ds);
                }
            }
            return result.OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase).ThenBy(d => d.Hosts.FirstOrDefault()).ToList();
        }

        // ------------------------------------------------------------ guests

        private VirtualMachine BaseGuest(JsonElement r)
        {
            var node = r.Str("node")!;
            var vmid = r.Long("vmid")!.Value;
            var lxc = r.Str("type") == "lxc";
            var vm = new VirtualMachine
            {
                Name = r.Str("name") ?? $"{(lxc ? "CT" : "VM")} {vmid}",
                VmId = vmid.ToString(),
                Kind = lxc ? GuestKind.Container : GuestKind.VirtualMachine,
                IsTemplate = r.Bool("template") == true,
                Platform = Platform.Proxmox,
                Host = node,
                Cluster = _cluster,
                Datacenter = _req.Group,
                SourceAddress = _req.Address,
                SourceProduct = _product,
                SourceApiVersion = _apiVersion,
                ResourcePool = r.Str("pool"),
                Tags = PveParsing.SplitTags(r.Str("tags")),
                RawState = r.Str("status"),
                PowerState = PveParsing.MapPowerState(r.Str("status")),
                Status = "green",
                CpuCount = r.Int("maxcpu") ?? 0,
                MemoryMiB = r.Long("maxmem") is { } mm ? mm / PveParsing.MiB : 0,
                ConfigPath = $"/etc/pve/nodes/{node}/{(lxc ? "lxc" : "qemu-server")}/{vmid}.conf",
                OwnerNode = node,
            };
            if (r.Str("hastate") is { Length: > 0 } hs)
            {
                vm.HaProtected = true;
                vm.HaState = hs;
            }
            if (vm.IsTemplate) vm.PowerState = PowerState.PoweredOff;
            vm.Extra["Object ID"] = r.Str("id") ?? $"{r.Str("type")}/{vmid}";
            return vm;
        }

        private async Task<VirtualMachine> CollectGuestAsync(JsonElement r)
        {
            var vm = BaseGuest(r);
            try
            {
                if (!_onlineNodes.Contains(vm.Host))
                {
                    vm.PowerState = PowerState.Unknown;
                    vm.RawState = "unknown";
                    vm.Status = "gray";
                    return vm;
                }

                if (vm.Kind == GuestKind.Container) await CollectLxcAsync(vm).ConfigureAwait(false);
                else await CollectQemuAsync(vm).ConfigureAwait(false);
            }
            catch (Exception ex) when (PveApiClient.IsSoftFailure(ex, _ct))
            {
                vm.Status = "gray";
                Warn($"{(vm.Kind == GuestKind.Container ? "CT" : "VM")} {vm.VmId} ({vm.Name}) on {vm.Host}: {ex.Message}");
            }
            finally
            {
                var done = Interlocked.Increment(ref _guestsDone);
                if (done % 10 == 0 || done == _guestsTotal) _progress?.Report($"Collected {done}/{_guestsTotal} guests...");
            }
            return vm;
        }

        private async Task CollectQemuAsync(VirtualMachine vm)
        {
            var basePath = $"/nodes/{Seg(vm.Host)}/qemu/{vm.VmId}";
            var cfgTask = _api.GetAsync($"{basePath}/config", _ct);
            var stTask = vm.IsTemplate ? Task.FromResult<JsonElement?>(null) : _api.TryGetAsync($"{basePath}/status/current", _ct);
            var snapTask = vm.IsTemplate ? Task.FromResult<JsonElement?>(null) : _api.TryGetAsync($"{basePath}/snapshot", _ct);

            var cfg = await cfgTask.ConfigureAwait(false);
            PveGuestMapper.ApplyQemuConfig(vm, cfg, s => StorageType(vm.Host, s));
            var agentEnabled = PveProps.Parse(cfg.Str("agent")) is var ap && (PveParsing.IsTrue(ap.Positional) || PveParsing.IsTrue(ap["enabled"]));

            if (await stTask.ConfigureAwait(false) is { } st)
                PveGuestMapper.ApplyQemuStatus(vm, st, hibernated: cfg.Has("vmstate"), _now);
            else if (!vm.IsTemplate)
                Warn($"VM {vm.VmId} ({vm.Name}): runtime status not available.");
            if (await snapTask.ConfigureAwait(false) is { } snaps) vm.Snapshots = PveGuestMapper.ParseSnapshots(snaps);

            vm.ToolsRunning = false;
            vm.Heartbeat = "gray";
            vm.ToolsStatus = agentEnabled ? "guestToolsNotRunning" : "agentDisabled";

            if (agentEnabled && vm.PowerState == PowerState.PoweredOn)
                await CollectAgentAsync(vm, basePath).ConfigureAwait(false);
        }

        private async Task CollectAgentAsync(VirtualMachine vm, string basePath)
        {
            var timeout = _owner.AgentTimeout;
            // "info" is the liveness probe: when it fails the agent is not running, so skip the other calls.
            var info = await _api.TryGetAsync($"{basePath}/agent/info", _ct, timeout).ConfigureAwait(false);
            if (info is null) return;

            vm.ToolsRunning = true;
            vm.ToolsStatus = "guestToolsRunning";
            vm.Heartbeat = "green";
            vm.ToolsVersion = PveGuestMapper.AgentResult(info.Value).Str("version");

            var os = _api.TryGetAsync($"{basePath}/agent/get-osinfo", _ct, timeout);
            var host = _api.TryGetAsync($"{basePath}/agent/get-host-name", _ct, timeout);
            var net = _api.TryGetAsync($"{basePath}/agent/network-get-interfaces", _ct, timeout);
            var fs = _api.TryGetAsync($"{basePath}/agent/get-fsinfo", _ct, timeout);
            await Task.WhenAll(os, host, net, fs).ConfigureAwait(false);
            try
            {
                if (os.Result is { } o) PveGuestMapper.ApplyAgentOsInfo(vm, o);
                if (host.Result is { } h) PveGuestMapper.ApplyAgentHostName(vm, h);
                if (net.Result is { } n) PveGuestMapper.ApplyAgentNetwork(vm, n);
                if (fs.Result is { } f) PveGuestMapper.ApplyAgentFsInfo(vm, f);
            }
            catch (Exception ex) when (ex is InvalidOperationException or FormatException or KeyNotFoundException)
            {
                Warn($"VM {vm.VmId} ({vm.Name}): unexpected guest agent data ({ex.Message}).");
            }
        }

        private async Task CollectLxcAsync(VirtualMachine vm)
        {
            var basePath = $"/nodes/{Seg(vm.Host)}/lxc/{vm.VmId}";
            var cfgTask = _api.GetAsync($"{basePath}/config", _ct);
            var stTask = vm.IsTemplate ? Task.FromResult<JsonElement?>(null) : _api.TryGetAsync($"{basePath}/status/current", _ct);
            var snapTask = vm.IsTemplate ? Task.FromResult<JsonElement?>(null) : _api.TryGetAsync($"{basePath}/snapshot", _ct);

            PveGuestMapper.ApplyLxcConfig(vm, await cfgTask.ConfigureAwait(false), s => StorageType(vm.Host, s));
            if (await stTask.ConfigureAwait(false) is { } st) PveGuestMapper.ApplyLxcStatus(vm, st, _now);
            if (await snapTask.ConfigureAwait(false) is { } snaps) vm.Snapshots = PveGuestMapper.ParseSnapshots(snaps);

            if (vm.PowerState == PowerState.PoweredOn
                && await _api.TryGetAsync($"{basePath}/interfaces", _ct, _owner.AgentTimeout).ConfigureAwait(false) is { } ifs)
                PveGuestMapper.ApplyLxcInterfaces(vm, ifs);
        }

        // ------------------------------------------------------------ HA

        private static void ApplyHa(List<VirtualMachine> vms, JsonElement? haRes, JsonElement? haStatus, JsonElement? haGroups)
        {
            if (haRes is null) return;
            var resources = haRes.Items().Where(r => r.Str("sid") is not null).ToDictionary(r => r.Str("sid")!, StringComparer.Ordinal);
            var services = haStatus.Items().Where(s => s.Str("type") == "service" && s.Str("sid") is not null)
                .GroupBy(s => s.Str("sid")!).ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
            var groups = haGroups.Items().Where(g => g.Str("group") is not null)
                .GroupBy(g => g.Str("group")!).ToDictionary(g => g.Key, g => g.First().Str("nodes"), StringComparer.Ordinal);

            foreach (var vm in vms)
            {
                var sid = $"{(vm.Kind == GuestKind.Container ? "ct" : "vm")}:{vm.VmId}";
                if (!resources.TryGetValue(sid, out var res))
                {
                    vm.HaProtected = false;
                    vm.HaState = null;
                    continue;
                }
                vm.HaProtected = true;
                vm.HaState = res.Str("state") ?? vm.HaState;
                if (services.TryGetValue(sid, out var svc))
                {
                    vm.HaState = svc.Str("state") ?? vm.HaState;
                    vm.OwnerNode = svc.Str("node") ?? vm.OwnerNode;
                }
                if (res.Str("group") is { } group)
                {
                    vm.FailoverPriority = group;
                    if (groups.TryGetValue(group, out var nodes) && nodes is not null)
                    {
                        // "pve1:2,pve2:1" → ordered by priority (higher first).
                        vm.PreferredOwners = nodes.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                            .Select(n => n.Split(':'))
                            .Select(p => (Node: p[0], Prio: p.Length > 1 && int.TryParse(p[1], out var pr) ? pr : 0))
                            .OrderByDescending(p => p.Prio).Select(p => p.Node).ToList();
                    }
                }
            }
        }

        // ------------------------------------------------------------ health

        private void AddHealth(InventorySnapshot snapshot, ClusterInfo? cluster)
        {
            if (cluster?.Quorate == false)
                snapshot.Health.Add(new HealthItem
                {
                    Name = cluster.Name,
                    Message = "Cluster is not quorate",
                    Severity = HealthSeverity.Error,
                    SourceAddress = _req.Address,
                });
            foreach (var h in snapshot.Hosts.Where(h => h.Status == "red"))
                snapshot.Health.Add(new HealthItem
                {
                    Name = h.Name,
                    Message = "Node is offline",
                    Severity = HealthSeverity.Error,
                    SourceAddress = _req.Address,
                });
            foreach (var vm in snapshot.VirtualMachines.Where(v => v.HaState is "error" or "fence"))
                snapshot.Health.Add(new HealthItem
                {
                    Name = vm.Name,
                    Message = $"HA resource is in state '{vm.HaState}'",
                    Severity = HealthSeverity.Error,
                    SourceAddress = _req.Address,
                });
        }
    }
}
