using System.Text.Json;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>
/// Maps the JSON produced by Collect-HyperV.ps1 (one host) or by the launcher (several cluster nodes) to an
/// <see cref="InventorySnapshot"/>. Also used to import a JSON file from a manual run of the script.
/// </summary>
public static class HyperVJsonMapper
{
    public const string ProductName = "Microsoft Hyper-V";
    public const int SupportedSchemaVersion = 1;

    private const string ImportHint =
        "Expected the JSON written by Collect-HyperV.ps1 (e.g. .\\Collect-HyperV.ps1 -OutFile hyperv.json).";

    /// <summary>Maps collection JSON text.</summary>
    public static InventorySnapshot Map(string json, ConnectionRequest request)
    {
        ArgumentNullException.ThrowIfNull(json);
        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(json.TrimStart('\uFEFF'), new JsonDocumentOptions { MaxDepth = 256 });
        }
        catch (JsonException ex)
        {
            throw new CollectionException(CollectionFailure.Protocol, "The Hyper-V collection data is not valid JSON: " + ex.Message, ImportHint, ex);
        }
        using (doc) return Map(doc.RootElement, request);
    }

    /// <summary>Reads (UTF-8 or UTF-16, with or without BOM) and maps a collection JSON file.</summary>
    public static InventorySnapshot MapFile(string path, ConnectionRequest request) => Map(File.ReadAllText(path), request);

    /// <summary>Maps an already-parsed collection document (single host, multi-node, or a launcher envelope).</summary>
    public static InventorySnapshot Map(JsonElement root, ConnectionRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (root.Prop("ok") is { } okEl)
        {
            if (okEl.ValueKind != JsonValueKind.True)
                throw HyperVErrorMapper.FromEnvelope(root.Str("stage"), root.Str("error"), root.Str("category"), request);
            root = root.Prop("data") ?? throw new CollectionException(CollectionFailure.Protocol, "The collection result contained no data.");
        }

        var warnings = new List<string>();
        List<JsonElement> nodeDocs;
        if (root.Prop("nodes") is not null)
        {
            nodeDocs = root.Arr("nodes").Where(n => n.Prop("host") is not null).ToList();
            warnings.AddRange(root.Strs("warnings"));
        }
        else if (root.Prop("host") is not null)
        {
            nodeDocs = [root];
        }
        else
        {
            throw new CollectionException(CollectionFailure.Protocol, "This is not a Hypervisor Explorer Hyper-V collection.", ImportHint);
        }
        if (nodeDocs.Count == 0)
            throw new CollectionException(CollectionFailure.Protocol, "The Hyper-V collection contained no hosts.", ImportHint);

        // One document per host (a node may be reached twice, e.g. via its name and via the cluster name).
        nodeDocs = nodeDocs
            .GroupBy(n => n.Prop("host").Str("name") ?? n.Str("computerName") ?? "", StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .ToList();

        foreach (var n in nodeDocs)
        {
            if (n.Int("schemaVersion") is { } v && v > SupportedSchemaVersion)
                warnings.Add($"{n.Str("computerName")}: collection schema v{v} is newer than this version understands (v{SupportedSchemaVersion}); some data may be missing.");
            warnings.AddRange(n.Strs("warnings"));
        }

        var primary = nodeDocs[0];
        var primaryHost = primary.Prop("host")!.Value;
        var os = primaryHost.Prop("os");
        var caption = os.Str("caption");
        var osVersion = os.Str("version");
        var build = os.Str("buildNumber");
        var clusterEl = nodeDocs.Select(n => n.Prop("cluster")).FirstOrDefault(c => c is not null);
        var clusterName = clusterEl.Str("name");

        var source = new Source
        {
            Address = request.Address,
            Platform = Platform.HyperV,
            Group = request.Group,
            ProductName = ProductName,
            Version = osVersion ?? "",
            Build = build,
            ApiVersion = "WinRM",
            Vendor = "Microsoft Corporation",
            OsType = "Windows",
            CollectedAt = primary.Date("collectedAt") ?? DateTimeOffset.Now,
        };
        if (caption is not null)
            source.Extra["Fullname"] = $"{ProductName} ({caption}) {osVersion} build-{build}".Replace("  ", " ");
        source.Extra["Product line"] = clusterName is null ? "Hyper-V" : "Hyper-V Failover Cluster";

        var ctx = new MapContext(request, clusterName, $"{ProductName} {caption ?? osVersion}".Trim());
        var snapshot = new InventorySnapshot { Source = source };

        // Cluster-wide shared storage first so VM disks can resolve against it.
        var csvDatastores = new List<(Datastore Ds, List<string> Prefixes)>();
        var clusterNodeNames = clusterEl.Arr("nodes").Select(n => n.Str("name")).Where(n => n is not null).Select(n => n!).ToList();
        if (clusterEl is { } ce)
        {
            foreach (var csv in ce.Arr("sharedVolumes"))
            {
                var path = csv.Str("friendlyVolumeName");
                var name = csv.Str("name") ?? path ?? "CSV";
                if (csvDatastores.Any(c => c.Ds.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) && path is not null)
                    name = $"{name} ({path})";
                var fs = csv.Str("fileSystem");
                var ds = new Datastore
                {
                    Name = name,
                    Platform = Platform.HyperV,
                    SourceAddress = request.Address,
                    Cluster = clusterName,
                    Type = fs is null ? "CSV" : $"CSV ({fs})",
                    Accessible = !string.Equals(csv.Str("state"), "Offline", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(csv.Str("state"), "Failed", StringComparison.OrdinalIgnoreCase),
                    Shared = true,
                    CapacityBytes = csv.Long("size"),
                    FreeBytes = csv.Long("freeSpace"),
                    Content = path,
                    Hosts = [.. clusterNodeNames],
                };
                if (csv.Str("ownerNode") is { } owner) ds.Extra["Owner node"] = owner;
                if (csv.Bool("redirectedAccess") == true)
                    snapshot.Health.Add(Health(ds.Name, "Cluster Shared Volume is in redirected access mode", HealthSeverity.Warning, request));
                if (csv.Bool("maintenanceMode") == true)
                    snapshot.Health.Add(Health(ds.Name, "Cluster Shared Volume is in maintenance mode", HealthSeverity.Warning, request));
                csvDatastores.Add((ds, path is null ? [] : [path]));
            }
        }

        // Cluster VM roles, by VmId and by role name.
        var groupsById = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        var groupsByName = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var g in clusterEl.Arr("vmGroups"))
        {
            if (g.Str("vmId") is { } id) groupsById.TryAdd(NormalizeGuid(id), g);
            if (g.Str("name") is { } gn) groupsByName.TryAdd(gn, g);
        }
        var nodeStates = clusterEl.Arr("nodes")
            .Where(n => n.Str("name") is not null)
            .GroupBy(n => n.Str("name")!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().Str("state"), StringComparer.OrdinalIgnoreCase);

        var smbDatastores = new Dictionary<string, Datastore>(StringComparer.OrdinalIgnoreCase);

        foreach (var node in nodeDocs)
        {
            var hostEl = node.Prop("host")!.Value;
            var inCluster = clusterName is not null;
            var host = MapHost(hostEl, ctx, inCluster ? clusterName : null);
            if (nodeStates.TryGetValue(host.Name, out var nodeState) && string.Equals(nodeState, "Paused", StringComparison.OrdinalIgnoreCase))
            {
                host.InMaintenance = true;
                host.Status = "yellow";
            }
            snapshot.Hosts.Add(host);

            // Local volumes -> datastores; mount table = local volumes + CSVs.
            var mounts = new List<KeyValuePair<string, string>>();
            foreach (var vol in hostEl.Arr("volumes"))
            {
                var paths = vol.Strs("paths");
                if (paths.Count == 0) continue;
                var label = vol.Str("label");
                var mount = DisplayMount(paths[0]);
                var ds = new Datastore
                {
                    Name = label is null ? $"{host.Name} {mount}" : $"{host.Name} {mount} ({label})",
                    Platform = Platform.HyperV,
                    SourceAddress = request.Address,
                    Cluster = host.Cluster,
                    Type = vol.Str("fileSystem") ?? "Local",
                    Accessible = !string.Equals(vol.Str("healthStatus"), "Unhealthy", StringComparison.OrdinalIgnoreCase),
                    Shared = false,
                    CapacityBytes = vol.Long("size"),
                    FreeBytes = vol.Long("sizeRemaining"),
                    Content = string.Join(", ", paths),
                    Hosts = [host.Name],
                };
                if (vol.Str("healthStatus") is { } hs) ds.Extra["Config status"] = HealthColour(hs);
                snapshot.Datastores.Add(ds);
                mounts.AddRange(paths.Select(p => new KeyValuePair<string, string>(p, ds.Name)));
            }
            foreach (var (ds, prefixes) in csvDatastores)
                mounts.AddRange(prefixes.Select(p => new KeyValuePair<string, string>(p, ds.Name)));

            var collectedAt = node.Date("collectedAt") ?? source.CollectedAt;
            foreach (var vmEl in node.Arr("vms"))
            {
                try
                {
                    var vm = MapVm(vmEl, host, collectedAt, ctx, mounts, smbDatastores, groupsById, groupsByName);
                    snapshot.VirtualMachines.Add(vm);
                }
                catch (Exception ex) when (ex is InvalidOperationException or FormatException or JsonException)
                {
                    warnings.Add($"{host.Name}: VM '{vmEl.Str("name")}' could not be mapped: {ex.Message}");
                }
            }
        }

        snapshot.Datastores.AddRange(csvDatastores.Select(c => c.Ds));
        snapshot.Datastores.AddRange(smbDatastores.Values);

        if (clusterEl is { } cel)
            snapshot.Clusters.Add(MapCluster(cel, snapshot, request, nodeDocs.Count));

        foreach (var w in warnings.Distinct())
        {
            snapshot.Warnings.Add(w);
            var colon = w.IndexOf(": ", StringComparison.Ordinal);
            var subject = colon is > 0 and < 64 && !w[..colon].Contains(' ') ? w[..colon] : request.Address;
            snapshot.Health.Add(Health(subject, colon is > 0 and < 64 && subject != request.Address ? w[(colon + 2)..] : w, HealthSeverity.Warning, request));
        }

        if (clusterEl is { } cl2)
        {
            foreach (var n in cl2.Arr("nodes"))
            {
                var st = n.Str("state");
                if (st is not null && !st.Equals("Up", StringComparison.OrdinalIgnoreCase))
                    snapshot.Health.Add(Health(n.Str("name") ?? "?", $"Cluster node is {st}",
                        st.Equals("Paused", StringComparison.OrdinalIgnoreCase) ? HealthSeverity.Info : HealthSeverity.Warning, request));
            }
            foreach (var g in cl2.Arr("vmGroups"))
            {
                var st = g.Str("state");
                if (st is not null && (st.Equals("Failed", StringComparison.OrdinalIgnoreCase) || st.Equals("PartialOnline", StringComparison.OrdinalIgnoreCase)))
                    snapshot.Health.Add(Health(g.Str("name") ?? "?", $"Cluster role is {st}", HealthSeverity.Error, request));
            }
        }

        return snapshot;
    }

    private sealed record MapContext(ConnectionRequest Request, string? ClusterName, string SourceProduct);

    private static HealthItem Health(string name, string message, HealthSeverity severity, ConnectionRequest request) =>
        new() { Name = name, Message = message, Severity = severity, SourceAddress = request.Address };

    // ------------------------------------------------------------------ host

    private static HostSystem MapHost(JsonElement h, MapContext ctx, string? cluster)
    {
        var os = h.Prop("os");
        var bios = h.Prop("bios");
        var procs = h.Arr("processors").ToList();
        var name = h.Str("name") ?? h.Str("dnsHostName") ?? "unknown";

        var cores = procs.Sum(p => p.Long("numberOfCores") ?? 0);
        var threads = procs.Sum(p => p.Long("numberOfLogicalProcessors") ?? 0);
        if (threads == 0) threads = h.Long("logicalProcessorCount") ?? 0;
        var loads = procs.Select(p => p.Long("loadPercentage")).Where(l => l is not null).Select(l => (double)l!.Value).ToList();

        long? memUsed = null;
        if (os.Long("totalVisibleMemoryKb") is { } totalKb && os.Long("freePhysicalMemoryKb") is { } freeKb)
            memUsed = (totalKb - freeKb) * 1024;

        var caption = os.Str("caption");
        var version = os.Str("version");
        var host = new HostSystem
        {
            Name = name,
            Platform = Platform.HyperV,
            SourceAddress = ctx.Request.Address,
            Datacenter = ctx.Request.Group,
            Cluster = cluster,
            Status = "green",
            CpuModel = procs.Select(p => CollapseSpaces(p.Str("name"))).FirstOrDefault(n => n is not null),
            CpuMhz = procs.Select(p => p.Int("maxClockSpeed")).FirstOrDefault(m => m is not null),
            CpuSockets = procs.Count > 0 ? procs.Count : h.Int("numberOfProcessors"),
            CoresPerSocket = procs.Count > 0 ? procs[0].Int("numberOfCores") : null,
            CpuCores = cores > 0 ? (int)cores : null,
            CpuThreads = threads > 0 ? (int)threads : null,
            HyperThreadingActive = cores > 0 && threads > 0 ? threads > cores : null,
            CpuUsagePercent = loads.Count > 0 ? loads.Average() : null,
            MemoryBytes = h.Long("memoryCapacity") ?? h.Long("totalPhysicalMemory"),
            MemoryUsedBytes = memUsed,
            Version = caption is null ? version : $"{caption} {version}".Trim(),
            KernelVersion = version is null ? null : os.Str("buildNumber") is { } b && !version.EndsWith(b, StringComparison.Ordinal) ? $"{version}.{b}" : version,
            BootTime = os.Date("lastBootUpTime"),
            Vendor = h.Str("manufacturer"),
            Model = h.Str("model"),
            SerialNumber = bios.Str("serialNumber"),
            BiosVendor = bios.Str("manufacturer"),
            BiosVersion = bios.Str("version"),
            BiosDate = bios.Date("releaseDate"),
            Uuid = h.Str("uuid")?.ToLowerInvariant(),
            DnsServers = h.Strs("dnsServers"),
            Domain = h.Str("domain"),
            DnsSearch = h.Strs("dnsSearch"),
            NtpServers = ParseNtp(h.Str("ntpSource")),
            TimeZone = h.Prop("timeZone").Str("id"),
        };
        if (h.Prop("timeZone") is { } tz)
        {
            if (tz.Str("displayName") is { } dn) host.Extra["Time Zone Name"] = dn;
            if (tz.Long("baseUtcOffsetMinutes") is { } off) host.Extra["GMT Offset"] = off / 60.0;
        }
        if (h.Bool("liveMigrationEnabled") is { } lm) host.Extra["VMotion support"] = lm;

        MapHostNetworking(h, host);
        MapStorageAdapters(h, host);
        return host;
    }

    private static void MapHostNetworking(JsonElement h, HostSystem host)
    {
        var nics = h.Arr("nics").ToList();
        var teams = h.Arr("lbfoTeams").ToList();

        // Resolve a switch uplink description to physical NIC names (directly, or via an LBFO team NIC).
        List<string> Uplinks(string description)
        {
            var direct = nics.Where(n => string.Equals(n.Str("interfaceDescription"), description, StringComparison.OrdinalIgnoreCase))
                .Select(n => n.Str("name")!).Where(n => n is not null).ToList();
            if (direct.Count > 0) return direct;
            var team = teams.FirstOrDefault(t => t.Strs("teamNicDescriptions").Contains(description, StringComparer.OrdinalIgnoreCase));
            if (team.ValueKind == JsonValueKind.Object)
            {
                var members = team.Strs("members");
                if (members.Count > 0) return members;
            }
            return [description];
        }

        foreach (var s in h.Arr("switches"))
        {
            var type = s.Str("switchType");
            if (s.Bool("embeddedTeamingEnabled") == true) type = $"{type} (SET)";
            var sw = new VirtualSwitch
            {
                Name = s.Str("name") ?? "",
                Type = type,
                Uplinks = s.Strs("netAdapterInterfaceDescriptions").SelectMany(Uplinks).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                Notes = s.Str("notes"),
            };
            if (s.Bool("allowManagementOS") is { } mos) sw.Extra["Allow management OS"] = mos;
            if (s.Str("bandwidthReservationMode") is { } bw) sw.Extra["Traffic Shaping"] = bw;
            host.Switches.Add(sw);
        }

        foreach (var n in nics)
        {
            var name = n.Str("name") ?? "";
            var speed = n.Long("speedBps");
            var up = string.Equals(n.Str("status"), "Up", StringComparison.OrdinalIgnoreCase);
            host.Nics.Add(new HostNic
            {
                Name = name,
                Description = n.Str("interfaceDescription"),
                Driver = n.Str("driverFileName") ?? n.Str("driverDescription"),
                SpeedMbps = up && speed is > 0 and < 10_000_000_000_000 ? speed / 1_000_000 : (up ? null : 0),
                FullDuplex = n.Bool("fullDuplex"),
                Mac = FormatMac(n.Str("macAddress")),
                Switch = host.Switches.FirstOrDefault(sw => sw.Uplinks.Contains(name, StringComparer.OrdinalIgnoreCase))?.Name,
                Pci = n.Str("pci"),
                Status = n.Str("status"),
            });
        }

        var mgmt = h.Arr("managementAdapters").ToList();
        foreach (var m in mgmt)
        {
            host.PortGroups.Add(new PortGroup
            {
                Name = m.Str("name") ?? "",
                Switch = m.Str("switchName"),
                Vlan = AccessVlan(m.Str("vlanMode"), m.Int("accessVlanId")) ?? 0,
            });
        }

        foreach (var ip in h.Arr("ipInterfaces"))
        {
            var mac = FormatMac(ip.Str("macAddress"));
            var alias = ip.Str("interfaceAlias") ?? "";
            var mg = mgmt.FirstOrDefault(m => mac is not null && FormatMac(m.Str("macAddress")) == mac);
            string? portGroup = mg.ValueKind == JsonValueKind.Object ? mg.Str("name") : null;
            if (portGroup is null && alias.StartsWith("vEthernet (", StringComparison.OrdinalIgnoreCase) && alias.EndsWith(')'))
                portGroup = alias[11..^1];
            var v4 = ip.Arr("ipv4").ToList();
            var v6 = ip.Arr("ipv6").ToList();
            host.IpInterfaces.Add(new HostIpInterface
            {
                Name = alias,
                PortGroup = portGroup,
                Mac = mac,
                Dhcp = ip.Bool("dhcp"),
                Ipv4 = v4.Count == 0 ? null : string.Join(", ", v4.Select(a => a.Str("address")).Where(a => a is not null)),
                SubnetMask = v4.Count == 0 ? null : PrefixToMask(v4[0].Long("prefixLength")),
                Gateway = ip.Str("gateway"),
                Ipv6 = v6.Count == 0 ? null : string.Join(", ", v6.Select(a => a.Long("prefixLength") is { } p ? $"{a.Str("address")}/{p}" : a.Str("address"))),
                Mtu = ip.Int("mtu"),
            });
        }
    }

    private static void MapStorageAdapters(JsonElement h, HostSystem host)
    {
        foreach (var p in h.Arr("initiatorPorts"))
        {
            var type = p.Str("connectionType");
            var node = p.Str("nodeAddress");
            var port = p.Str("portAddress");
            var isFc = type?.Contains("Fibre", StringComparison.OrdinalIgnoreCase) == true;
            host.StorageAdapters.Add(new StorageAdapter
            {
                Device = p.Str("instanceName") ?? node ?? "",
                Type = type,
                Status = p.Str("operationalStatus"),
                Wwn = isFc ? string.Join(" ", new[] { FormatWwn(node), FormatWwn(port) }.Where(x => x is not null)) : node,
            });
        }
        foreach (var c in h.Arr("scsiControllers"))
        {
            host.StorageAdapters.Add(new StorageAdapter
            {
                Device = c.Str("name") ?? "",
                Type = "SCSI",
                Status = c.Str("status"),
                Driver = c.Str("driverName"),
                Model = c.Str("name"),
            });
        }
    }

    // ------------------------------------------------------------------ cluster

    private static ClusterInfo MapCluster(JsonElement c, InventorySnapshot snapshot, ConnectionRequest request, int collectedNodes)
    {
        var networks = c.Arr("networks").ToList();
        var netRoles = networks
            .Where(n => n.Str("name") is not null)
            .GroupBy(n => n.Str("name")!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => RoleText(g.First()), StringComparer.OrdinalIgnoreCase);
        var netIfs = c.Arr("networkInterfaces").ToList();

        var info = new ClusterInfo
        {
            Name = c.Str("name") ?? "",
            Platform = Platform.HyperV,
            SourceAddress = request.Address,
            Datacenter = request.Group,
            Id = c.Str("id"),
            Domain = c.Str("domain"),
            QuorumType = c.Str("quorumType"),
            QuorumWitness = c.Str("quorumWitnessType") is { } wt && c.Str("quorumWitness") is { } w && !w.Equals(wt, StringComparison.OrdinalIgnoreCase)
                ? $"{w} ({wt})" : c.Str("quorumWitness") ?? c.Str("quorumWitnessType"),
            FunctionalLevel = FunctionalLevelText(c.Int("functionalLevel")),
            HaEnabled = true,
            Quorate = true,
            DrsEnabled = false,
        };

        foreach (var n in c.Arr("nodes"))
        {
            var name = n.Str("name") ?? "";
            var addr = netIfs
                .Where(i => string.Equals(i.Str("node"), name, StringComparison.OrdinalIgnoreCase) && i.Str("address") is not null)
                .OrderBy(i => netRoles.TryGetValue(i.Str("network") ?? "", out var r) && r == "Cluster and client" ? 0 : 1)
                .Select(i => i.Str("address"))
                .FirstOrDefault();
            var node = new ClusterNode
            {
                Name = name,
                State = n.Str("state"),
                DrainStatus = n.Str("drainStatus"),
                Id = n.Str("id"),
                Address = addr,
                Votes = n.Int("nodeWeight"),
            };
            if (n.Int("dynamicWeight") is { } dw) node.Extra["Dynamic vote"] = dw;
            if (n.Str("model") is { } model) node.Extra["Model"] = model;
            info.Nodes.Add(node);
        }

        foreach (var n in networks)
        {
            var prefix = n.Strs("ipv4PrefixLengths").FirstOrDefault();
            var address = n.Str("address");
            if (address is not null && prefix is not null) address = $"{address}/{prefix}";
            else if (address is not null && n.Str("addressMask") is { } mask) address = $"{address} / {mask}";
            var metric = n.Long("metric")?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (metric is not null && n.Bool("autoMetric") == true) metric += " (auto)";
            info.Networks.Add(new ClusterNetwork
            {
                Name = n.Str("name") ?? "",
                Role = RoleText(n),
                Address = address,
                State = n.Str("state"),
                Metric = metric,
            });
        }

        var hosts = snapshot.Hosts.Where(h => string.Equals(h.Cluster, info.Name, StringComparison.OrdinalIgnoreCase)).ToList();
        info.NumHosts = info.Nodes.Count > 0 ? info.Nodes.Count : hosts.Count;
        info.NumEffectiveHosts = info.Nodes.Count > 0
            ? info.Nodes.Count(n => string.Equals(n.State, "Up", StringComparison.OrdinalIgnoreCase))
            : hosts.Count;
        info.NumCpuCores = hosts.Sum(h => h.CpuCores ?? 0);
        info.NumCpuThreads = hosts.Sum(h => h.CpuThreads ?? 0);
        info.TotalCpuMhz = hosts.Sum(h => (long)(h.CpuMhz ?? 0) * (h.CpuCores ?? 0));
        info.TotalMemoryBytes = hosts.Sum(h => h.MemoryBytes ?? 0);
        info.OverallStatus = info.Nodes.Any(n => string.Equals(n.State, "Down", StringComparison.OrdinalIgnoreCase)) ? "yellow" : "green";
        if (info.Nodes.Count > 0 && collectedNodes < info.Nodes.Count)
            info.Extra["Config status"] = $"{collectedNodes} of {info.Nodes.Count} nodes collected";
        if (c.Long("s2dEnabled") is { } s2d) info.Extra["Storage Spaces Direct"] = s2d != 0;
        return info;
    }

    private static string? RoleText(JsonElement n)
    {
        var role = n.Str("role");
        var value = n.Long("roleValue");
        return (role, value) switch
        {
            ("ClusterAndClient", _) or (_, 3) or ("3", _) => "Cluster and client",
            ("Cluster", _) or (_, 1) or ("1", _) => "Cluster only",
            ("None", _) or (_, 0) or ("0", _) => "None",
            _ => role,
        };
    }

    /// <summary>ClusterFunctionalLevel number to text.</summary>
    public static string? FunctionalLevelText(int? level) => level switch
    {
        null => null,
        8 => "8 (Windows Server 2012 R2)",
        9 => "9 (Windows Server 2016)",
        10 => "10 (Windows Server 2019)",
        11 => "11 (Windows Server 2022)",
        12 => "12 (Windows Server 2025)",
        _ => level.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    // ------------------------------------------------------------------ VMs

    private static VirtualMachine MapVm(
        JsonElement v, HostSystem host, DateTimeOffset collectedAt, MapContext ctx,
        List<KeyValuePair<string, string>> mounts, Dictionary<string, Datastore> smb,
        Dictionary<string, JsonElement> groupsById, Dictionary<string, JsonElement> groupsByName)
    {
        var id = v.Str("id") is { } rawId ? NormalizeGuid(rawId) : null;
        var state = v.Str("state");
        var power = MapPowerState(state);
        var gen = v.Int("generation");
        var guest = v.Prop("guest");
        var proc = v.Prop("processor");
        var fw = v.Prop("firmware");
        var uptime = v.Long("uptimeSeconds") is { } up and > 0 ? TimeSpan.FromSeconds(up) : (TimeSpan?)null;

        var vm = new VirtualMachine
        {
            Name = v.Str("name") ?? id ?? "",
            VmId = id,
            Uuid = id,
            BiosUuid = v.Str("biosGuid") is { } bg ? NormalizeGuid(bg) : null,
            Kind = GuestKind.VirtualMachine,
            Platform = Platform.HyperV,
            Host = host.Name,
            Cluster = host.Cluster,
            Datacenter = ctx.Request.Group,
            SourceAddress = ctx.Request.Address,
            SourceProduct = ctx.SourceProduct,
            SourceApiVersion = "WinRM",
            PowerState = power,
            RawState = state,
            Status = v.Str("status"),
            CreationDate = v.Date("creationTime"),
            Uptime = uptime,
            PowerOnTime = power == PowerState.PoweredOn && uptime is { } u ? collectedAt - u : null,
            CpuCount = v.Int("processorCount") ?? proc.Int("count") ?? 0,
            CpuReservation = proc.Int("reserve"),
            CpuLimit = proc.Int("maximum"),
            CpuShares = proc.Int("relativeWeight"),
            MemoryMiB = ToMiB(v.Long("memoryStartup")) ?? 0,
            MemoryAssignedMiB = ToMiB(v.Long("memoryAssigned")),
            MemoryDemandMiB = ToMiB(v.Long("memoryDemand")),
            MemoryMinMiB = ToMiB(v.Long("memoryMinimum")),
            MemoryMaxMiB = ToMiB(v.Long("memoryMaximum")),
            DynamicMemory = v.Bool("dynamicMemoryEnabled"),
            Firmware = gen >= 2 ? "efi" : "bios",
            SecureBoot = gen >= 2 ? fw.Bool("secureBoot") : false,
            HardwareVersion = v.Str("version"),
            Generation = gen is { } g ? $"Gen {g}" : null,
            OsGuest = guest.Str("osName"),
            DnsName = guest.Str("fullyQualifiedDomainName"),
            Annotation = v.Str("notes"),
            ConfigPath = v.Str("configurationLocation") ?? v.Str("path"),
            SnapshotDirectory = v.Str("snapshotFileLocation"),
            BootOrder = gen >= 2
                ? JoinOrNull(fw.Strs("bootOrder").Select(BootDeviceText))
                : JoinOrNull(v.Prop("bios").Strs("startupOrder")),
            AutoStart = v.Str("automaticStartAction") is { } asa ? !asa.Equals("Nothing", StringComparison.OrdinalIgnoreCase) : null,
            StartDelaySeconds = v.Int("automaticStartDelay"),
        };

        if (proc is { } p)
        {
            if (p.Bool("exposeVirtualizationExtensions") is { } nested) vm.Extra["Nested virtualization"] = nested;
            if (p.Bool("compatibilityForMigrationEnabled") is { } compat) vm.Extra["Processor compatibility"] = compat;
        }
        if (fw.Str("secureBootTemplate") is { } sbt) vm.Extra["Secure boot template"] = sbt;
        if (v.Str("automaticStopAction") is { } stop) vm.Extra["Automatic stop action"] = stop;
        if (v.Str("checkpointType") is { } ct) vm.Extra["Checkpoint type"] = ct;
        if (v.Str("replicationState") is { } rs && !rs.Equals("Disabled", StringComparison.OrdinalIgnoreCase))
            vm.Extra["Replication"] = v.Str("replicationHealth") is { } rh ? $"{rs} ({rh})" : rs;
        if (guest is { } gu)
        {
            var details = new List<string>();
            if (gu.Str("osName") is { } on) details.Add($"osName='{on}'");
            if (gu.Str("osVersion") is { } ov) details.Add($"osVersion='{ov}'");
            if (gu.Str("osBuildNumber") is { } ob) details.Add($"buildNumber='{ob}'");
            if (details.Count > 0) vm.Extra["Guest Detailed Data"] = string.Join(" ", details);
        }

        // Failover cluster role.
        var isClustered = v.Bool("isClustered");
        if (ctx.ClusterName is not null)
        {
            vm.HaProtected = isClustered ?? false;
            JsonElement grp = default;
            var found = (id is not null && groupsById.TryGetValue(id, out grp)) || groupsByName.TryGetValue(vm.Name, out grp);
            if (found)
            {
                vm.HaProtected = true;
                vm.HaState = grp.Str("state");
                vm.OwnerNode = grp.Str("ownerNode");
                vm.PreferredOwners = grp.Strs("preferredOwners");
                vm.FailoverPriority = PriorityText(grp.Long("priority"));
                vm.Extra["HA Restart Priority"] = PriorityRvTools(grp.Long("priority"));
                if (grp.Long("failoverThreshold") is { } ft) vm.Extra["Max Failures"] = ft;
                if (grp.Long("failoverPeriod") is { } fp) vm.Extra["Max Failure Window"] = fp;
                if (grp.Long("autoFailbackType") is { } af) vm.Extra["Failback"] = af == 1 ? "Allow" : "Prevent";
            }
            vm.Extra["DAS protection"] = vm.HaProtected;
        }

        // Integration services / heartbeat.
        foreach (var ic in v.Arr("integrationServices"))
        {
            vm.IntegrationComponents.Add(new IntegrationComponent
            {
                Name = ic.Str("name") ?? "",
                Enabled = ic.Bool("enabled") ?? false,
                Status = ic.Str("primaryStatusDescription") ?? ic.Str("primaryOperationalStatus"),
            });
        }
        vm.Heartbeat = MapHeartbeat(v.Str("heartbeat"), v.Str("heartbeatStatus"), power);
        var icVersion = guest.Str("integrationServicesVersion") ?? v.Str("integrationServicesVersion");
        if (icVersion is "0.0" or "0.0.0.0") icVersion = null;
        vm.ToolsVersion = icVersion;
        vm.ToolsRunning = power == PowerState.PoweredOn ? vm.Heartbeat is "green" or "yellow" : false;
        vm.ToolsStatus = power != PowerState.PoweredOn
            ? "toolsNotRunning"
            : vm.ToolsRunning == true
                ? (v.Str("integrationServicesState")?.Contains("Update required", StringComparison.OrdinalIgnoreCase) == true ? "toolsOld" : "toolsOk")
                : (icVersion is null && guest is null ? "toolsNotInstalled" : "toolsNotRunning");

        MapDisks(v, vm, host, mounts, smb, ctx);
        MapNics(v, vm);
        MapCdDrives(v, vm);

        foreach (var s in v.Arr("snapshots"))
        {
            vm.Snapshots.Add(new VmSnapshot
            {
                Name = s.Str("name") ?? "",
                Created = s.Date("creationTime"),
                Parent = s.Str("parentSnapshotName"),
                Type = s.Str("snapshotType"),
                Path = s.Str("path"),
            });
        }

        return vm;
    }

    private static void MapDisks(JsonElement v, VirtualMachine vm, HostSystem host,
        List<KeyValuePair<string, string>> mounts, Dictionary<string, Datastore> smb, MapContext ctx)
    {
        var index = 0;
        foreach (var d in v.Arr("disks"))
        {
            var ctype = d.Str("controllerType");
            var cnum = d.Int("controllerNumber");
            var loc = d.Int("controllerLocation");
            var path = d.Str("path");
            var passthrough = d.Long("diskNumber") is not null;
            var format = d.Str("vhdFormat")?.ToLowerInvariant();
            var vhdType = d.Str("vhdType");

            string? datastore = null;
            if (!passthrough && path is not null)
            {
                datastore = MatchDatastore(path, mounts);
                if (datastore is null && SmbShareRoot(path) is { } share)
                {
                    if (!smb.TryGetValue(share, out var sds))
                    {
                        sds = new Datastore
                        {
                            Name = share,
                            Platform = Platform.HyperV,
                            SourceAddress = ctx.Request.Address,
                            Cluster = ctx.ClusterName,
                            Type = "SMB",
                            Shared = true,
                            Content = share,
                        };
                        smb[share] = sds;
                    }
                    if (!sds.Hosts.Contains(host.Name, StringComparer.OrdinalIgnoreCase)) sds.Hosts.Add(host.Name);
                    datastore = share;
                }
            }

            var disk = new VmDisk
            {
                Index = index,
                Label = $"Hard disk {index + 1}",
                Controller = ctype is null ? null : $"{ctype} {cnum}:{loc}",
                ControllerType = ctype,
                ControllerNumber = cnum,
                Unit = loc,
                Path = path,
                Datastore = datastore,
                CapacityBytes = d.Long("size"),
                UsedBytes = d.Long("fileSize"),
                Format = format,
                Thin = vhdType is null ? null : vhdType.Equals("Dynamic", StringComparison.OrdinalIgnoreCase) || vhdType.Equals("Differencing", StringComparison.OrdinalIgnoreCase),
                Shared = d.Bool("supportPersistentReservations") == true || format == "vhdset",
                ParentPath = d.Str("parentPath"),
                Passthrough = passthrough,
                Options = vhdType,
            };
            if (ctype is not null && ctype.Equals("SCSI", StringComparison.OrdinalIgnoreCase) && loc is { } l) disk.Extra["SCSI Unit #"] = l;
            if (passthrough) disk.Extra["Raw LUN ID"] = $"Disk {d.Long("diskNumber")}";
            vm.Disks.Add(disk);
            index++;
        }
    }

    private static void MapNics(JsonElement v, VirtualMachine vm)
    {
        var nics = v.Arr("networkAdapters").ToList();
        var nameCounts = nics.GroupBy(n => n.Str("name") ?? "", StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.Count(), StringComparer.OrdinalIgnoreCase);
        var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var index = 0;
        foreach (var n in nics)
        {
            var name = n.Str("name") ?? "Network Adapter";
            seen[name] = seen.GetValueOrDefault(name) + 1;
            var label = nameCounts.GetValueOrDefault(name) > 1 ? $"{name} #{seen[name]}" : name;
            var ips = n.Strs("ipAddresses");
            var sw = n.Str("switchName");
            var mode = n.Str("vlanMode");
            var options = new List<string> { n.Bool("dynamicMacAddressEnabled") == false ? "Static MAC" : "Dynamic MAC" };
            if (mode is not null && mode.Equals("Trunk", StringComparison.OrdinalIgnoreCase))
                options.Add($"Trunk {n.Str("allowedVlanIds")} native {n.Int("nativeVlanId")}");
            else if (mode is not null && !mode.Equals("Access", StringComparison.OrdinalIgnoreCase) && !mode.Equals("Untagged", StringComparison.OrdinalIgnoreCase))
                options.Add($"VLAN mode {mode}");

            vm.Nics.Add(new VmNic
            {
                Index = index++,
                Label = label,
                AdapterType = n.Bool("isLegacy") == true ? "Legacy Network Adapter" : "Synthetic",
                Network = sw ?? "Not connected",
                Switch = sw,
                Vlan = AccessVlan(mode, n.Int("accessVlanId")),
                Mac = FormatMac(n.Str("macAddress")),
                Connected = n.Bool("connected") ?? sw is not null,
                StartsConnected = sw is not null,
                Ipv4 = ips.Where(ip => !ip.Contains(':')).ToList(),
                Ipv6 = ips.Where(ip => ip.Contains(':')).ToList(),
                Options = string.Join("; ", options),
            });
        }
    }

    private static void MapCdDrives(JsonElement v, VirtualMachine vm)
    {
        foreach (var c in v.Arr("dvdDrives"))
        {
            var path = c.Str("path");
            var media = c.Str("dvdMediaType");
            vm.CdDrives.Add(new VmCdDrive
            {
                DeviceNode = $"{c.Str("controllerType") ?? "DVD"} {c.Int("controllerNumber")}:{c.Int("controllerLocation")}",
                Media = path,
                Connected = path is not null,
                DeviceType = path is null
                    ? "Empty"
                    : media is not null && media.Equals("PassThrough", StringComparison.OrdinalIgnoreCase) ? $"Host device {path}" : $"ISO {path}",
            });
        }
    }

    // ------------------------------------------------------------------ helpers (public where tests need them)

    /// <summary>Hyper-V VM state to the neutral power state.</summary>
    public static PowerState MapPowerState(string? state)
    {
        if (string.IsNullOrEmpty(state)) return PowerState.Unknown;
        var s = state.Replace("Critical", "", StringComparison.OrdinalIgnoreCase).Replace("-", "").Trim();
        return s.ToLowerInvariant() switch
        {
            "running" or "starting" or "stopping" or "saving" or "pausing" or "resuming" or "reset" or "fastsaving"
                or "forceshutdown" or "forcereboot" or "componentservicing" => PowerState.PoweredOn,
            "off" => PowerState.PoweredOff,
            "saved" or "paused" or "fastsaved" or "hibernated" => PowerState.Suspended,
            _ => PowerState.Unknown,
        };
    }

    /// <summary>
    /// RVTools heartbeat colour: green = OK, yellow = OK but applications critical, red = lost/error,
    /// gray = no contact / disabled / VM not running.
    /// </summary>
    public static string MapHeartbeat(string? vmHeartbeat, string? integrationStatus, PowerState power)
    {
        if (power != PowerState.PoweredOn) return "gray";
        var h = (vmHeartbeat ?? integrationStatus ?? "").Replace(" ", "").ToLowerInvariant();
        if (h.Length == 0) return "gray";
        if (h == "okapplicationscritical") return "yellow";
        if (h.StartsWith("ok", StringComparison.Ordinal)) return "green";
        if (h is "lostcommunication" or "error" or "degraded" or "nonrecoverableerror" or "stressed" or "predictivefailure") return "red";
        return "gray";
    }

    /// <summary>Cluster group priority (0/1000/2000/3000) to text.</summary>
    public static string? PriorityText(long? priority) => priority switch
    {
        null => null,
        3000 => "High",
        2000 => "Medium",
        1000 => "Low",
        0 => "No auto start",
        _ => priority.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    private static string? PriorityRvTools(long? priority) => priority switch
    {
        3000 => "high",
        2000 => "medium",
        1000 => "low",
        0 => "disabled",
        _ => null,
    };

    /// <summary>
    /// The datastore whose mount path is the longest prefix of <paramref name="path"/> (case-insensitive,
    /// on path-segment boundaries), or null. <paramref name="mounts"/> maps mount path to datastore name.
    /// </summary>
    public static string? MatchDatastore(string? path, IEnumerable<KeyValuePair<string, string>> mounts)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        var p = WithTrailingSlash(NormalizePath(path));
        string? best = null;
        var bestLen = -1;
        foreach (var (mount, name) in mounts)
        {
            if (string.IsNullOrWhiteSpace(mount)) continue;
            var pre = WithTrailingSlash(NormalizePath(mount));
            if (pre.Length > bestLen && p.StartsWith(pre, StringComparison.OrdinalIgnoreCase))
            {
                best = name;
                bestLen = pre.Length;
            }
        }
        return best;
    }

    /// <summary>"\\server\share" for a UNC path, else null.</summary>
    public static string? SmbShareRoot(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        var p = NormalizePath(path);
        if (!p.StartsWith(@"\\", StringComparison.Ordinal)) return null;
        var parts = p[2..].Split('\\', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 2 ? $@"\\{parts[0]}\{parts[1]}" : null;
    }

    private static string NormalizePath(string path)
    {
        var p = path.Trim().Replace('/', '\\');
        if (p.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase)) p = @"\\" + p[8..];
        else if (p.StartsWith(@"\\?\", StringComparison.Ordinal)) p = p[4..];
        return p;
    }

    private static string WithTrailingSlash(string p) => p.EndsWith('\\') ? p : p + "\\";

    private static string DisplayMount(string path)
    {
        var p = path.TrimEnd('\\');
        return p.Length == 2 && p[1] == ':' ? p.ToUpperInvariant() : p;
    }

    private static string HealthColour(string health) => health.ToLowerInvariant() switch
    {
        "healthy" => "green",
        "warning" => "yellow",
        "unhealthy" => "red",
        _ => "gray",
    };

    private static int? AccessVlan(string? mode, int? vlan) =>
        vlan is > 0 && (mode is null || mode.Equals("Access", StringComparison.OrdinalIgnoreCase)) ? vlan : null;

    private static string BootDeviceText(string entry) => entry switch
    {
        "HardDiskDrive" => "Hard drive",
        "DvdDrive" => "DVD",
        "VMNetworkAdapter" => "Network",
        _ when entry.StartsWith("File:", StringComparison.Ordinal) => entry.Length > 5 ? $"File ({entry[5..]})" : "File",
        _ => entry,
    };

    private static List<string> ParseNtp(string? source)
    {
        if (string.IsNullOrWhiteSpace(source)) return [];
        var s = source.Trim();
        // Non-NTP sources ("Local CMOS Clock", "VM IC Time Synchronization Provider") are reported verbatim.
        if (s.Contains("CMOS", StringComparison.OrdinalIgnoreCase) || s.Contains("Provider", StringComparison.OrdinalIgnoreCase)
            || s.Contains("Free-running", StringComparison.OrdinalIgnoreCase))
            return [s];
        // "time.contoso.com,0x9 dc01.contoso.com" -> host names without w32time flags.
        return s.Split([' ', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(x => x.Split(',')[0])
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string? JoinOrNull(IEnumerable<string> items)
    {
        var s = string.Join(", ", items.Where(i => !string.IsNullOrWhiteSpace(i)));
        return s.Length == 0 ? null : s;
    }

    private static long? ToMiB(long? bytes) => bytes is null ? null : (long)Math.Round(bytes.Value / 1048576.0);

    private static string NormalizeGuid(string id) => id.Trim().Trim('{', '}').ToLowerInvariant();

    private static string? CollapseSpaces(string? s) =>
        s is null ? null : string.Join(' ', s.Split(' ', StringSplitOptions.RemoveEmptyEntries));

    /// <summary>"00155D012A0B" or "00-15-5D-01-2A-0B" to "00:15:5d:01:2a:0b".</summary>
    public static string? FormatMac(string? mac)
    {
        if (string.IsNullOrWhiteSpace(mac)) return null;
        var hex = new string(mac.Where(Uri.IsHexDigit).ToArray());
        if (hex.Length != 12) return mac;
        return string.Join(':', Enumerable.Range(0, 6).Select(i => hex.Substring(i * 2, 2))).ToLowerInvariant();
    }

    private static string? FormatWwn(string? wwn)
    {
        if (string.IsNullOrWhiteSpace(wwn)) return null;
        var hex = new string(wwn.Where(Uri.IsHexDigit).ToArray());
        if (hex.Length != 16) return wwn;
        return string.Join(':', Enumerable.Range(0, 8).Select(i => hex.Substring(i * 2, 2))).ToLowerInvariant();
    }

    /// <summary>IPv4 prefix length to dotted mask.</summary>
    public static string? PrefixToMask(long? prefix)
    {
        if (prefix is not { } p || p < 0 || p > 32) return null;
        var mask = p == 0 ? 0u : uint.MaxValue << (int)(32 - p);
        return $"{mask >> 24}.{(mask >> 16) & 255}.{(mask >> 8) & 255}.{mask & 255}";
    }
}
