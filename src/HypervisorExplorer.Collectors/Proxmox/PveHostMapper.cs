using System.Text.Json;
using System.Text.RegularExpressions;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>Maps PVE node endpoints (status, dns, time, network, hardware/pci) and storage config onto the model.</summary>
internal static partial class PveHostMapper
{
    [GeneratedRegex(@"^(?<raw>.+)\.(?<vid>\d+)$")]
    private static partial Regex VlanName();

    /// <summary>Applies /nodes/{n}/status.</summary>
    public static void ApplyStatus(HostSystem host, JsonElement st, DateTimeOffset now)
    {
        if (st.Prop("cpuinfo") is { } ci)
        {
            host.CpuModel = ci.Str("model")?.Trim();
            host.CpuMhz = ci.Dbl("mhz") is { } mhz ? (int)Math.Round(mhz) : null;
            host.CpuSockets = ci.Int("sockets");
            // PVE's "cores" is the total physical core count across sockets; "cpus" is logical CPUs.
            host.CpuCores = ci.Int("cores");
            host.CpuThreads = ci.Int("cpus");
            if (host.CpuCores is > 0 && host.CpuSockets is > 0)
                host.CoresPerSocket = host.CpuCores / host.CpuSockets;
            if (host.CpuCores is { } c && host.CpuThreads is { } t)
                host.HyperThreadingActive = t > c;
        }
        if (st.Dbl("cpu") is { } cpu) host.CpuUsagePercent = Math.Round(cpu * 100, 1);
        if (st.Prop("memory") is { } mem)
        {
            host.MemoryBytes = mem.Long("total") ?? host.MemoryBytes;
            host.MemoryUsedBytes = mem.Long("used") ?? host.MemoryUsedBytes;
        }
        if (st.Long("uptime") is > 0 and var up)
        {
            var boot = now - TimeSpan.FromSeconds(up);
            host.BootTime = new DateTimeOffset(boot.Ticks - boot.Ticks % TimeSpan.TicksPerSecond, boot.Offset);
        }

        var pve = st.Str("pveversion");
        if (PveParsing.PveManagerVersion(pve) is { } v) host.Version = $"Proxmox VE {v}";
        host.KernelVersion = st.Prop("current-kernel")?.Str("release") ?? KernelRelease(st.Str("kversion"));
        if (st.Prop("boot-info") is { } bi && bi.Str("mode") is { } mode)
            host.Extra["Boot mode"] = bi.Bool("secureboot") == true ? $"{mode} (secure boot)" : mode;
    }

    /// <summary>"Linux 6.8.8-2-pve #1 SMP ..." → "6.8.8-2-pve".</summary>
    private static string? KernelRelease(string? kversion)
    {
        if (string.IsNullOrWhiteSpace(kversion)) return null;
        var parts = kversion.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 2 && parts[0] == "Linux" ? parts[1] : kversion;
    }

    /// <summary>Applies /nodes/{n}/dns.</summary>
    public static void ApplyDns(HostSystem host, JsonElement dns)
    {
        foreach (var key in new[] { "dns1", "dns2", "dns3" })
            if (dns.Str(key) is { Length: > 0 } s) host.DnsServers.Add(s);
        if (dns.Str("search") is { Length: > 0 } search)
        {
            host.DnsSearch = search.Split([' ', ',', ';'], StringSplitOptions.RemoveEmptyEntries).ToList();
            host.Domain = host.DnsSearch.FirstOrDefault();
        }
    }

    /// <summary>Applies /nodes/{n}/time.</summary>
    public static void ApplyTime(HostSystem host, JsonElement time)
    {
        host.TimeZone = time.Str("timezone");
        if (time.Long("localtime") is { } local && time.Long("time") is { } utc)
        {
            var offset = TimeSpan.FromSeconds(local - utc);
            host.Extra["GMT Offset"] = (offset < TimeSpan.Zero ? "-" : "+") + offset.ToString(@"hh\:mm");
        }
    }

    /// <summary>
    /// Applies /nodes/{n}/network: physical NICs (eth) → <see cref="HostNic"/>, Linux/OVS bridges →
    /// <see cref="VirtualSwitch"/>, VLAN interfaces → <see cref="PortGroup"/>, addressed interfaces → <see cref="HostIpInterface"/>.
    /// Bonds are folded into their member NICs and the bridge they uplink.
    /// </summary>
    public static void ApplyNetwork(HostSystem host, JsonElement list)
    {
        var ifaces = list.Items().Where(i => i.Str("iface") is not null)
            .OrderBy(i => i.Str("iface"), StringComparer.Ordinal).ToList();

        static List<string> Split(string? s) =>
            s is null ? [] : s.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).ToList();

        var bridgePorts = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var bondSlaves = new Dictionary<string, (List<string> Slaves, string? Mode)>(StringComparer.Ordinal);
        foreach (var i in ifaces)
        {
            var name = i.Str("iface")!;
            switch (i.Str("type"))
            {
                case "bridge":
                    bridgePorts[name] = Split(i.Str("bridge_ports"));
                    break;
                case "OVSBridge":
                    bridgePorts[name] = Split(i.Str("ovs_ports"));
                    break;
                case "bond":
                    bondSlaves[name] = (Split(i.Str("slaves") ?? i.Str("bond-slaves")), i.Str("bond_mode") ?? i.Str("bond-mode"));
                    break;
                case "OVSBond":
                    bondSlaves[name] = (Split(i.Str("ovs_bonds")), i.Str("ovs_options"));
                    break;
            }
        }

        string? BridgeOf(string port) =>
            bridgePorts.FirstOrDefault(b => b.Value.Contains(port, StringComparer.Ordinal)).Key;

        string? BondOf(string nic) =>
            bondSlaves.FirstOrDefault(b => b.Value.Slaves.Contains(nic, StringComparer.Ordinal)).Key;

        foreach (var i in ifaces)
        {
            var name = i.Str("iface")!;
            var type = i.Str("type");
            var mtu = i.Int("mtu");

            switch (type)
            {
                case "eth":
                {
                    var bond = BondOf(name);
                    var nic = new HostNic
                    {
                        Name = name,
                        Status = i.Bool("active") == true ? "up" : i.Bool("exists") == false ? "missing" : "down",
                        Switch = BridgeOf(name) ?? (bond is null ? null : BridgeOf(bond)),
                        Description = bond is null
                            ? NullIfEmpty(i.Str("comments"))
                            : $"member of {bond}{(bondSlaves[bond].Mode is { } mode ? $" ({mode})" : "")}",
                    };
                    if (bond is not null) nic.Extra["Uplink port"] = bond;
                    if (VlanName().IsMatch(name)) break; // "eno1.20" handled as a VLAN below
                    host.Nics.Add(nic);
                    break;
                }
                case "bridge" or "OVSBridge":
                {
                    var ports = bridgePorts[name];
                    var notes = new List<string>();
                    foreach (var p in ports.Where(bondSlaves.ContainsKey))
                    {
                        var (slaves, mode) = bondSlaves[p];
                        notes.Add($"{p}{(mode is null ? "" : $" ({mode})")}: {string.Join(", ", slaves)}");
                    }
                    var vlanAware = i.Bool("bridge_vlan_aware") == true;
                    if (vlanAware) notes.Add($"VLAN-aware{(i.Str("bridge_vids") is { } vids ? $" ({vids})" : "")}");
                    if (NullIfEmpty(i.Str("comments")) is { } c) notes.Add(c.Trim());
                    host.Switches.Add(new VirtualSwitch
                    {
                        Name = name,
                        Type = type == "OVSBridge" ? "OVS bridge" : vlanAware ? "Linux bridge (VLAN-aware)" : "Linux bridge",
                        Uplinks = ports,
                        Mtu = mtu ?? 1500,
                        Notes = notes.Count == 0 ? null : string.Join("; ", notes),
                    });
                    break;
                }
            }

            // VLAN interfaces: type "vlan" (vmbr0.20 / eno1.20 / named with vlan-raw-device) or OVSIntPort with a tag.
            if (type == "vlan" || (type is "eth" && VlanName().IsMatch(name)))
            {
                var m = VlanName().Match(name);
                var raw = i.Str("vlan-raw-device") ?? (m.Success ? m.Groups["raw"].Value : null);
                var vid = i.Int("vlan-id") ?? (m.Success ? int.Parse(m.Groups["vid"].Value) : null);
                host.PortGroups.Add(new PortGroup
                {
                    Name = name,
                    Switch = raw is null ? null : bridgePorts.ContainsKey(raw) ? raw : BridgeOf(raw) ?? raw,
                    Vlan = vid,
                });
            }
            else if (type == "OVSIntPort")
            {
                host.PortGroups.Add(new PortGroup { Name = name, Switch = i.Str("ovs_bridge"), Vlan = i.Int("ovs_tag") });
            }

            // Addressed interfaces.
            var cidr = i.Str("cidr");
            var address = i.Str("address") ?? (cidr is null ? null : PveParsing.StripPrefix(cidr));
            var cidr6 = i.Str("cidr6");
            var address6 = i.Str("address6") ?? (cidr6 is null ? null : PveParsing.StripPrefix(cidr6));
            var dhcp = i.Str("method") == "dhcp";
            if (address is not null || address6 is not null || dhcp)
            {
                var netmask = i.Str("netmask") ?? (cidr is not null && cidr.Contains('/') ? cidr[(cidr.IndexOf('/') + 1)..] : null);
                var ip = new HostIpInterface
                {
                    Name = name,
                    PortGroup = name,
                    Dhcp = dhcp,
                    Ipv4 = address,
                    SubnetMask = PveParsing.ToSubnetMask(netmask),
                    Gateway = i.Str("gateway"),
                    Ipv6 = address6 is null ? null : i.Str("netmask6") is { } p6 ? $"{address6}/{p6}" : cidr6 ?? address6,
                    Mtu = mtu ?? 1500,
                };
                if (i.Str("gateway6") is { } gw6) ip.Extra["IP 6 Gateway"] = gw6;
                host.IpInterfaces.Add(ip);
            }
        }
    }

    /// <summary>Applies /nodes/{n}/hardware/pci: mass-storage controllers (class 0x01xxxx) become storage adapters.</summary>
    public static void ApplyPci(HostSystem host, JsonElement list)
    {
        foreach (var d in list.Items())
        {
            var cls = d.Str("class");
            if (cls is null) continue;
            var hex = cls.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? cls[2..] : cls;
            if (hex.Length < 4 || !hex.StartsWith("01", StringComparison.Ordinal)) continue;
            var type = hex.Substring(2, 2) switch
            {
                "00" => "SCSI",
                "01" => "IDE",
                "04" => "RAID",
                "05" => "ATA",
                "06" => "SATA (AHCI)",
                "07" => "SAS",
                "08" => "NVMe",
                _ => "Storage",
            };
            var id = d.Str("id") ?? "";
            host.StorageAdapters.Add(new StorageAdapter
            {
                Device = id,
                Pci = id,
                Type = type,
                Model = string.Join(" ", new[] { d.Str("vendor_name"), d.Str("device_name") }.Where(s => !string.IsNullOrWhiteSpace(s))),
            });
        }
    }

    /// <summary>Human-readable address for a /storage config entry (NFS export, Ceph pool, LVM VG, path...).</summary>
    public static string? StorageAddress(JsonElement cfg)
    {
        string? S(string k) => cfg.Str(k);
        return S("type") switch
        {
            "nfs" => $"{S("server")}:{S("export")}",
            "cifs" => $"//{S("server")}/{S("share")}",
            "glusterfs" => $"{S("server")}:{S("volume")}",
            "iscsi" => $"{S("portal")} {S("target")}",
            "iscsidirect" => $"{S("portal")} {S("target")}",
            "rbd" => S("monhost") is { } mon ? $"rbd:{S("pool") ?? "rbd"} @ {mon}" : $"rbd:{S("pool") ?? "rbd"}",
            "cephfs" => S("monhost") is { } mon ? $"cephfs:{S("fs-name") ?? "default"} @ {mon}" : $"cephfs:{S("fs-name") ?? "default"}",
            "pbs" => $"{S("server")}:{S("datastore")}",
            "zfspool" => S("pool"),
            "zfs" => $"{S("portal")}:{S("pool")}",
            "lvm" => S("vgname"),
            "lvmthin" => $"{S("vgname")}/{S("thinpool")}",
            "esxi" => S("server"),
            _ => S("path") ?? S("server"),
        };
    }

    private static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
}
