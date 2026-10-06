using System.Text.Json;
using System.Text.RegularExpressions;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>Maps QEMU VM / LXC container config, status, snapshot and guest-agent JSON onto <see cref="VirtualMachine"/>.</summary>
internal static partial class PveGuestMapper
{
    [GeneratedRegex(@"^(ide|sata|scsi|virtio)(\d+)$", RegexOptions.IgnoreCase)]
    private static partial Regex QemuDiskKey();

    [GeneratedRegex(@"^(net|usb|hostpci|mp|unused)(\d+)$", RegexOptions.IgnoreCase)]
    private static partial Regex IndexedKey();

    private static readonly string[] DiskSkipOptions = ["size", "format", "media", "file", "volume", "mp"];

    private static readonly HashSet<string> PseudoFileSystems = new(
        ["tmpfs", "devtmpfs", "proc", "sysfs", "cgroup", "cgroup2", "overlay", "squashfs", "devpts", "securityfs",
         "pstore", "efivarfs", "bpf", "tracefs", "debugfs", "mqueue", "hugetlbfs", "configfs", "fusectl", "autofs",
         "nsfs", "ramfs", "binfmt_misc", "rpc_pipefs", "fuse.lxcfs", "fuse.gvfsd-fuse", "fuse.portal", "iso9660",
         "udf", "nfsd", "selinuxfs", "none"],
        StringComparer.OrdinalIgnoreCase);

    // ---------------------------------------------------------------- QEMU

    /// <summary>Applies /nodes/{n}/qemu/{id}/config.</summary>
    public static void ApplyQemuConfig(VirtualMachine vm, JsonElement cfg, Func<string, string?> storageType)
    {
        if (cfg.Str("name") is { Length: > 0 } name) vm.Name = name;

        var cores = cfg.Int("cores") ?? 1;
        var sockets = cfg.Int("sockets") ?? 1;
        vm.Sockets = sockets;
        vm.CoresPerSocket = cores;
        vm.CpuCount = cfg.Int("vcpus") is > 0 and var v ? v : cores * sockets;
        vm.CpuShares = cfg.Int("cpuunits");
        var cpu = PveProps.Parse(cfg.Str("cpu"));
        vm.CpuType = cpu.Positional ?? cpu["cputype"];
        var hotplug = cfg.Str("hotplug");
        vm.CpuHotAdd = hotplug is null ? false : hotplug.Split(',').Contains("cpu", StringComparer.OrdinalIgnoreCase);

        var memProps = PveProps.Parse(cfg.Str("memory"));
        // "memory: 4096" or (PVE 8.1+) "memory: current=4096"; MiB, default 512.
        var memory = long.TryParse(memProps.Positional ?? memProps["current"], out var memVal) ? memVal : 512;
        vm.MemoryMiB = memory;
        vm.MemoryMaxMiB = memory;
        var balloon = cfg.Long("balloon");
        if (balloon is > 0)
        {
            vm.MemoryMinMiB = balloon;
            vm.DynamicMemory = balloon < memory;
        }
        else
        {
            vm.DynamicMemory = false;
        }

        var bios = cfg.Str("bios");
        var efi = string.Equals(bios, "ovmf", StringComparison.OrdinalIgnoreCase);
        vm.Firmware = efi ? "efi" : "bios";
        if (efi)
        {
            var efidisk = PveProps.Parse(cfg.Str("efidisk0"));
            vm.SecureBoot = PveParsing.IsTrue(efidisk["pre-enrolled-keys"]);
        }
        else
        {
            vm.SecureBoot = false;
        }

        vm.HardwareVersion = PveProps.Parse(cfg.Str("machine")) is var mach && (mach.Positional ?? mach["type"]) is { } mt
            ? mt
            : "pc (i440fx)";
        vm.OsConfigured = PveParsing.QemuOsName(cfg.Str("ostype")) ?? vm.OsConfigured;
        vm.Annotation = NullIfEmpty(cfg.Str("description")?.TrimEnd());

        var smbios = PveProps.Parse(cfg.Str("smbios1"));
        vm.BiosUuid = smbios["uuid"];
        vm.Uuid = cfg.Str("vmgenid") is { Length: > 0 } gen && gen != "0" ? gen : vm.BiosUuid;

        ApplyStartup(vm, cfg);
        if (cfg.Str("tags") is { } tags) vm.Tags = PveParsing.SplitTags(tags);

        var boot = cfg.Str("boot");
        if (!string.IsNullOrWhiteSpace(boot))
        {
            var bp = PveProps.Parse(boot);
            vm.BootOrder = bp["order"] is { } order ? string.Join(", ", order.Split(';', StringSplitOptions.RemoveEmptyEntries)) : boot;
        }

        var meta = PveProps.Parse(cfg.Str("meta"));
        if (long.TryParse(meta["ctime"], out var ctime)) vm.CreationDate = PveParsing.FromUnix(ctime);

        if (cfg.Str("lock") is { Length: > 0 } lck)
        {
            vm.Status = "yellow";
            vm.Extra["Lock"] = lck;
        }

        // Disks and CD drives, ordered by bus then index.
        var scsihw = cfg.Str("scsihw") ?? "lsi";
        var diskEntries = cfg.StringProps()
            .Select(p => (p.Key, p.Value, Match: QemuDiskKey().Match(p.Key)))
            .Where(p => p.Match.Success || p.Key is "efidisk0" or "tpmstate0")
            .OrderBy(p => p.Match.Success ? BusRank(p.Match.Groups[1].Value) : 10)
            .ThenBy(p => p.Match.Success ? int.Parse(p.Match.Groups[2].Value) : 0)
            .ToList();
        foreach (var (key, value, match) in diskEntries)
        {
            var props = PveProps.Parse(value);
            if (string.Equals(props["media"], "cdrom", StringComparison.OrdinalIgnoreCase))
            {
                vm.CdDrives.Add(CdDrive(key, props));
                continue;
            }

            var bus = match.Success ? match.Groups[1].Value.ToLowerInvariant() : key;
            var disk = BuildDisk(key, props, storageType);
            disk.Index = vm.Disks.Count;
            disk.Unit = match.Success ? int.Parse(match.Groups[2].Value) : 0;
            disk.ControllerType = bus switch
            {
                "scsi" => $"SCSI ({scsihw})",
                "virtio" => "VirtIO Block",
                "ide" => "IDE",
                "sata" => "SATA",
                "efidisk0" => "EFI vars disk",
                "tpmstate0" => "TPM state",
                _ => bus,
            };
            if (bus == "ide") disk.ControllerNumber = disk.Unit / 2;
            vm.Disks.Add(disk);
        }

        // NICs, USB and PCI passthrough.
        var passthrough = new List<string>();
        foreach (var (key, value) in cfg.StringProps().OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            var m = IndexedKey().Match(key);
            if (!m.Success) continue;
            var idx = int.Parse(m.Groups[2].Value);
            switch (m.Groups[1].Value.ToLowerInvariant())
            {
                case "net":
                    vm.Nics.Add(QemuNic(key, idx, value));
                    break;
                case "usb":
                    var usb = PveProps.Parse(value);
                    vm.UsbDevices.Add(new VmUsbDevice
                    {
                        DeviceNode = key,
                        DeviceType = usb["host"] is { } host ? $"host={host}" : usb["mapping"] is { } map ? $"mapping={map}" : usb.Positional ?? value,
                        Extra = { ["Auto connect"] = "True" },
                    });
                    break;
                case "hostpci":
                    passthrough.Add($"{key}: {value}");
                    break;
            }
        }
        vm.Nics.Sort((a, b) => a.Index.CompareTo(b.Index));
        if (passthrough.Count > 0) vm.Extra["Fixed Passthru HotPlug"] = string.Join("; ", passthrough);
    }

    private static int BusRank(string bus) => bus.ToLowerInvariant() switch
    {
        "scsi" => 0,
        "virtio" => 1,
        "sata" => 2,
        "ide" => 3,
        _ => 9,
    };

    private static VmCdDrive CdDrive(string key, PveProps props)
    {
        var media = props.Positional ?? props["file"];
        var empty = media is null || media.Equals("none", StringComparison.OrdinalIgnoreCase);
        var host = media is not null && media.Equals("cdrom", StringComparison.OrdinalIgnoreCase);
        return new VmCdDrive
        {
            DeviceNode = key,
            Media = empty ? null : media,
            Connected = !empty,
            DeviceType = empty ? "Empty drive" : host ? "Host CD/DVD drive" : $"ISO {media}",
        };
    }

    private static VmDisk BuildDisk(string key, PveProps props, Func<string, string?> storageType)
    {
        var volid = props.Positional ?? props["file"] ?? props["volume"];
        var (storage, volume) = PveParsing.SplitVolume(volid);
        var type = storage is null ? null : storageType(storage);
        var format = props["format"] ?? PveParsing.FormatFromVolume(volume)
            ?? (PveParsing.IsBlockStorage(type) ? "raw" : null);
        var passthrough = volid is not null && volid.StartsWith('/');
        return new VmDisk
        {
            Label = key,
            Controller = key,
            Path = volid,
            Datastore = storage,
            CapacityBytes = PveParsing.ParseSize(props["size"]),
            Format = format,
            Thin = passthrough ? false : PveParsing.IsThin(type, format),
            Shared = props.ContainsKey("shared") ? PveParsing.IsTrue(props["shared"]) : null,
            Cache = props["cache"] ?? (key.StartsWith("mp") || key == "rootfs" ? null : "none (default)"),
            Passthrough = passthrough,
            Options = PveParsing.JoinOptions(props, [.. DiskSkipOptions, "cache"]),
            Extra = { ["Disk Key"] = key },
        };
    }

    private static VmNic QemuNic(string key, int idx, string value)
    {
        var p = PveProps.Parse(value);
        string? model = p["model"];
        string? mac = p["macaddr"];
        foreach (var kv in p.Pairs)
        {
            if (PveParsing.NicModels.Contains(kv.Key))
            {
                model ??= kv.Key;
                mac ??= kv.Value;
                break;
            }
        }
        model ??= p.Positional ?? "nic";
        var linkDown = PveParsing.IsTrue(p["link_down"]);
        return new VmNic
        {
            Index = idx,
            Label = key,
            AdapterType = PveParsing.NicModelName(model),
            Network = p["bridge"],
            Switch = p["bridge"],
            Vlan = int.TryParse(p["tag"], out var tag) ? tag : null,
            Mac = PveParsing.NormalizeMac(mac),
            Connected = !linkDown,
            StartsConnected = !linkDown,
            Firewall = p.ContainsKey("firewall") ? PveParsing.IsTrue(p["firewall"]) : false,
            Options = PveParsing.JoinOptions(p, ["bridge", "tag", "firewall", "link_down", "macaddr", "model", model]),
        };
    }

    private static void ApplyStartup(VirtualMachine vm, JsonElement cfg)
    {
        vm.AutoStart = cfg.Bool("onboot") ?? false;
        var startup = PveProps.Parse(cfg.Str("startup"));
        if (int.TryParse(startup["up"], out var up)) vm.StartDelaySeconds = up;
        if (startup["order"] is { } order) vm.Extra["Startup order"] = order;
    }

    /// <summary>Applies /nodes/{n}/qemu/{id}/status/current.</summary>
    public static void ApplyQemuStatus(VirtualMachine vm, JsonElement st, bool hibernated, DateTimeOffset now)
    {
        var status = st.Str("status");
        var qmp = st.Str("qmpstatus");
        vm.PowerState = PveParsing.MapPowerState(status, qmp, hibernated);
        vm.RawState = qmp is { Length: > 0 } && !string.Equals(qmp, status, StringComparison.OrdinalIgnoreCase) ? qmp : status;
        if (hibernated && vm.PowerState == PowerState.Suspended) vm.RawState = "hibernated";
        ApplyUptime(vm, st, now);

        if (vm.PowerState == PowerState.PoweredOn)
        {
            var maxmem = st.Long("maxmem");
            var bi = st.Prop("ballooninfo");
            var actual = bi?.Long("actual") ?? st.Long("balloon");
            vm.MemoryAssignedMiB = ToMiB(actual ?? maxmem);
            if (st.Long("mem") is > 0 and var used) vm.MemoryDemandMiB = ToMiB(used);
            var max = bi?.Long("max_mem") ?? maxmem;
            if (actual is { } a && max is { } mx && mx > a) vm.MemoryBalloonedMiB = ToMiB(mx - a);
            else vm.MemoryBalloonedMiB = 0;
        }

        ApplyHaFromStatus(vm, st);
    }

    // ---------------------------------------------------------------- LXC

    /// <summary>Applies /nodes/{n}/lxc/{id}/config.</summary>
    public static void ApplyLxcConfig(VirtualMachine vm, JsonElement cfg, Func<string, string?> storageType)
    {
        var hostname = cfg.Str("hostname");
        if (string.IsNullOrEmpty(vm.Name) && hostname is not null) vm.Name = hostname;
        vm.DnsName = hostname;

        if (cfg.Int("cores") is { } cores) vm.CpuCount = cores;
        vm.Sockets = 1;
        vm.CoresPerSocket = vm.CpuCount;
        vm.CpuShares = cfg.Int("cpuunits");
        vm.MemoryMiB = cfg.Long("memory") ?? 512;
        vm.MemoryMaxMiB = vm.MemoryMiB;
        vm.DynamicMemory = false;

        var ostype = cfg.Str("ostype");
        vm.OsConfigured = ostype is null ? "LXC" : $"LXC: {ostype}";
        vm.Annotation = NullIfEmpty(cfg.Str("description")?.TrimEnd());
        ApplyStartup(vm, cfg);
        if (cfg.Str("tags") is { } tags) vm.Tags = PveParsing.SplitTags(tags);

        if (cfg.Str("lock") is { Length: > 0 } lck)
        {
            vm.Status = "yellow";
            vm.Extra["Lock"] = lck;
        }

        var details = new List<string>();
        if (cfg.Str("arch") is { } arch) details.Add($"arch='{arch}'");
        details.Add($"unprivileged='{(cfg.Bool("unprivileged") == true ? 1 : 0)}'");
        if (cfg.Str("features") is { Length: > 0 } features) details.Add($"features='{features}'");
        if (cfg.Long("swap") is { } swap) details.Add($"swap='{swap} MiB'");
        vm.Extra["Guest Detailed Data"] = string.Join(" ", details);
        vm.Extra["Unprivileged"] = cfg.Bool("unprivileged") == true;

        var entries = cfg.StringProps()
            .Where(p => p.Key == "rootfs" || (IndexedKey().Match(p.Key) is { Success: true } m && m.Groups[1].Value == "mp"))
            .OrderBy(p => p.Key == "rootfs" ? -1 : int.Parse(p.Key[2..]))
            .ToList();
        foreach (var (key, value) in entries)
        {
            var props = PveProps.Parse(value);
            var disk = BuildDisk(key, props, storageType);
            disk.Index = vm.Disks.Count;
            disk.Unit = key == "rootfs" ? 0 : int.Parse(key[2..]) + 1;
            disk.ControllerType = disk.Passthrough == true ? "Bind mount" : "LXC mount point";
            disk.Options = PveParsing.JoinOptions(props, ["size", "format", "volume"]);
            if (key == "rootfs") disk.Options = disk.Options is null ? "mp=/" : "mp=/, " + disk.Options;
            vm.Disks.Add(disk);
        }

        foreach (var (key, value) in cfg.StringProps())
        {
            var m = IndexedKey().Match(key);
            if (!m.Success || m.Groups[1].Value != "net") continue;
            var p = PveProps.Parse(value);
            var nic = new VmNic
            {
                Index = int.Parse(m.Groups[2].Value),
                Label = p["name"] ?? key,
                AdapterType = PveParsing.NicModelName(p["type"] ?? "veth"),
                Network = p["bridge"],
                Switch = p["bridge"],
                Vlan = int.TryParse(p["tag"], out var tag) ? tag : null,
                Mac = PveParsing.NormalizeMac(p["hwaddr"]),
                Connected = !PveParsing.IsTrue(p["link_down"]),
                StartsConnected = !PveParsing.IsTrue(p["link_down"]),
                Firewall = PveParsing.IsTrue(p["firewall"]),
                Options = PveParsing.JoinOptions(p, ["name", "bridge", "tag", "hwaddr", "firewall", "link_down", "type"]),
            };
            if (p["ip"] is { } ip && PveParsing.IsIpAddress(ip)) nic.Ipv4.Add(PveParsing.StripPrefix(ip));
            if (p["ip6"] is { } ip6 && PveParsing.IsIpAddress(ip6)) nic.Ipv6.Add(PveParsing.StripPrefix(ip6));
            vm.Nics.Add(nic);
        }
        vm.Nics.Sort((a, b) => a.Index.CompareTo(b.Index));
    }

    /// <summary>Applies /nodes/{n}/lxc/{id}/status/current.</summary>
    public static void ApplyLxcStatus(VirtualMachine vm, JsonElement st, DateTimeOffset now)
    {
        var status = st.Str("status");
        vm.PowerState = PveParsing.MapPowerState(status);
        vm.RawState = status;
        ApplyUptime(vm, st, now);
        if (vm.PowerState == PowerState.PoweredOn)
        {
            vm.MemoryAssignedMiB = ToMiB(st.Long("maxmem"));
            if (st.Long("mem") is > 0 and var used) vm.MemoryDemandMiB = ToMiB(used);
            // For containers "disk" is the root filesystem usage.
            if (st.Long("disk") is > 0 and var diskUsed && vm.Disks.FirstOrDefault(d => d.Label == "rootfs") is { } root)
                root.UsedBytes = diskUsed;
        }
        ApplyHaFromStatus(vm, st);
    }

    /// <summary>Applies /nodes/{n}/lxc/{id}/interfaces (PVE 7.2+): runtime addresses of a running container.</summary>
    public static void ApplyLxcInterfaces(VirtualMachine vm, JsonElement interfaces)
    {
        foreach (var iface in interfaces.Items())
        {
            var name = iface.Str("name");
            if (name == "lo") continue;
            var mac = PveParsing.NormalizeMac(iface.Str("hwaddr") ?? iface.Str("hardware-address"));
            var nic = vm.Nics.FirstOrDefault(n => mac is not null && n.Mac == mac)
                      ?? vm.Nics.FirstOrDefault(n => n.Label == name);
            if (nic is null) continue;

            var v4 = new List<string>();
            var v6 = new List<string>();
            if (iface.Str("inet") is { } inet) v4.AddRange(inet.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(PveParsing.StripPrefix));
            if (iface.Str("inet6") is { } inet6) v6.AddRange(inet6.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(PveParsing.StripPrefix));
            foreach (var a in iface.Prop("ip-addresses").Items())
            {
                var ip = a.Str("ip-address");
                if (ip is null) continue;
                if (a.Str("ip-address-type") is "inet6" or "ipv6") v6.Add(ip); else v4.Add(ip);
            }
            MergeIps(nic, v4, v6);
        }
    }

    // ---------------------------------------------------------------- shared

    private static void ApplyUptime(VirtualMachine vm, JsonElement st, DateTimeOffset now)
    {
        if (st.Long("uptime") is > 0 and var up && vm.PowerState != PowerState.PoweredOff)
        {
            vm.Uptime = TimeSpan.FromSeconds(up);
            var on = now - vm.Uptime.Value;
            vm.PowerOnTime = new DateTimeOffset(on.Ticks - on.Ticks % TimeSpan.TicksPerSecond, on.Offset);
        }
    }

    private static void ApplyHaFromStatus(VirtualMachine vm, JsonElement st)
    {
        if (st.Prop("ha") is not { ValueKind: JsonValueKind.Object } ha) return;
        var managed = ha.Bool("managed") == true;
        vm.HaProtected ??= managed;
        if (managed)
        {
            vm.HaState ??= ha.Str("state");
            vm.FailoverPriority ??= ha.Str("group");
        }
    }

    /// <summary>Parses /snapshot, excluding the "current" pseudo-entry.</summary>
    public static List<VmSnapshot> ParseSnapshots(JsonElement list)
    {
        var result = new List<VmSnapshot>();
        foreach (var s in list.Items())
        {
            var name = s.Str("name");
            if (name is null || name == "current") continue;
            var mem = s.Bool("vmstate") == true;
            result.Add(new VmSnapshot
            {
                Name = name,
                Description = NullIfEmpty(s.Str("description")?.TrimEnd()),
                Created = PveParsing.FromUnix(s.Long("snaptime")),
                Parent = s.Str("parent"),
                IncludesMemory = mem,
                Type = mem ? "poweredOn" : "poweredOff",
            });
        }
        return result.OrderBy(s => s.Created ?? DateTimeOffset.MinValue).ToList();
    }

    /// <summary>Unwraps the guest-agent envelope: PVE returns {"result": ...} inside "data".</summary>
    public static JsonElement AgentResult(JsonElement data) =>
        data.Prop("result") is { } r ? r : data;

    /// <summary>Applies agent network-get-interfaces: IPs matched to VM NICs by MAC.</summary>
    public static void ApplyAgentNetwork(VirtualMachine vm, JsonElement result)
    {
        foreach (var iface in AgentResult(result).Items())
        {
            var mac = PveParsing.NormalizeMac(iface.Str("hardware-address"));
            if (iface.Str("name") is "lo" or "Loopback Pseudo-Interface 1" || mac is null || mac == "00:00:00:00:00:00") continue;
            var nic = vm.Nics.FirstOrDefault(n => n.Mac == mac);
            if (nic is null) continue;
            var v4 = new List<string>();
            var v6 = new List<string>();
            foreach (var a in iface.Prop("ip-addresses").Items())
            {
                var ip = a.Str("ip-address");
                if (ip is null || ip.StartsWith("127.") || ip == "::1") continue;
                if (a.Str("ip-address-type") == "ipv6") v6.Add(ip); else v4.Add(ip);
            }
            MergeIps(nic, v4, v6);
        }
    }

    private static void MergeIps(VmNic nic, List<string> v4, List<string> v6)
    {
        foreach (var ip in v4.Where(i => !i.StartsWith("127.")))
            if (!nic.Ipv4.Contains(ip)) nic.Ipv4.Add(ip);
        foreach (var ip in v6.Where(i => i != "::1"))
            if (!nic.Ipv6.Contains(ip, StringComparer.OrdinalIgnoreCase)) nic.Ipv6.Add(ip);
    }

    /// <summary>Applies agent get-osinfo.</summary>
    public static void ApplyAgentOsInfo(VirtualMachine vm, JsonElement result)
    {
        var r = AgentResult(result);
        vm.OsGuest = r.Str("pretty-name") ?? JoinNonEmpty(r.Str("name"), r.Str("version"));
        var details = new List<string>();
        if (r.Str("pretty-name") is { } pn) details.Add($"prettyName='{pn}'");
        if (r.Str("id") is { } id) details.Add($"distroName='{id}'");
        if (r.Str("version-id") is { } vid) details.Add($"distroVersion='{vid}'");
        if (r.Str("kernel-release") is { } kr) details.Add($"kernelVersion='{kr}'");
        if (r.Str("machine") is { } mach) details.Add($"architecture='{mach}'");
        if (details.Count > 0) vm.Extra["Guest Detailed Data"] = string.Join(" ", details);
    }

    /// <summary>Applies agent get-fsinfo: real file systems become partitions.</summary>
    public static void ApplyAgentFsInfo(VirtualMachine vm, JsonElement result)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var fs in AgentResult(result).Items())
        {
            var mount = fs.Str("mountpoint");
            var type = fs.Str("type");
            var total = fs.Long("total-bytes");
            if (mount is null || total is not > 0 || (type is not null && PseudoFileSystems.Contains(type))) continue;
            if (!seen.Add(mount)) continue;
            var used = fs.Long("used-bytes");
            vm.Partitions.Add(new VmPartition
            {
                Name = mount,
                FileSystem = type,
                CapacityBytes = total,
                FreeBytes = used is { } u ? Math.Max(0, total.Value - u) : null,
            });
        }
        vm.Partitions.Sort((a, b) => string.CompareOrdinal(a.Name, b.Name));
    }

    /// <summary>Applies agent get-host-name.</summary>
    public static void ApplyAgentHostName(VirtualMachine vm, JsonElement result)
    {
        if (AgentResult(result).Str("host-name") is { Length: > 0 } h) vm.DnsName = h;
    }

    private static long? ToMiB(long? bytes) => bytes is null ? null : (long)Math.Round(bytes.Value / (double)PveParsing.MiB);

    private static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;

    private static string? JoinNonEmpty(params string?[] parts)
    {
        var p = parts.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        return p.Count == 0 ? null : string.Join(" ", p);
    }
}
