using System.Globalization;
using System.Net;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.Proxmox;

/// <summary>A parsed PVE property string such as <c>local-lvm:vm-100-disk-0,size=32G,discard=on</c>.</summary>
internal sealed class PveProps
{
    private readonly Dictionary<string, string> _values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The leading value without a key (the volume, the first positional "1" of <c>agent: 1,...</c>), if any.</summary>
    public string? Positional { get; private set; }

    /// <summary>Key=value pairs in their original order (first segment included when it has a key).</summary>
    public List<KeyValuePair<string, string>> Pairs { get; } = [];

    public string? this[string key] => _values.GetValueOrDefault(key);

    public bool ContainsKey(string key) => _values.ContainsKey(key);

    public static PveProps Parse(string? value)
    {
        var p = new PveProps();
        if (string.IsNullOrWhiteSpace(value)) return p;
        var first = true;
        foreach (var raw in value.Split(','))
        {
            var seg = raw.Trim();
            if (seg.Length == 0) { first = false; continue; }
            var eq = seg.IndexOf('=');
            if (eq < 0)
            {
                if (first) p.Positional = seg;
                else p.Add(seg, "1"); // bare flag, e.g. "ro"
            }
            else
            {
                p.Add(seg[..eq].Trim(), seg[(eq + 1)..].Trim());
            }
            first = false;
        }
        return p;
    }

    private void Add(string key, string value)
    {
        _values.TryAdd(key, value);
        Pairs.Add(new(key, value));
    }
}

/// <summary>Parsing helpers for PVE config values.</summary>
internal static class PveParsing
{
    public const long MiB = 1024 * 1024;

    public static bool IsTrue(string? s) =>
        s is not null && (s == "1" || s.Equals("on", StringComparison.OrdinalIgnoreCase)
            || s.Equals("yes", StringComparison.OrdinalIgnoreCase) || s.Equals("true", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Parses a PVE size (<c>32G</c>, <c>500M</c>, <c>1T</c>, <c>4M</c>, <c>512K</c>, or bare bytes <c>10737418240</c>).
    /// Units are binary (1K = 1024). Returns null when unparseable.
    /// </summary>
    public static long? ParseSize(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Trim();
        var unit = char.ToUpperInvariant(s[^1]);
        if (unit == 'B' && s.Length > 1 && char.IsLetter(s[^2])) // tolerate "32GB"/"32GiB"-ish suffixes
        {
            s = s.TrimEnd('B', 'b', 'i');
            unit = char.ToUpperInvariant(s[^1]);
        }
        long mult = unit switch
        {
            'K' => 1024L,
            'M' => MiB,
            'G' => 1024L * MiB,
            'T' => 1024L * 1024 * MiB,
            'P' => 1024L * 1024 * 1024 * MiB,
            _ => 1,
        };
        var num = mult == 1 ? s : s[..^1];
        if (!double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) || d < 0) return null;
        return (long)Math.Round(d * mult);
    }

    /// <summary>Splits a volume id <c>storage:volname</c>. Paths (<c>/dev/sdb</c>) and keywords (<c>none</c>, <c>cdrom</c>) have no storage.</summary>
    public static (string? Storage, string? Volume) SplitVolume(string? volid)
    {
        if (string.IsNullOrWhiteSpace(volid) || volid.StartsWith('/')) return (null, volid);
        var colon = volid.IndexOf(':');
        return colon <= 0 ? (null, volid) : (volid[..colon], volid[(colon + 1)..]);
    }

    /// <summary>Disk format from the volume name suffix, or null when not encoded in the name.</summary>
    public static string? FormatFromVolume(string? volume)
    {
        if (volume is null) return null;
        var lower = volume.ToLowerInvariant();
        if (lower.EndsWith(".qcow2")) return "qcow2";
        if (lower.EndsWith(".vmdk")) return "vmdk";
        if (lower.EndsWith(".raw") || lower.EndsWith(".img")) return "raw";
        if (lower.EndsWith(".iso")) return "iso";
        if (lower.Contains("subvol-")) return "subvol";
        return null;
    }

    private static readonly HashSet<string> BlockStorageTypes =
        new(["lvm", "lvmthin", "zfspool", "rbd", "iscsi", "iscsidirect", "zfs", "drbd"], StringComparer.OrdinalIgnoreCase);

    public static bool IsBlockStorage(string? type) => type is not null && BlockStorageTypes.Contains(type);

    /// <summary>
    /// Whether a disk is thin-provisioned, inferred from storage type and format. Null when it can't be known
    /// (e.g. raw files on a directory/NFS storage are sparse but not reported as such).
    /// </summary>
    public static bool? IsThin(string? storageType, string? format)
    {
        if (string.Equals(format, "qcow2", StringComparison.OrdinalIgnoreCase)) return true;
        return storageType?.ToLowerInvariant() switch
        {
            "lvmthin" or "zfspool" or "zfs" or "rbd" or "cephfs" => true,
            "lvm" or "iscsi" or "iscsidirect" => false,
            _ => null,
        };
    }

    private static readonly HashSet<string> SharedStorageTypes =
        new(["nfs", "cifs", "rbd", "cephfs", "glusterfs", "iscsi", "iscsidirect", "pbs", "zfs", "esxi"], StringComparer.OrdinalIgnoreCase);

    public static bool IsInherentlyShared(string? type) => type is not null && SharedStorageTypes.Contains(type);

    /// <summary>Friendly name for a QEMU <c>ostype</c>.</summary>
    public static string? QemuOsName(string? ostype) => ostype?.ToLowerInvariant() switch
    {
        null or "" => null,
        "l26" => "Linux 2.6+ kernel",
        "l24" => "Linux 2.4 kernel",
        "win11" => "Microsoft Windows 11/2022/2025",
        "win10" => "Microsoft Windows 10/2016/2019",
        "win8" => "Microsoft Windows 8.x/2012/2012 R2",
        "win7" => "Microsoft Windows 7/2008 R2",
        "wvista" => "Microsoft Windows Vista",
        "w2k8" => "Microsoft Windows Vista/2008",
        "w2k3" => "Microsoft Windows 2003",
        "w2k" => "Microsoft Windows 2000",
        "wxp" => "Microsoft Windows XP",
        "solaris" => "Solaris kernel",
        "other" => "Other",
        _ => ostype,
    };

    /// <summary>Friendly NIC model name.</summary>
    public static string NicModelName(string model) => model.ToLowerInvariant() switch
    {
        "virtio" => "VirtIO (paravirtualized)",
        "e1000" => "Intel E1000",
        "e1000e" => "Intel E1000E",
        "vmxnet3" => "VMware vmxnet3",
        "rtl8139" => "Realtek RTL8139",
        "ne2k_pci" => "NE2000 PCI",
        "ne2k_isa" => "NE2000 ISA",
        "pcnet" => "AMD PCnet",
        "veth" => "veth",
        _ => model,
    };

    public static readonly HashSet<string> NicModels = new(
        ["virtio", "e1000", "e1000e", "e1000-82540em", "e1000-82544gc", "e1000-82545em", "vmxnet3", "rtl8139",
         "ne2k_pci", "ne2k_isa", "pcnet", "i82551", "i82557b", "i82559er"],
        StringComparer.OrdinalIgnoreCase);

    public static List<string> SplitTags(string? tags) =>
        string.IsNullOrWhiteSpace(tags)
            ? []
            : tags.Split([';', ',', ' '], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

    /// <summary>Maps PVE status (+ QMP status) to a power state.</summary>
    public static PowerState MapPowerState(string? status, string? qmpStatus = null, bool hibernated = false)
    {
        switch (status?.ToLowerInvariant())
        {
            case "running":
                return qmpStatus?.ToLowerInvariant() is "paused" or "suspended" or "prelaunch"
                    ? PowerState.Suspended
                    : PowerState.PoweredOn;
            case "stopped":
                return hibernated ? PowerState.Suspended : PowerState.PoweredOff;
            case "paused" or "suspended":
                return PowerState.Suspended;
            default:
                return PowerState.Unknown;
        }
    }

    /// <summary>Converts "24" or "255.255.255.0" to a dotted IPv4 mask.</summary>
    public static string? ToSubnetMask(string? netmask)
    {
        if (string.IsNullOrWhiteSpace(netmask)) return null;
        if (netmask.Contains('.')) return netmask;
        if (!int.TryParse(netmask, out var prefix) || prefix is < 0 or > 32) return netmask;
        var mask = prefix == 0 ? 0u : uint.MaxValue << (32 - prefix);
        return new IPAddress([(byte)(mask >> 24), (byte)(mask >> 16), (byte)(mask >> 8), (byte)mask]).ToString();
    }

    /// <summary>Strips a "/prefix" from an address ("10.0.0.5/24" → "10.0.0.5").</summary>
    public static string StripPrefix(string address)
    {
        var slash = address.IndexOf('/');
        return slash < 0 ? address : address[..slash];
    }

    public static bool IsIpAddress(string? s) => s is not null && IPAddress.TryParse(StripPrefix(s), out _);

    public static string? NormalizeMac(string? mac) =>
        string.IsNullOrWhiteSpace(mac) ? null : mac.Trim().Replace('-', ':').ToUpperInvariant();

    /// <summary>"pve-manager/8.2.4/faa83925c9641325" → "8.2.4".</summary>
    public static string? PveManagerVersion(string? pveversion)
    {
        if (string.IsNullOrWhiteSpace(pveversion)) return null;
        var parts = pveversion.Split('/');
        return parts.Length >= 2 ? parts[1] : pveversion;
    }

    public static DateTimeOffset? FromUnix(long? seconds) =>
        seconds is > 0 ? DateTimeOffset.FromUnixTimeSeconds(seconds.Value) : null;

    /// <summary>Joins key=value options for display, skipping the given keys.</summary>
    public static string? JoinOptions(PveProps props, params string[] skip)
    {
        var parts = props.Pairs
            .Where(p => !skip.Contains(p.Key, StringComparer.OrdinalIgnoreCase))
            .Select(p => $"{p.Key}={p.Value}")
            .ToList();
        return parts.Count == 0 ? null : string.Join(", ", parts);
    }
}
