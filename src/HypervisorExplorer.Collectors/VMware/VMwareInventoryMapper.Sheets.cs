using System.Xml.Linq;

namespace HypervisorExplorer.Collectors.VMware;

// Whole RVTools sheets that the Core model has no objects for. Each row is keyed by exact RVTools headers.
public sealed partial class VMwareInventoryMapper
{
    // ---------------------------------------------------------------- vRP

    private List<Dictionary<string, object?>> ResourcePoolRows()
    {
        var vmById = _d.VirtualMachines.ToDictionary(v => v.Ref.Value, StringComparer.Ordinal);
        var rows = new List<Dictionary<string, object?>>();
        foreach (var rp in _d.ResourcePools)
        {
            var s = rp.Get("summary");
            var cpu = s.At("config.cpuAllocation");
            var mem = s.At("config.memoryAllocation");
            var rcpu = s.At("runtime.cpu");
            var rmem = s.At("runtime.memory");
            var q = s.El("quickStats");
            var vms = rp.Refs("vm").Select(r => vmById.GetValueOrDefault(r.Value)).OfType<VimObject>().ToList();

            var row = SheetRow();
            row["Resource Pool name"] = rp.Str("name") ?? s.Str("name");
            row["Resource Pool path"] = InventoryPath(rp.Ref);
            row["Status"] = s.Str("runtime.overallStatus") ?? rp.Str("overallStatus");
            row["# VMs total"] = vms.Count;
            row["# VMs"] = vms.Count(v => v.Str("runtime.powerState") == "poweredOn");
            row["# vCPUs"] = vms.Sum(v => v.Int("config.hardware.numCPU") ?? 0);
            row["CPU limit"] = cpu.Long("limit");
            row["CPU overheadLimit"] = cpu.Long("overheadLimit");
            row["CPU reservation"] = cpu.Long("reservation");
            row["CPU level"] = cpu.Str("shares.level");
            row["CPU shares"] = cpu.Int("shares.shares");
            row["CPU expandableReservation"] = cpu.Bool("expandableReservation");
            row["CPU maxUsage"] = rcpu.Long("maxUsage");
            row["CPU overallUsage"] = rcpu.Long("overallUsage");
            row["CPU reservationUsed"] = rcpu.Long("reservationUsed");
            row["CPU reservationUsedForVm"] = rcpu.Long("reservationUsedForVm");
            row["CPU unreservedForPool"] = rcpu.Long("unreservedForPool");
            row["CPU unreservedForVm"] = rcpu.Long("unreservedForVm");
            row["Mem Configured"] = s.Long("configuredMemoryMB");
            row["Mem limit"] = mem.Long("limit");
            row["Mem overheadLimit"] = mem.Long("overheadLimit");
            row["Mem reservation"] = mem.Long("reservation");
            row["Mem level"] = mem.Str("shares.level");
            row["Mem shares"] = mem.Int("shares.shares");
            row["Mem expandableReservation"] = mem.Bool("expandableReservation");
            // Runtime memory usage is reported in bytes; RVTools shows MiB like the allocation values.
            row["Mem maxUsage"] = Mib(rmem.Long("maxUsage"));
            row["Mem overallUsage"] = Mib(rmem.Long("overallUsage"));
            row["Mem reservationUsed"] = Mib(rmem.Long("reservationUsed"));
            row["Mem reservationUsedForVm"] = Mib(rmem.Long("reservationUsedForVm"));
            row["Mem unreservedForPool"] = Mib(rmem.Long("unreservedForPool"));
            row["Mem unreservedForVm"] = Mib(rmem.Long("unreservedForVm"));
            row["QS overallCpuDemand"] = q.Long("overallCpuDemand");
            row["QS overallCpuUsage"] = q.Long("overallCpuUsage");
            row["QS staticCpuEntitlement"] = q.Long("staticCpuEntitlement");
            row["QS distributedCpuEntitlement"] = q.Long("distributedCpuEntitlement");
            row["QS balloonedMemory"] = q.Long("balloonedMemory");
            row["QS compressedMemory"] = q.Long("compressedMemory");
            row["QS consumedOverheadMemory"] = q.Long("consumedOverheadMemory");
            row["QS distributedMemoryEntitlement"] = q.Long("distributedMemoryEntitlement");
            row["QS guestMemoryUsage"] = q.Long("guestMemoryUsage");
            row["QS hostMemoryUsage"] = q.Long("hostMemoryUsage");
            row["QS overheadMemory"] = q.Long("overheadMemory");
            row["QS privateMemory"] = q.Long("privateMemory");
            row["QS sharedMemory"] = q.Long("sharedMemory");
            row["QS staticMemoryEntitlement"] = q.Long("staticMemoryEntitlement");
            row["QS swappedMemory"] = q.Long("swappedMemory");
            row["Object ID"] = rp.Ref.Value;
            rows.Add(row);
        }
        return rows;
    }

    // ---------------------------------------------------------------- dvSwitch / dvPort

    /// <summary>Value of an inheritable DVS policy (BoolPolicy/StringPolicy/IntPolicy/LongPolicy): &lt;x&gt;&lt;value&gt;.</summary>
    private static XElement? PolicyValue(XElement? e, string name) => e.El(name).El("value");

    private static void ShapingColumns(Dictionary<string, object?> row, XElement? portConfig)
    {
        foreach (var (dir, el) in new[] { ("In", "inShapingPolicy"), ("Out", "outShapingPolicy") })
        {
            var sp = portConfig.El(el);
            row[$"{dir} Traffic Shaping"] = PolicyValue(sp, "enabled").Bool();
            row[$"{dir} Avg"] = PolicyValue(sp, "averageBandwidth").Long();
            row[$"{dir} Peak"] = PolicyValue(sp, "peakBandwidth").Long();
            row[$"{dir} Burst"] = PolicyValue(sp, "burstSize").Long();
        }
    }

    private List<Dictionary<string, object?>> DvSwitchRows()
    {
        var rows = new List<Dictionary<string, object?>>();
        foreach (var sw in _d.DistributedSwitches)
        {
            var cfg = sw.Get("config");
            var sum = sw.Get("summary");
            var product = cfg.El("productInfo") ?? sum.El("productInfo");
            var hostMembers = sum.Els("hostMember").Select(MoRef.From).OfType<MoRef>().ToList();
            if (hostMembers.Count == 0)
                hostMembers = cfg.Els("host").Select(h => MoRef.From(h.At("config.host"))).OfType<MoRef>().ToList();
            var lacp = cfg.Els("lacpGroupConfig").ToList();

            var row = SheetRow();
            row["Switch"] = sw.Str("name") ?? cfg.Str("name");
            row["Datacenter"] = DatacenterOf(sw.Ref);
            row["Name"] = product.Str("name");
            row["Vendor"] = product.Str("vendor");
            row["Version"] = product.Str("version");
            row["Description"] = cfg.Str("description") ?? sum.Str("description");
            row["Created"] = cfg.Date("createTime");
            row["Host members"] = Join(hostMembers.Select(NameOf));
            row["Max Ports"] = cfg.Int("maxPorts");
            row["# Ports"] = cfg.Int("numPorts") ?? sum.Int("numPorts");
            row["# VMs"] = sum.Els("vm").Count();
            ShapingColumns(row, cfg.El("defaultPortConfig"));
            row["CDP Type"] = cfg.Str("linkDiscoveryProtocolConfig.protocol");
            row["CDP Operation"] = cfg.Str("linkDiscoveryProtocolConfig.operation");
            row["LACP Name"] = Join(lacp.Select(l => l.Str("name")));
            row["LACP Mode"] = Join(lacp.Select(l => l.Str("mode")));
            row["LACP Load Balance Alg."] = Join(lacp.Select(l => l.Str("loadbalanceAlgorithm")));
            row["Max MTU"] = cfg.Int("maxMtu");
            row["Contact"] = cfg.Str("contact.contact") ?? sum.Str("contact.contact");
            row["Admin Name"] = cfg.Str("contact.name") ?? sum.Str("contact.name");
            row["Object ID"] = sw.Ref.Value;
            rows.Add(row);
        }
        return rows;
    }

    private static string? VlanText(XElement? vlan) => vlan.XsiType() switch
    {
        "VmwareDistributedVirtualSwitchTrunkVlanSpec" => Join(vlan.Els("vlanId").Select(r =>
            r.Int("start") == r.Int("end") ? r.Str("start") : $"{r.Str("start")}-{r.Str("end")}")),
        "VmwareDistributedVirtualSwitchPvlanSpec" => $"PVLAN {vlan.Str("pvlanId")}",
        _ => vlan.Str("vlanId"),
    };

    private List<Dictionary<string, object?>> DvPortRows()
    {
        var rows = new List<Dictionary<string, object?>>();
        foreach (var pg in _d.DistributedPortgroups)
        {
            var cfg = pg.Get("config");
            var pc = cfg.El("defaultPortConfig");
            var sec = pc.El("securityPolicy");
            var mac = pc.El("macManagementPolicy");
            var team = pc.El("uplinkTeamingPolicy");
            var fc = team.El("failureCriteria");
            var order = team.El("uplinkPortOrder");
            var pol = cfg.El("policy");

            var row = SheetRow();
            row["Port"] = pg.Str("name") ?? cfg.Str("name");
            row["Switch"] = cfg.El("distributedVirtualSwitch") is { } swr ? NameOf(MoRef.From(swr)) : null;
            row["Type"] = cfg.Str("type");
            row["# Ports"] = cfg.Int("numPorts");
            row["VLAN"] = VlanText(pc.El("vlan"));
            row["Speed"] = PolicyValue(fc, "speed").Int();
            row["Full Duplex"] = PolicyValue(fc, "fullDuplex").Bool();
            row["Blocked"] = PolicyValue(pc, "blocked").Bool();
            row["Allow Promiscuous"] = PolicyValue(sec, "allowPromiscuous").Bool() ?? mac.Bool("allowPromiscuous");
            row["Mac Changes"] = PolicyValue(sec, "macChanges").Bool() ?? mac.Bool("macChanges");
            row["Forged Transmits"] = PolicyValue(sec, "forgedTransmits").Bool() ?? mac.Bool("forgedTransmits");
            row["Active Uplink"] = Join(order.Strs("activeUplinkPort"));
            row["Standby Uplink"] = Join(order.Strs("standbyUplinkPort"));
            row["Policy"] = PolicyValue(team, "policy").Str();
            ShapingColumns(row, pc);
            row["Reverse Policy"] = PolicyValue(team, "reversePolicy").Bool();
            row["Notify Switch"] = PolicyValue(team, "notifySwitches").Bool();
            row["Rolling Order"] = PolicyValue(team, "rollingOrder").Bool();
            row["Check Beacon"] = PolicyValue(fc, "checkBeacon").Bool();
            row["Live Port Moving"] = pol.Bool("livePortMovingAllowed");
            row["Check Duplex"] = PolicyValue(fc, "checkDuplex").Bool();
            row["Check Error %"] = PolicyValue(fc, "checkErrorPercent").Bool();
            row["Check Speed"] = PolicyValue(fc, "checkSpeed").Str();
            row["Percentage"] = PolicyValue(fc, "percentage").Int();
            row["Block Override"] = pol.Bool("blockOverrideAllowed");
            row["Config Reset"] = pol.Bool("portConfigResetAtDisconnect");
            row["Shaping Override"] = pol.Bool("shapingOverrideAllowed");
            row["Vendor Config Override"] = pol.Bool("vendorConfigOverrideAllowed");
            row["Sec. Policy Override"] = pol.Bool("securityPolicyOverrideAllowed");
            row["Teaming Override"] = pol.Bool("uplinkTeamingOverrideAllowed");
            row["Vlan Override"] = pol.Bool("vlanOverrideAllowed");
            row["Object ID"] = pg.Ref.Value;
            rows.Add(row);
        }
        return rows;
    }

    // ---------------------------------------------------------------- vMultiPath

    private List<Dictionary<string, object?>> MultiPathRows()
    {
        var rows = new List<Dictionary<string, object?>>();
        foreach (var h in _d.Hosts)
        {
            var hostName = h.Str("name") ?? h.Ref.Value;
            var luns = h.Items("config.storageDevice.scsiLun")
                .Where(l => l.Str("key") is not null)
                .ToDictionary(l => l.Str("key")!, StringComparer.Ordinal);
            // canonical disk name → datastore name (VMFS extents)
            var diskDs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var mi in h.Items("config.fileSystemVolume.mountInfo"))
            {
                var vol = mi.El("volume");
                if (vol.XsiType() != "HostVmfsVolume") continue;
                foreach (var ext in vol.Els("extent"))
                {
                    if (ext.Str("diskName") is { } dn) diskDs[dn] = vol.Str("name") ?? "";
                }
            }

            foreach (var lu in h.Get("config.storageDevice.multipathInfo").Els("lun"))
            {
                var scsi = lu.Str("lun") is { } lk ? luns.GetValueOrDefault(lk) : null;
                if (scsi is not null && scsi.Str("lunType") is { } lt && lt != "disk") continue;
                var canonical = scsi.Str("canonicalName") ?? lu.Str("id");
                var row = SheetRow();
                row["Host"] = hostName;
                row["Cluster"] = ClusterOfHost(h.Ref);
                row["Datacenter"] = DatacenterOf(h.Ref);
                row["Datastore"] = canonical is not null ? diskDs.GetValueOrDefault(canonical) : null;
                row["Disk"] = canonical;
                row["Display name"] = scsi.Str("displayName");
                row["Policy"] = lu.Str("policy.policy");
                row["Oper. State"] = Join(scsi.Strs("operationalState"));
                var paths = lu.Els("path").ToList();
                for (var i = 0; i < 8; i++)
                {
                    row[$"Path {i + 1}"] = i < paths.Count ? paths[i].Str("name") : null;
                    row[$"Path {i + 1} state"] = i < paths.Count ? paths[i].Str("pathState") ?? paths[i].Str("state") : null;
                }
                row["vStorage"] = scsi.Str("vStorageSupport");
                row["Queue depth"] = scsi.Int("queueDepth");
                row["Vendor"] = scsi.Str("vendor")?.Trim();
                row["Model"] = scsi.Str("model")?.Trim();
                row["Revision"] = scsi.Str("revision")?.Trim();
                row["Level"] = scsi.Int("scsiLevel");
                row["Serial #"] = scsi.Str("serialNumber")?.Trim();
                row["UUID"] = scsi.Str("uuid");
                row["Object ID"] = h.Ref.Value;
                rows.Add(row);
            }
        }
        return rows;
    }

    // ---------------------------------------------------------------- vLicense

    /// <summary>Masks every letter/digit except the last five characters ("XXXXX-XXXXX-XXXXX-XXXXX-AB12C").</summary>
    public static string? MaskLicenseKey(string? key)
    {
        if (string.IsNullOrEmpty(key)) return key;
        var chars = key.ToCharArray();
        for (var i = 0; i < chars.Length - 5; i++)
        {
            if (char.IsLetterOrDigit(chars[i])) chars[i] = 'X';
        }
        return new string(chars);
    }

    private static DateTimeOffset? LicenseExpiration(XElement lic) =>
        lic.Els("properties").FirstOrDefault(p => p.Str("key") == "expirationDate")?.El("value").Date();

    private List<Dictionary<string, object?>> LicenseRows()
    {
        var rows = new List<Dictionary<string, object?>>();
        foreach (var lic in _d.Licenses)
        {
            var features = lic.Els("properties").Where(p => p.Str("key") == "feature")
                .Select(p => p.El("value") is { } v ? v.Str("value") ?? v.Str("key") ?? v.Value : null);
            var row = SheetRow();
            row["Name"] = lic.Str("name");
            row["Key"] = MaskLicenseKey(lic.Str("licenseKey"));
            row["Labels"] = Join(lic.Els("labels").Select(l => $"{l.Str("key")}={l.Str("value")}"));
            row["Cost Unit"] = lic.Str("costUnit");
            row["Total"] = lic.Long("total");
            row["Used"] = lic.Long("used");
            row["Expiration Date"] = LicenseExpiration(lic) is { } e ? e : "Never";
            row["Features"] = Join(features);
            rows.Add(row);
        }
        return rows;
    }
}
