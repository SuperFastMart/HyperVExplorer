using System.Globalization;
using System.Xml.Linq;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.VMware;

public sealed partial class VMwareInventoryMapper
{
    // ---------------------------------------------------------------- vCluster

    private void MapClusters()
    {
        foreach (var c in _d.ComputeResources.Where(c => c.Ref.Type == "ClusterComputeResource"))
        {
            var s = c.Get("summary");
            var cfg = c.Get("configurationEx");
            var das = cfg.El("dasConfig");
            var drs = cfg.El("drsConfig");
            var dpm = cfg.El("dpmConfigInfo");
            var cl = new ClusterInfo
            {
                Name = c.Str("name") ?? NameOf(c.Ref),
                Platform = Platform.VMware,
                SourceAddress = _addr,
                Datacenter = DatacenterOf(c.Ref),
                OverallStatus = c.Str("overallStatus") ?? s.Str("overallStatus"),
                HaEnabled = das.Bool("enabled"),
                DrsEnabled = drs.Bool("enabled"),
                NumHosts = s.Int("numHosts") ?? 0,
                NumEffectiveHosts = s.Int("numEffectiveHosts") ?? 0,
                NumCpuCores = s.Int("numCpuCores") ?? 0,
                NumCpuThreads = s.Int("numCpuThreads") ?? 0,
                TotalCpuMhz = s.Long("totalCpu") ?? 0,
                TotalMemoryBytes = s.Long("totalMemory") ?? 0,
                Id = c.Ref.Value,
            };
            var x = cl.Extra;
            x["Config status"] = c.Str("configStatus");
            x["Effective Cpu"] = s.Long("effectiveCpu");
            x["Effective Memory"] = s.Long("effectiveMemory");
            x["Num VMotions"] = s.Int("numVmotions");
            x["Failover Level"] = das.Int("failoverLevel");
            x["AdmissionControlEnabled"] = das.Bool("admissionControlEnabled");
            x["Host monitoring"] = das.Str("hostMonitoring");
            x["HB Datastore Candidate Policy"] = das.Str("hBDatastoreCandidatePolicy");
            var dvs = das.El("defaultVmSettings");
            x["Isolation Response"] = dvs.Str("isolationResponse");
            x["Restart Priority"] = dvs.Str("restartPriority");
            var tools = dvs.El("vmToolsMonitoringSettings");
            x["Cluster Settings"] = tools.Bool("clusterSettings");
            x["Max Failures"] = tools.Int("maxFailures");
            x["Max Failure Window"] = tools.Int("maxFailureWindow");
            x["Failure Interval"] = tools.Int("failureInterval");
            x["Min Up Time"] = tools.Int("minUpTime");
            x["VM Monitoring"] = das.Str("vmMonitoring") ?? tools.Str("vmMonitoring");
            x["DRS enabled"] = drs.Bool("enabled");
            x["DRS default VM behavior"] = drs.Str("defaultVmBehavior");
            x["DRS vmotion rate"] = drs.Int("vmotionRate");
            x["DPM enabled"] = dpm.Bool("enabled");
            x["DPM default behavior"] = dpm.Str("defaultDpmBehavior");
            x["DPM Host Power Action Rate"] = dpm.Int("hostPowerActionRate");
            x["Current EVC"] = s.Str("currentEVCModeKey");
            _snap.Clusters.Add(cl);

            var ha = new ClusterHa { HaEnabled = das.Bool("enabled") == true, DasConfig = das };
            foreach (var vc in cfg.Els("dasVmConfig"))
            {
                if (MoRef.From(vc.El("key")) is { } k) ha.VmOverrides[k.Value] = vc;
            }
            var groups = cfg.Els("group").Where(g => g.XsiType() == "ClusterVmGroup")
                .ToDictionary(g => g.Str("name") ?? "", g => g.Els("vm").Select(v => v.Value).ToHashSet(StringComparer.Ordinal));
            foreach (var r in cfg.Els("rule"))
            {
                var rt = r.XsiType();
                var (label, vms) = rt switch
                {
                    "ClusterAffinityRuleSpec" => ("Affinity", r.Els("vm").Select(v => v.Value).ToHashSet(StringComparer.Ordinal)),
                    "ClusterAntiAffinityRuleSpec" => ("Anti-Affinity", r.Els("vm").Select(v => v.Value).ToHashSet(StringComparer.Ordinal)),
                    "ClusterVmHostRuleInfo" => (r.Str("antiAffineHostGroupName") is not null ? "VM-Host Anti-Affinity" : "VM-Host Affinity",
                        groups.GetValueOrDefault(r.Str("vmGroupName") ?? "") ?? []),
                    "ClusterDependencyRuleInfo" => ("VM-VM Dependency",
                        (groups.GetValueOrDefault(r.Str("vmGroup") ?? "") ?? []).Concat(groups.GetValueOrDefault(r.Str("dependsOnVmGroup") ?? "") ?? [])
                        .ToHashSet(StringComparer.Ordinal)),
                    _ => (rt ?? "Rule", new HashSet<string>()),
                };
                if (r.Bool("enabled") != false) ha.Rules.Add(new ClusterRule(r.Str("name") ?? "", label, vms));
            }
            _clusterHa[c.Ref.Value] = ha;
        }
    }

    private void FinishClusters()
    {
        foreach (var cl in _snap.Clusters)
        {
            var c = _d.ComputeResources.FirstOrDefault(r => r.Ref.Value == cl.Id);
            if (c is null) continue;
            foreach (var hr in c.Refs("host"))
            {
                var h = _hostObjs.GetValueOrDefault(hr.Value);
                cl.Nodes.Add(new ClusterNode
                {
                    Name = NameOf(hr),
                    Id = hr.Value,
                    State = h?.Str("runtime.connectionState"),
                    DrainStatus = h?.Bool("runtime.inMaintenanceMode") == true ? "inMaintenance" : null,
                });
            }
        }
    }

    // ---------------------------------------------------------------- vHost & friends

    private void MapHosts()
    {
        var assigned = _d.LicenseAssignments
            .GroupBy(a => a.Str("entityId") ?? "")
            .ToDictionary(g => g.Key, g => Join(g.Select(a => a.Str("assignedLicense.name"))), StringComparer.Ordinal);
        var swappedByHost = _d.VirtualMachines
            .GroupBy(v => v.Ref1("runtime.host")?.Value ?? "")
            .ToDictionary(g => g.Key, g => g.Sum(v => v.Long("summary.quickStats.swappedMemory") ?? 0), StringComparer.Ordinal);

        foreach (var o in _d.Hosts)
        {
            try
            {
                _snap.Hosts.Add(MapHost(o, assigned.GetValueOrDefault(o.Ref.Value), swappedByHost.GetValueOrDefault(o.Ref.Value)));
            }
            catch (Exception ex)
            {
                _snap.Warnings.Add($"Host '{o.Str("name") ?? o.Ref.Value}' could not be mapped: {ex.Message}");
            }
        }
    }

    private HostSystem MapHost(VimObject o, string? licenses, long vmSwappedMiB)
    {
        var hw = o.Get("summary.hardware");
        var qs = o.Get("summary.quickStats");
        var net = o.Get("config.network");
        var dt = o.Get("config.dateTimeInfo");
        var dns = net.El("dnsConfig");
        var sockets = hw.Int("numCpuPkgs");
        var cores = hw.Int("numCpuCores");
        var mhz = hw.Int("cpuMhz");
        var ident = hw.Els("otherIdentifyingInfo").Concat(o.Get("hardware.systemInfo").Els("otherIdentifyingInfo")).ToList();
        string? IdentValue(string key) => ident.FirstOrDefault(i => i.Str("identifierType.key") == key)?.Str("identifierValue");

        var h = new HostSystem
        {
            Name = o.Str("name") ?? NameOf(o.Ref),
            Platform = Platform.VMware,
            SourceAddress = _addr,
            Datacenter = DatacenterOf(o.Ref),
            Cluster = ClusterOfHost(o.Ref),
            Status = o.Str("overallStatus"),
            InMaintenance = o.Bool("runtime.inMaintenanceMode") ?? false,
            CpuModel = hw.Str("cpuModel"),
            CpuMhz = mhz,
            CpuSockets = sockets,
            CoresPerSocket = sockets is > 0 && cores is not null ? cores / sockets : null,
            CpuCores = cores,
            CpuThreads = hw.Int("numCpuThreads"),
            HyperThreadingActive = o.Bool("config.hyperThread.active"),
            CpuUsagePercent = qs.Long("overallCpuUsage") is { } used && mhz is > 0 && cores is > 0
                ? used * 100.0 / (mhz.Value * (double)cores.Value) : null,
            MemoryBytes = hw.Long("memorySize"),
            MemoryUsedBytes = qs.Long("overallMemoryUsage") * 1024 * 1024,
            Version = o.Str("summary.config.product.fullName"),
            KernelVersion = o.Str("summary.config.product.build"),
            BootTime = o.Date("runtime.bootTime"),
            Vendor = hw.Str("vendor"),
            Model = hw.Str("model"),
            SerialNumber = o.Str("hardware.systemInfo.serialNumber") is { Length: > 0 } sn ? sn
                : IdentValue("SerialNumberTag") ?? IdentValue("EnclosureSerialNumberTag") ?? IdentValue("ServiceTag"),
            BiosVendor = o.Str("hardware.biosInfo.vendor"),
            BiosVersion = o.Str("hardware.biosInfo.biosVersion"),
            BiosDate = o.Date("hardware.biosInfo.releaseDate"),
            Uuid = hw.Str("uuid"),
            DnsServers = dns.Strs("address"),
            Domain = dns.Str("domainName"),
            DnsSearch = dns.Strs("searchDomain"),
            NtpServers = dt.El("ntpConfig").Strs("server"),
            TimeZone = dt.Str("timeZone.name") ?? dt.Str("timeZone.key"),
        };

        var x = h.Extra;
        x["Config status"] = o.Str("configStatus");
        x["in Quarantine Mode"] = o.Bool("runtime.inQuarantineMode");
        x["HT Available"] = o.Bool("config.hyperThread.available");
        x["CPU usage %"] = h.CpuUsagePercent is { } p ? (object)Math.Round(p) : null;
        x["VMotion support"] = o.Bool("capability.vmotionSupported");
        x["Storage VMotion support"] = o.Bool("capability.storageVMotionSupported");
        x["Current EVC"] = o.Str("summary.currentEVCModeKey");
        x["Max EVC"] = o.Str("summary.maxEVCModeKey");
        x["Assigned License(s)"] = licenses;
        x["Current CPU power man. policy"] = o.Str("hardware.cpuPowerManagementInfo.currentPolicy");
        x["Supported CPU power man."] = o.Str("hardware.cpuPowerManagementInfo.hardwareSupport");
        x["Host Power Policy"] = o.Str("config.powerSystemInfo.currentPolicy.name") ?? o.Str("config.powerSystemInfo.currentPolicy.shortName");
        x["DHCP"] = dns.Bool("dhcp");
        x["NTPD running"] = o.Get("config.service").Els("service").FirstOrDefault(s => s.Str("key") == "ntpd")?.Bool("running");
        x["Time Zone"] = dt.Str("timeZone.description") ?? dt.Str("timeZone.key");
        x["Time Zone Name"] = dt.Str("timeZone.name");
        x["GMT Offset"] = dt.Int("timeZone.gmtOffset");
        x["Service tag"] = IdentValue("ServiceTag");
        x["OEM specific string"] = Join(ident.Where(i => i.Str("identifierType.key") == "OemSpecificString").Select(i => i.Str("identifierValue")));
        x["Object ID"] = o.Ref.Value;
        x["Datacenter"] = h.Datacenter;

        x["VM Memory Swapped"] = vmSwappedMiB;

        MapHostNetwork(h, net);
        MapHostStorage(h, o);
        return h;
    }

    private void MapHostNetwork(HostSystem h, XElement? net)
    {
        var pnicDevice = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var p in net.Els("pnic"))
        {
            if (p.Str("key") is { } k && p.Str("device") is { } d) pnicDevice[k] = d;
        }
        var nicSwitch = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var vs in net.Els("vswitch"))
        {
            foreach (var pk in vs.Strs("pnic"))
            {
                if (pnicDevice.TryGetValue(pk, out var dev)) nicSwitch[dev] = vs.Str("name") ?? "";
            }
        }
        foreach (var ps in net.Els("proxySwitch"))
        {
            foreach (var pk in ps.Strs("pnic"))
            {
                if (pnicDevice.TryGetValue(pk, out var dev)) nicSwitch[dev] = ps.Str("dvsName") ?? "";
            }
        }

        foreach (var p in net.Els("pnic"))
        {
            var dev = p.Str("device") ?? "";
            var link = p.El("linkSpeed");
            var nic = new HostNic
            {
                Name = dev,
                Driver = p.Str("driver"),
                SpeedMbps = link.Long("speedMb"),
                FullDuplex = link.Bool("duplex"),
                Mac = p.Str("mac"),
                Switch = nicSwitch.GetValueOrDefault(dev),
                Pci = p.Str("pci"),
                Status = link is null ? "Down" : "Up",
            };
            nic.Extra["Speed"] = link.Long("speedMb") ?? 0;
            nic.Extra["WakeOn"] = p.Bool("wakeOnLanSupported");
            nic.Extra["Uplink port"] = null;
            h.Nics.Add(nic);
        }
        // Uplink names on distributed switches (proxySwitch.spec.backing.pnicSpec[].uplinkPortKey → uplinkPort names).
        foreach (var ps in net.Els("proxySwitch"))
        {
            var portNames = ps.Els("uplinkPort").ToDictionary(u => u.Str("key") ?? "", u => u.Str("value"), StringComparer.Ordinal);
            foreach (var spec in ps.At("spec.backing").Els("pnicSpec"))
            {
                var nic = h.Nics.FirstOrDefault(n => n.Name == spec.Str("pnicDevice"));
                if (nic is not null && spec.Str("uplinkPortKey") is { } upk)
                    nic.Extra["Uplink port"] = portNames.GetValueOrDefault(upk) ?? upk;
            }
        }

        foreach (var vs in net.Els("vswitch"))
        {
            var policy = vs.At("spec.policy");
            var sw = new VirtualSwitch
            {
                Name = vs.Str("name") ?? "",
                Type = "Standard",
                Uplinks = vs.Strs("pnic").Select(k => pnicDevice.GetValueOrDefault(k) ?? k).ToList(),
                Mtu = vs.Int("mtu"),
                Ports = vs.Int("numPorts"),
            };
            ApplyNetworkPolicy(sw.Extra, policy);
            sw.Extra["Free Ports"] = vs.Int("numPortsAvailable");
            sw.Extra["MTU"] = vs.Int("mtu");
            h.Switches.Add(sw);
        }

        foreach (var pg in net.Els("portgroup"))
        {
            var port = new PortGroup
            {
                Name = pg.Str("spec.name") ?? "",
                Switch = pg.Str("spec.vswitchName"),
                Vlan = pg.Int("spec.vlanId"),
            };
            ApplyNetworkPolicy(port.Extra, pg.El("computedPolicy") ?? pg.At("spec.policy"));
            h.PortGroups.Add(port);
        }

        var gateway = net.Str("ipRouteConfig.defaultGateway");
        var gateway6 = net.Str("ipRouteConfig.ipV6DefaultGateway");
        foreach (var vn in net.Els("vnic"))
        {
            var spec = vn.El("spec");
            var pgName = vn.Str("portgroup");
            if (string.IsNullOrEmpty(pgName) && spec.Str("distributedVirtualPort.portgroupKey") is { } pgKey)
                pgName = _pgByKey.TryGetValue(pgKey, out var dpg) ? dpg.Str("name") : pgKey;
            var v6 = spec.At("ip.ipV6Config").Els("ipV6Address")
                .Select(a => a.Str("ipAddress") is { } ip ? $"{ip}/{a.Str("prefixLength")}" : null);
            var ipi = new HostIpInterface
            {
                Name = vn.Str("device") ?? "",
                PortGroup = pgName,
                Mac = spec.Str("mac"),
                Dhcp = spec.Bool("ip.dhcp"),
                Ipv4 = spec.Str("ip.ipAddress"),
                SubnetMask = spec.Str("ip.subnetMask"),
                Gateway = gateway,
                Ipv6 = Join(v6),
                Mtu = spec.Int("mtu"),
            };
            ipi.Extra["IP 6 Gateway"] = gateway6;
            h.IpInterfaces.Add(ipi);
        }
    }

    /// <summary>Fills vSwitch/vPort policy columns from a HostNetworkPolicy.</summary>
    private static void ApplyNetworkPolicy(Dictionary<string, object?> x, XElement? policy)
    {
        var sec = policy.El("security");
        var shaping = policy.El("shapingPolicy");
        var team = policy.El("nicTeaming");
        var off = policy.El("offloadPolicy");
        x["Promiscuous Mode"] = sec.Bool("allowPromiscuous");
        x["Mac Changes"] = sec.Bool("macChanges");
        x["Forged Transmits"] = sec.Bool("forgedTransmits");
        x["Traffic Shaping"] = shaping.Bool("enabled");
        x["Width"] = shaping.Long("averageBandwidth");
        x["Peak"] = shaping.Long("peakBandwidth");
        x["Burst"] = shaping.Long("burstSize");
        x["Policy"] = team.Str("policy");
        x["Reverse Policy"] = team.Bool("reversePolicy");
        x["Notify Switch"] = team.Bool("notifySwitches");
        x["Rolling Order"] = team.Bool("rollingOrder");
        x["Offload"] = off.Bool("csumOffload");
        x["TSO"] = off.Bool("tcpSegmentation");
        x["Zero Copy Xmit"] = off.Bool("zeroCopyXmit");
    }

    private static string HbaType(string? xsiType) => xsiType switch
    {
        "HostBlockHba" => "Block SCSI",
        "HostFibreChannelHba" => "Fibre Channel",
        "HostFibreChannelOverEthernetHba" => "Fibre Channel over Ethernet",
        "HostInternetScsiHba" => "iSCSI",
        "HostParallelScsiHba" => "Parallel SCSI",
        "HostSerialAttachedHba" => "SAS",
        "HostPcieHba" => "PCIe",
        "HostRdmaHba" => "RDMA",
        "HostTcpHba" => "TCP",
        null => "Unknown",
        _ => xsiType.StartsWith("Host", StringComparison.Ordinal) ? xsiType[4..] : xsiType,
    };

    /// <summary>Formats an xsd:long WWN as colon-separated hex ("20:00:00:25:b5:aa:00:0f").</summary>
    public static string? FormatWwn(long? wwn)
    {
        if (wwn is null) return null;
        var hex = unchecked((ulong)wwn.Value).ToString("x16", CultureInfo.InvariantCulture);
        return string.Join(":", Enumerable.Range(0, 8).Select(i => hex.Substring(i * 2, 2)));
    }

    private static void MapHostStorage(HostSystem h, VimObject o)
    {
        foreach (var a in o.Items("config.storageDevice.hostBusAdapter"))
        {
            var t = a.XsiType();
            string? wwn = t switch
            {
                "HostFibreChannelHba" or "HostFibreChannelOverEthernetHba" =>
                    $"{FormatWwn(a.Long("nodeWorldWideName"))} {FormatWwn(a.Long("portWorldWideName"))}".Trim(),
                "HostInternetScsiHba" => a.Str("iScsiName"),
                "HostSerialAttachedHba" => a.Str("nodeWorldWideName"),
                _ => null,
            };
            var hba = new StorageAdapter
            {
                Device = a.Str("device") ?? "",
                Type = HbaType(t),
                Status = a.Str("status"),
                Driver = a.Str("driver"),
                Model = a.Str("model"),
                Wwn = string.IsNullOrWhiteSpace(wwn) ? null : wwn,
                Pci = a.Str("pci"),
            };
            hba.Extra["Bus"] = a.Int("bus");
            h.StorageAdapters.Add(hba);
        }
    }

    // ---------------------------------------------------------------- vDatastore

    private void MapDatastores()
    {
        var poweredOn = _d.VirtualMachines.Where(v => v.Str("runtime.powerState") == "poweredOn")
            .Select(v => v.Ref.Value).ToHashSet(StringComparer.Ordinal);
        foreach (var o in _d.Datastores)
        {
            var s = o.Get("summary");
            var info = o.Get("info");
            var vmfs = info.El("vmfs");
            var nas = info.El("nas");
            var hosts = o.Items("host").Select(m => MoRef.From(m.El("key"))).OfType<MoRef>().ToList();
            var cap = s.Long("capacity");
            var free = s.Long("freeSpace");
            var vms = o.Refs("vm");
            var extents = vmfs.Els("extent").Select(e => e.Str("diskName")).ToList();
            string? address = nas is not null
                ? $"{Join(nas.Strs("remoteHostNames")) ?? nas.Str("remoteHost")}:{nas.Str("remotePath")}"
                : Join(extents);

            var clusters = hosts.Select(ClusterOfHost).Distinct().ToList();
            var parent = o.Ref1("parent") ?? ParentOf(o.Ref);
            var ds = new Datastore
            {
                Name = o.Str("name") ?? s.Str("name") ?? NameOf(o.Ref),
                Platform = Platform.VMware,
                SourceAddress = _addr,
                Cluster = clusters.Count == 1 ? clusters[0] : null,
                Type = s.Str("type"),
                Address = address,
                Accessible = s.Bool("accessible") ?? true,
                Shared = s.Bool("multipleHostAccess"),
                CapacityBytes = cap,
                FreeBytes = free,
                Content = s.Str("url"),
                Hosts = hosts.Select(NameOf).ToList(),
            };
            var x = ds.Extra;
            x["Config status"] = o.Str("configStatus");
            x["# VMs total"] = vms.Count;
            x["# VMs"] = vms.Count(v => poweredOn.Contains(v.Value));
            if (cap is not null && free is not null)
                x["Provisioned MiB"] = Mib(cap - free + (s.Long("uncommitted") ?? 0));
            x["SIOC enabled"] = o.Bool("iormConfiguration.enabled");
            x["SIOC Threshold"] = o.Int("iormConfiguration.congestionThreshold");
            x["Cluster name"] = parent?.Type == "StoragePod" ? NameOf(parent) : null;
            x["Block size"] = vmfs.Int("blockSizeMb");
            x["Max Blocks"] = vmfs.Long("maxBlocks");
            x["# Extents"] = vmfs is null ? null : extents.Count;
            x["Major Version"] = vmfs.Int("majorVersion");
            x["Version"] = vmfs.Str("version") ?? nas.Str("type");
            x["VMFS Upgradeable"] = vmfs.Bool("vmfsUpgradable");
            x["MHA"] = s.Bool("multipleHostAccess");
            x["URL"] = s.Str("url");
            x["Object ID"] = o.Ref.Value;
            _snap.Datastores.Add(ds);
        }
    }
}
