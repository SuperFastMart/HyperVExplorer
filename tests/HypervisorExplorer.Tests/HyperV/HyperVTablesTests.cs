using HypervisorExplorer.Collectors.HyperV;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Tests.HyperV;

/// <summary>Pushes a mapped Hyper-V snapshot through the store and every table definition.</summary>
public class HyperVTablesTests
{
    private static Inventory Load()
    {
        var store = new InventoryStore();
        store.Upsert(HyperVJsonMapper.Map(HyperVFixtures.ClusterData, new ConnectionRequest
        {
            Platform = Platform.HyperV,
            Address = "hvclu01.contoso.local",
            Group = "London DC",
        }));
        store.Upsert(HyperVJsonMapper.Map(HyperVFixtures.StandaloneHost, new ConnectionRequest
        {
            Platform = Platform.HyperV,
            Address = "HVSTD01",
        }));
        return store.Current;
    }

    private static Dictionary<string, string> Row(TableDefinition t, object row) =>
        t.Columns.ToDictionary(c => c.Header, c => TableColumn.Format(c.GetValue(row)));

    private static List<Dictionary<string, string>> Rows(Inventory inv, TableDefinition t) =>
        t.Rows(inv).Select(r => Row(t, r)).ToList();

    [Fact]
    public void EveryTable_RendersWithoutErrors()
    {
        var inv = Load();
        foreach (var table in RvToolsTables.All.Concat(ExtendedTables.All))
        {
            foreach (var row in table.Rows(inv))
            {
                foreach (var col in table.Columns) TableColumn.Format(col.GetValue(row));
                table.ScopeOf(row);
                table.VmOf(row);
            }
        }
    }

    [Fact]
    public void VInfo_HasHyperVValues()
    {
        var inv = Load();
        var rows = Rows(inv, RvToolsTables.Get("vInfo"));
        Assert.Equal(4, rows.Count);
        var sql = rows.Single(r => r["VM"] == "SQL01");
        Assert.Equal("poweredOn", sql["Powerstate"]);
        Assert.Equal("green", sql["Heartbeat"]);
        Assert.Equal("sql01.contoso.local", sql["DNS Name"]);
        Assert.Equal("8", sql["CPUs"]);
        Assert.Equal("16384", sql["Memory"]);
        Assert.Equal("2", sql["Disks"]);
        Assert.Equal("10.0.20.50", sql["Primary IP Address"]);
        Assert.Equal("SETswitch", sql["Network #1"]);
        Assert.Equal("efi", sql["Firmware"]);
        Assert.Equal("True", sql["EFI Secure boot"]);
        Assert.Equal("10.0", sql["HW version"]);
        Assert.Equal("high", sql["HA Restart Priority"]);
        Assert.Equal("True", sql["DAS protection"]);
        Assert.Equal("120000", sql["Boot delay"]);
        Assert.Equal("London DC", sql["Datacenter"]);
        Assert.Equal("HVCLU01", sql["Cluster"]);
        Assert.Equal("HV01", sql["Host"]);
        Assert.Equal("Windows Server 2022 Datacenter", sql["OS according to the VMware Tools"]);
        Assert.Equal("2b9f1e4a-6c7d-4e8f-9a0b-1c2d3e4f5a6b", sql["SMBIOS UUID"]);
        Assert.Equal("6f1d3a2e-8b4c-4d5e-9f00-112233445566", sql["VM UUID"]);
        Assert.Equal("Microsoft Hyper-V Microsoft Windows Server 2022 Datacenter", sql["VI SDK Server type"]);
        Assert.Equal("WinRM", sql["VI SDK API Version"]);
        Assert.Equal("hvclu01.contoso.local", sql["VI SDK Server"]);
        Assert.Equal(((136365211648L + 536870912000L) / 1048576).ToString(), sql["Provisioned MiB"]);

        Assert.Equal("suspended", rows.Single(r => r["VM"] == "DC01")["Powerstate"]);
    }

    [Fact]
    public void VHost_HasHostValuesAndVmCounts()
    {
        var inv = Load();
        var rows = Rows(inv, RvToolsTables.Get("vHost"));
        var hv01 = rows.Single(r => r["Host"] == "HV01");
        Assert.Equal("HVCLU01", hv01["Cluster"]);
        Assert.Equal("London DC", hv01["Datacenter"]);
        Assert.Equal("2", hv01["# CPU"]);
        Assert.Equal("16", hv01["Cores per CPU"]);
        Assert.Equal("32", hv01["# Cores"]);
        Assert.Equal("True", hv01["HT Active"]);
        Assert.Equal("524288", hv01["# Memory"]);
        Assert.Equal("2", hv01["# VMs total"]);
        Assert.Equal("1", hv01["# VMs"]);
        Assert.Equal("3", hv01["# NICs"]);
        Assert.Equal("2", hv01["# HBAs"]);
        Assert.Equal("Microsoft Windows Server 2022 Datacenter 10.0.20348", hv01["ESX Version"]);
        Assert.Equal("10.0.0.10, 10.0.0.11", hv01["DNS Servers"]);
        Assert.Equal("dc01.contoso.local", hv01["NTP Server(s)"]);
        Assert.Equal("ABC1234", hv01["Serial number"]);
        Assert.Equal("True", hv01["VMotion support"]);
        Assert.Equal("(UTC+00:00) Dublin, Edinburgh, Lisbon, London", hv01["Time Zone Name"]);
    }

    [Fact]
    public void ClusterSheets_HaveNodesNetworksAndHa()
    {
        var inv = Load();
        var nodes = Rows(inv, ExtendedTables.All.Single(t => t.Name == "hvClusterNodes"));
        Assert.Equal(3, nodes.Count);
        var hv02 = nodes.Single(r => r["Node"] == "HV02");
        Assert.Equal("HVCLU01", hv02["Cluster"]);
        Assert.Equal("Up", hv02["State"]);
        Assert.Equal("10.0.10.22", hv02["Address"]);
        Assert.Equal("1", hv02["Votes"]);
        Assert.Equal("NodeAndFileShareMajority", hv02["Quorum"]);
        Assert.Equal("File Share Witness", hv02["Witness"]);
        Assert.Equal("Down", nodes.Single(r => r["Node"] == "HV03")["State"]);

        var nets = Rows(inv, ExtendedTables.All.Single(t => t.Name == "hvClusterNetworks"));
        Assert.Equal("Cluster only", nets.Single(r => r["Network"] == "LiveMigration")["Role"]);

        var ha = Rows(inv, ExtendedTables.All.Single(t => t.Name == "hvHA"));
        Assert.DoesNotContain(ha, r => r["VM"] == "DC01"); // standalone host VM
        var sql = ha.Single(r => r["VM"] == "SQL01");
        Assert.Equal("HV01, HV02", sql["Preferred owners"]);
        Assert.Equal("High", sql["Priority"]);
        Assert.Equal("True", sql["HA protected"]);

        var cluster = Rows(inv, RvToolsTables.Get("vCluster")).Single();
        Assert.Equal("3", cluster["NumHosts"]);
        Assert.Equal("2", cluster["numEffectiveHosts"]);
        Assert.Equal("56", cluster["NumCpuCores"]);
    }

    [Fact]
    public void DiskAndDatastoreSheets_Resolve()
    {
        var inv = Load();
        var disks = Rows(inv, RvToolsTables.Get("vDisk"));
        var sqlData = disks.Single(r => r["VM"] == "SQL01" && r["Disk"] == "Hard disk 2");
        Assert.Equal("True", sqlData["Thin"]);
        Assert.Equal("vhdx", sqlData["Disk Mode"]);
        Assert.Equal("SCSI 0:1", sqlData["Controller"]);
        Assert.Equal("1", sqlData["SCSI Unit #"]);
        Assert.Equal("True", disks.Single(r => r["VM"] == "WEB01" && r["Disk"] == "Hard disk 2")["Raw"]);

        var ds = Rows(inv, RvToolsTables.Get("vDatastore"));
        var csv1 = ds.Single(r => r["Name"] == "Cluster Disk 1");
        Assert.Equal("1", csv1["# VMs total"]);
        Assert.Equal("3", csv1["# Hosts"]);
        Assert.Equal("50", csv1["Free %"]);

        var net = Rows(inv, RvToolsTables.Get("vNetwork")).Single(r => r["VM"] == "SQL01");
        Assert.Equal("SETswitch (VLAN 20)", net["Network"]);
        Assert.Equal("2001:db8:20::50, fe80::1c2d:3e4f:5a6b:7c8d".Split(", ").OrderBy(x => x), net["IPv6 Address"].Split(", ").OrderBy(x => x));

        var source = Rows(inv, RvToolsTables.Get("vSource")).Single(r => r["VI SDK Server"] == "hvclu01.contoso.local");
        Assert.Equal("WinRM", source["API type"]);
        Assert.Equal("Microsoft Corporation", source["Vendor"]);
        Assert.Equal("Hyper-V Failover Cluster", source["Product line"]);

        var health = Rows(inv, RvToolsTables.Get("vHealth"));
        Assert.Contains(health, r => r["Name"] == "SQL01" && r["Message"].Contains("Media mounted"));
        Assert.Contains(health, r => r["Name"] == "DC01" && r["Message"].Contains("Snapshot"));
    }
}
