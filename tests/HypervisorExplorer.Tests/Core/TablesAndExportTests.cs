using System.IO.Compression;
using System.Xml.Linq;
using HypervisorExplorer.Core.Export;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Tests.Core;

public class TablesAndExportTests
{
    private static Inventory Sample() => Inventory.FromSnapshots(SampleInventory.Create());

    [Fact]
    public void Every_rvtools_sheet_has_a_table_with_identical_headers_in_order()
    {
        Assert.Equal(RvToolsSchema.Sheets.Select(s => s.Sheet), RvToolsTables.All.Select(t => t.Name));
        foreach (var (sheet, headers) in RvToolsSchema.Sheets)
            Assert.Equal(headers, RvToolsTables.Get(sheet).Columns.Select(c => c.Header));
    }

    [Fact]
    public void All_tables_evaluate_against_sample_without_throwing()
    {
        var inv = Sample();
        foreach (var table in RvToolsTables.All.Concat(ExtendedTables.All))
        foreach (var row in table.Rows(inv))
        foreach (var col in table.Columns)
            _ = col.GetValue(row);
    }

    [Fact]
    public void VInfo_maps_core_fields()
    {
        var inv = Sample();
        var table = RvToolsTables.Get("vInfo");
        var vm = inv.VirtualMachines.First(v => v.PowerState == PowerState.PoweredOn);
        var row = table.Rows(inv).First(r => table.VmOf(r) == vm);
        object? Cell(string h) => table.Columns.First(c => c.Header == h).GetValue(row);

        Assert.Equal(vm.Name, Cell("VM"));
        Assert.Equal("poweredOn", Cell("Powerstate"));
        Assert.Equal("False", Cell("Template"));
        Assert.Equal(vm.CpuCount, Cell("CPUs"));
        Assert.Equal(vm.MemoryMiB, Cell("Memory"));
        Assert.Equal((long)Math.Round(vm.ProvisionedBytes / 1048576.0), Cell("Provisioned MiB"));
        Assert.Equal(vm.Nics[0].Network, Cell("Network #1"));
        Assert.Equal(vm.Host, Cell("Host"));
        Assert.Equal(vm.SourceAddress, Cell("VI SDK Server"));
    }

    [Fact]
    public void Extra_bag_overrides_mapped_value()
    {
        var snaps = SampleInventory.Create();
        var vm = snaps[2].VirtualMachines[0];
        vm.Extra["HW version"] = 21L;
        vm.Extra["Config status"] = "yellow";
        var inv = Inventory.FromSnapshots(snaps);
        var table = RvToolsTables.Get("vInfo");
        var row = table.Rows(inv).First(r => table.VmOf(r) == vm);
        Assert.Equal(21L, table.Columns.First(c => c.Header == "HW version").GetValue(row));
        Assert.Equal("yellow", table.Columns.First(c => c.Header == "Config status").GetValue(row));
    }

    [Fact]
    public void Raw_sheet_rows_from_source_extra_are_rendered()
    {
        var inv = Sample();
        var lic = RvToolsTables.Get("vLicense");
        var rows = lic.Rows(inv);
        Assert.Single(rows);
        Assert.Equal("VMware vSphere 8 Enterprise Plus", lic.Columns.First(c => c.Header == "Name").GetValue(rows[0]));
    }

    [Fact]
    public void Xlsx_is_a_valid_package_with_all_sheets()
    {
        var inv = Sample();
        var path = Path.Combine(Path.GetTempPath(), $"hve-{Guid.NewGuid():N}.xlsx");
        try
        {
            InventoryExporter.ExportXlsx(inv, path);
            using var zip = ZipFile.OpenRead(path);
            foreach (var entry in zip.Entries.Where(e => e.Name.EndsWith(".xml") || e.Name.EndsWith(".rels")))
            {
                using var s = entry.Open();
                _ = XDocument.Load(s); // throws on malformed XML
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var wb = XDocument.Load(zip.GetEntry("xl/workbook.xml")!.Open());
            var names = wb.Descendants(ns + "sheet").Select(e => (string)e.Attribute("name")!).ToList();
            Assert.Equal(RvToolsSchema.Sheets.Select(s => s.Sheet), names.Take(27));
            Assert.Contains("hvOverview", names);

            var vinfo = XDocument.Load(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
            var rows = vinfo.Descendants(ns + "row").ToList();
            Assert.Equal(inv.VirtualMachines.Count + 1, rows.Count);
            Assert.NotNull(vinfo.Descendants(ns + "pane").SingleOrDefault());
            Assert.NotNull(vinfo.Descendants(ns + "autoFilter").SingleOrDefault());
            // CPUs column (Q) is numeric, not a shared string
            var cpuCell = rows[1].Elements(ns + "c").First(c => ((string)c.Attribute("r")!).StartsWith("Q"));
            Assert.Null(cpuCell.Attribute("t"));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Xlsx_strips_illegal_xml_characters()
    {
        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, [new XlsxSheet { Name = "T", Headers = ["A"], Rows = [["bad\u0001value <&>"]] }]);
        ms.Position = 0;
        using var zip = new ZipArchive(ms);
        var sst = XDocument.Load(zip.GetEntry("xl/sharedStrings.xml")!.Open());
        Assert.Contains(sst.Descendants().Select(e => e.Value), v => v == "badvalue <&>");
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(89, "CL")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    public void Column_names(int index, string expected) => Assert.Equal(expected, XlsxWriter.ColumnName(index));

    [Fact]
    public void Sheet_names_are_sanitised_and_unique()
    {
        var names = XlsxWriter.UniqueSheetNames(["a/b", "A_B", new string('x', 40)]);
        Assert.Equal("a_b", names[0]);
        Assert.Equal("A_B (2)", names[1]);
        Assert.Equal(31, names[2].Length);
    }

    [Theory]
    [InlineData("plain", "plain")]
    [InlineData("a,b", "\"a,b\"")]
    [InlineData("say \"hi\"", "\"say \"\"hi\"\"\"")]
    [InlineData("=cmd|'/c calc'!A1", "'=cmd|'/c calc'!A1")]
    [InlineData("-5", "-5")]
    public void Csv_escaping(string input, string expected) => Assert.Equal(expected, InventoryExporter.CsvEscape(input));

    [Fact]
    public void Json_round_trip_preserves_inventory_and_extra_values()
    {
        var snaps = SampleInventory.Create();
        snaps[2].VirtualMachines[0].Extra["Change Version"] = new DateTime(2025, 1, 2, 3, 4, 5);
        var json = InventoryJson.Serialize(snaps);
        var back = InventoryJson.Deserialize(json);

        var a = Inventory.FromSnapshots(snaps);
        var b = Inventory.FromSnapshots(back);
        Assert.Equal(a.VirtualMachines.Count, b.VirtualMachines.Count);
        Assert.Equal(a.Hosts.Count, b.Hosts.Count);
        Assert.Equal(a.Clusters.SelectMany(c => c.Nodes).Count(), b.Clusters.SelectMany(c => c.Nodes).Count());
        Assert.Equal(a.VirtualMachines.Sum(v => v.ProvisionedBytes), b.VirtualMachines.Sum(v => v.ProvisionedBytes));
        Assert.Single(RvToolsTables.Get("vLicense").Rows(b));
        var vm = b.VirtualMachines.First(v => v.Name == snaps[2].VirtualMachines[0].Name);
        Assert.IsType<DateTime>(vm.Extra["Change Version"]);
    }

    [Fact]
    public void Json_rejects_foreign_files() =>
        Assert.Throws<InvalidDataException>(() => InventoryJson.Deserialize("{\"format\":\"other\",\"snapshots\":[]}"));

    [Fact]
    public void Json_rejects_files_without_format_marker() =>
        // e.g. a Hyper-V collection file: must not load as an empty inventory
        Assert.Throws<InvalidDataException>(() => InventoryJson.Deserialize("{\"schemaVersion\":1,\"host\":{}}"));

    [Fact]
    public void Json_round_trip_keeps_extra_case_insensitive()
    {
        var snaps = SampleInventory.Create();
        snaps[2].VirtualMachines[0].Extra["Config status"] = "yellow";
        var back = InventoryJson.Deserialize(InventoryJson.Serialize(snaps));
        var vm = back[2].VirtualMachines.First(v => v.Name == snaps[2].VirtualMachines[0].Name);
        Assert.True(vm.Extra.ContainsKey("config STATUS"));
    }

    [Fact]
    public void Dates_format_invariantly_regardless_of_culture()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("th-TH");
            Assert.Equal("2024-05-01 10:30:00", TableColumn.Format(new DateTime(2024, 5, 1, 10, 30, 0)));
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public void Rekey_updates_all_references_including_extra_values()
    {
        var snap = SampleInventory.Create()[2];
        var old = snap.Source.Address;
        snap.VirtualMachines[0].Extra["VI SDK Server"] = old;
        snap.RekeySource(old + ":8443");
        Assert.All(snap.VirtualMachines, v => Assert.Equal(old + ":8443", v.SourceAddress));
        Assert.Equal(old + ":8443", snap.VirtualMachines[0].Extra["VI SDK Server"]);
        var lic = (List<Dictionary<string, object?>>)snap.Source.Extra["__sheet:vLicense"]!;
        Assert.Equal(old + ":8443", lic[0]["VI SDK Server"]);
    }

    [Fact]
    public void Html_report_renders_and_encodes()
    {
        var snaps = SampleInventory.Create();
        snaps[0].VirtualMachines[0].Annotation = "<script>alert(1)</script>";
        var html = HtmlReport.Render(Inventory.FromSnapshots(snaps));
        Assert.Contains("Hypervisor Estate Report", html);
        Assert.DoesNotContain("<script>alert(1)</script>", html);
    }

    [Fact]
    public void Health_flags_old_snapshots_and_low_datastore_space()
    {
        var inv = Sample();
        Assert.Contains(inv.Health, h => h.Message.StartsWith("Snapshot") || h.Message.Contains("free space"));
        Assert.Contains(inv.Health, h => h.Name == "vsanDatastore" && h.Severity == HealthSeverity.Warning);
    }

    [Fact]
    public void Store_upsert_replaces_source_snapshot()
    {
        var store = new InventoryStore();
        var snaps = SampleInventory.Create();
        foreach (var s in snaps) store.Upsert(s);
        var total = store.Current.VirtualMachines.Count;
        store.Upsert(new InventorySnapshot { Source = snaps[0].Source });
        Assert.Equal(total - snaps[0].VirtualMachines.Count, store.Current.VirtualMachines.Count);
        Assert.True(store.Remove(snaps[1].Source.Address));
        Assert.Equal(2, store.Current.Sources.Count);
    }
}
