using System.IO.Compression;
using System.Text;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Core.Export;

public sealed class XlsxExportOptions
{
    /// <summary>Include RVTools sheets that have no rows (RVTools always writes all 27).</summary>
    public bool IncludeEmptySheets { get; init; } = true;
    /// <summary>Append the extended (non-RVTools) sheets after the RVTools ones.</summary>
    public bool IncludeExtendedSheets { get; init; } = true;
}

/// <summary>Writes the inventory to RVTools-compatible XLSX, CSV, and HTML using the shared table definitions.</summary>
public static class InventoryExporter
{
    public static void ExportXlsx(Inventory inventory, string path, XlsxExportOptions? options = null)
    {
        options ??= new XlsxExportOptions();
        var tables = RvToolsTables.All.AsEnumerable();
        if (options.IncludeExtendedSheets) tables = tables.Concat(ExtendedTables.All);

        var sheets = new List<XlsxSheet>();
        foreach (var table in tables)
        {
            var rows = table.Rows(inventory);
            var isRvTools = RvToolsTables.All.Contains(table);
            if (rows.Count == 0 && !(options.IncludeEmptySheets && isRvTools)) continue;
            sheets.Add(ToSheet(table, rows));
        }
        XlsxWriter.Write(path, sheets);
    }

    public static XlsxSheet ToSheet(TableDefinition table, IReadOnlyList<object> rows) => new()
    {
        Name = table.Name,
        Headers = table.Columns.Select(c => c.Header).ToList(),
        Rows = rows.Select(r => table.Columns.Select(c => c.GetValue(r)).ToArray()).ToList(),
        FrozenColumns = table.FrozenColumns,
        AutoFilter = table.AutoFilter,
    };

    /// <summary>Writes one table (optionally pre-filtered rows and a column subset) as UTF-8 CSV with BOM for Excel.</summary>
    public static void ExportCsv(TableDefinition table, IReadOnlyList<object> rows, string path,
        IReadOnlyList<TableColumn>? columns = null)
    {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(true));
        WriteCsv(writer, table, rows, columns);
    }

    public static void WriteCsv(TextWriter writer, TableDefinition table, IReadOnlyList<object> rows,
        IReadOnlyList<TableColumn>? columns = null)
    {
        columns ??= table.Columns;
        writer.WriteLine(string.Join(",", columns.Select(c => CsvEscape(c.Header))));
        foreach (var row in rows)
            writer.WriteLine(string.Join(",", columns.Select(c => CsvEscape(TableColumn.Format(c.GetValue(row))))));
    }

    /// <summary>Writes every non-empty table as a CSV inside one zip archive.</summary>
    public static void ExportCsvZip(Inventory inventory, string zipPath)
    {
        using var fs = new FileStream(zipPath, FileMode.Create);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
        foreach (var table in RvToolsTables.All.Concat(ExtendedTables.All))
        {
            var rows = table.Rows(inventory);
            if (rows.Count == 0) continue;
            var entry = zip.CreateEntry(table.Name + ".csv", CompressionLevel.Optimal);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, new UTF8Encoding(true));
            WriteCsv(writer, table, rows);
        }
    }

    public static string CsvEscape(string value)
    {
        if (value.Length == 0) return "";
        // Neutralise spreadsheet formula injection from guest-controlled strings (VM notes, hostnames).
        if (value[0] is '=' or '+' or '-' or '@' && !double.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _))
            value = "'" + value;
        return value.IndexOfAny([',', '"', '\n', '\r']) >= 0 ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;
    }
}
