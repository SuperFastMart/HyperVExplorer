using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Core.Tables;

public enum CellKind
{
    Text,
    Number,
    Date,
    Bool,
}

public sealed class TableColumn
{
    public TableColumn(string header, CellKind kind, Func<object, object?> getter)
    {
        Header = header;
        Kind = kind;
        Getter = getter;
    }

    public string Header { get; }
    public CellKind Kind { get; }
    private Func<object, object?> Getter { get; }

    /// <summary>True when the column maps to collected data (vs. a VMware-only column that is always blank).</summary>
    public bool IsMapped { get; init; } = true;

    public object? GetValue(object row) => Normalize(Getter(row), Kind);

    private static object? Normalize(object? value, CellKind kind) => value switch
    {
        null => null,
        string s when s.Length == 0 => null,
        bool b => b ? "True" : "False",
        DateTimeOffset dto => dto.LocalDateTime,
        Enum e => e.ToString(),
        _ => value,
    };

    /// <summary>Formats a value for display/CSV, matching how the cell appears in Excel.</summary>
    public static string Format(object? value) => value switch
    {
        null => "",
        DateTime dt => dt.TimeOfDay == TimeSpan.Zero ? dt.ToString("yyyy-MM-dd") : dt.ToString("yyyy-MM-dd HH:mm:ss"),
        double d => d.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture),
        float f => f.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture),
        decimal m => m.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture),
        IFormattable fmt => fmt.ToString(null, System.Globalization.CultureInfo.InvariantCulture),
        _ => value.ToString() ?? "",
    };
}

/// <summary>
/// A tabular view over the inventory. One definition drives the UI grid, CSV, XLSX and HTML outputs,
/// so every surface shows identical columns and values.
/// </summary>
public sealed class TableDefinition
{
    public required string Name { get; init; }
    public string Description { get; init; } = "";
    public required IReadOnlyList<TableColumn> Columns { get; init; }
    public required Func<Inventory, IReadOnlyList<object>> Rows { get; init; }
    /// <summary>The VM a row belongs to (for details panel / selection), if any.</summary>
    public Func<object, VirtualMachine?> VmOf { get; init; } = _ => null;
    /// <summary>(source address, host name) the row belongs to, for tree filtering.</summary>
    public Func<object, (string Source, string? Host, string? Cluster)> ScopeOf { get; init; } = _ => ("", null, null);
    /// <summary>Sheet-level frozen columns in XLSX (RVTools freezes the first column).</summary>
    public int FrozenColumns { get; init; } = 1;
    public bool AutoFilter { get; init; } = true;
}

/// <summary>
/// Builds a table whose columns follow an RVTools sheet's exact header list. Unmapped headers become
/// blank columns; values in an object's <see cref="InventoryObject.Extra"/> bag override mapped values.
/// </summary>
public sealed class RvSheetBuilder<T> where T : class
{
    private readonly string _sheet;
    private readonly HashSet<string> _headers;
    private readonly Dictionary<string, (CellKind Kind, Func<T, object?> Get)> _map = new(StringComparer.Ordinal);
    private readonly List<Func<T, InventoryObject?>> _extraSources = [];

    public RvSheetBuilder(string sheet)
    {
        _sheet = sheet;
        _headers = RvToolsSchema.HeadersFor(sheet).ToHashSet(StringComparer.Ordinal);
    }

    public IReadOnlySet<string> Headers => _headers;

    /// <summary>Objects whose Extra bags are consulted (first match wins), e.g. the disk, then its VM.</summary>
    public RvSheetBuilder<T> ExtrasFrom(params Func<T, InventoryObject?>[] sources)
    {
        _extraSources.AddRange(sources);
        return this;
    }

    public RvSheetBuilder<T> Text(string header, Func<T, object?> get) => Add(header, CellKind.Text, get);
    public RvSheetBuilder<T> Num(string header, Func<T, object?> get) => Add(header, CellKind.Number, get);
    public RvSheetBuilder<T> Date(string header, Func<T, object?> get) => Add(header, CellKind.Date, get);
    public RvSheetBuilder<T> Bool(string header, Func<T, bool?> get) => Add(header, CellKind.Bool, r => get(r));

    private RvSheetBuilder<T> Add(string header, CellKind kind, Func<T, object?> get)
    {
        if (!_headers.Contains(header))
            throw new InvalidOperationException($"RVTools sheet '{_sheet}' has no column '{header}'.");
        _map[header] = (kind, get);
        return this;
    }

    public TableDefinition Build(
        Func<Inventory, IEnumerable<T>> rows,
        string description,
        Func<T, VirtualMachine?>? vmOf = null,
        Func<T, (string, string?, string?)>? scopeOf = null)
    {
        var columns = new List<TableColumn>();
        foreach (var header in RvToolsSchema.HeadersFor(_sheet))
        {
            var mapped = _map.TryGetValue(header, out var m);
            var kind = mapped ? m.Kind : CellKind.Text;
            var get = mapped ? m.Get : null;
            var h = header;
            columns.Add(new TableColumn(header, kind, row =>
            {
                var r = (T)row;
                foreach (var src in _extraSources)
                {
                    if (src(r) is { } obj && obj.Extra.TryGetValue(h, out var extra))
                        return extra;
                }
                return get?.Invoke(r);
            })
            { IsMapped = mapped || _extraSources.Count > 0 });
        }

        return new TableDefinition
        {
            Name = _sheet,
            Description = description,
            Columns = columns,
            Rows = inv => rows(inv).Cast<object>().ToList(),
            VmOf = vmOf is null ? _ => null : o => vmOf((T)o),
            ScopeOf = scopeOf is null ? _ => ("", null, null) : o => scopeOf((T)o),
        };
    }
}
