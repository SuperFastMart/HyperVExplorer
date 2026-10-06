using Avalonia.Collections;
using CommunityToolkit.Mvvm.ComponentModel;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.ViewModels;

/// <summary>A precomputed grid row: display text for binding, raw values for sorting/export.</summary>
public sealed class GridRow
{
    public GridRow(object source, string[] text, object?[] raw, (string Source, string? Host, string? Cluster) scope, VirtualMachine? vm)
    {
        Source = source;
        Text = text;
        Raw = raw;
        Scope = scope;
        Vm = vm;
        SearchText = string.Join("\u001f", text);
    }

    public object Source { get; }
    public string[] Text { get; }
    public object?[] Raw { get; }
    public (string Source, string? Host, string? Cluster) Scope { get; }
    public VirtualMachine? Vm { get; }
    public string SearchText { get; }
}

/// <summary>Sorts grid rows by a column's raw value: numbers numerically, dates chronologically, text ordinally.</summary>
public sealed class GridRowComparer(int column) : System.Collections.IComparer
{
    public int Compare(object? x, object? y)
    {
        var a = (x as GridRow)?.Raw[column];
        var b = (y as GridRow)?.Raw[column];
        if (a is null) return b is null ? 0 : -1;
        if (b is null) return 1;
        if (IsNumber(a) && IsNumber(b)) return Convert.ToDouble(a).CompareTo(Convert.ToDouble(b));
        if (a is DateTime da && b is DateTime db) return da.CompareTo(db);
        return string.Compare(TableColumn.Format(a), TableColumn.Format(b), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumber(object o) => o is int or long or double or float or decimal or short or byte;
}

public sealed partial class TableViewModel : ObservableObject
{
    private List<GridRow> _allRows = [];
    private Inventory? _builtFor;
    private Inventory _inventory = Inventory.Empty;

    public TableViewModel(TableDefinition definition, string category)
    {
        Definition = definition;
        Category = category;
        View = new DataGridCollectionView(Array.Empty<GridRow>());
    }

    public TableDefinition Definition { get; }
    public string Category { get; }
    public string Name => Definition.Name;
    public string Description => Definition.Description;

    [ObservableProperty] private int _totalCount;
    [ObservableProperty] private int _visibleCount;
    [ObservableProperty] private DataGridCollectionView _view;

    /// <summary>Per-column flag: true when at least one row has a value (used to hide always-empty VMware-only columns).</summary>
    public bool[] ColumnHasData { get; private set; } = [];

    public string Header => TotalCount > 0 ? $"{Name}  {TotalCount:N0}" : Name;

    partial void OnTotalCountChanged(int value) => OnPropertyChanged(nameof(Header));

    public IReadOnlyList<GridRow> AllRows => _allRows;

    /// <summary>Updates the row count cheaply; full rows are built only when the table is shown.</summary>
    public void SetInventory(Inventory inventory)
    {
        _inventory = inventory;
        TotalCount = Definition.Rows(inventory).Count;
    }

    /// <summary>Builds display rows if the inventory changed since the last build. Returns true when rebuilt.</summary>
    public bool EnsureBuilt()
    {
        if (ReferenceEquals(_builtFor, _inventory)) return false;
        var cols = Definition.Columns;
        var rows = Definition.Rows(_inventory);
        var hasData = new bool[cols.Count];
        var list = new List<GridRow>(rows.Count);
        foreach (var r in rows)
        {
            var raw = new object?[cols.Count];
            var text = new string[cols.Count];
            for (var i = 0; i < cols.Count; i++)
            {
                object? v;
                try { v = cols[i].GetValue(r); }
                catch (Exception) { v = null; }
                raw[i] = v;
                text[i] = TableColumn.Format(v);
                if (text[i].Length > 0) hasData[i] = true;
            }
            list.Add(new GridRow(r, text, raw, Definition.ScopeOf(r), Definition.VmOf(r)));
        }
        _allRows = list;
        ColumnHasData = hasData;
        _builtFor = _inventory;
        TotalCount = list.Count;
        return true;
    }

    public void ApplyFilter(string? search, Func<GridRow, bool>? scope)
    {
        var terms = string.IsNullOrWhiteSpace(search)
            ? []
            : search.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var filtered = _allRows.Where(r =>
            (scope is null || scope(r)) &&
            terms.All(t => r.SearchText.Contains(t, StringComparison.OrdinalIgnoreCase))).ToList();
        View = new DataGridCollectionView(filtered);
        VisibleCount = filtered.Count;
    }
}
