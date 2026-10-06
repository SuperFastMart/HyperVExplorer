using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using HypervisorExplorer.App.Services;
using HypervisorExplorer.App.ViewModels;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.App.Views;

public partial class MainWindow : Window, IDialogService
{
    private MainViewModel? _vm;

    public MainWindow()
    {
        InitializeComponent();
        KeyBindings.Add(new KeyBinding
        {
            Gesture = new KeyGesture(Key.F, KeyModifiers.Control),
            Command = new CommunityToolkit.Mvvm.Input.RelayCommand(() => SearchBox.Focus()),
        });
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (_vm is not null)
        {
            _vm.ColumnsChanged -= RebuildColumns;
            _vm.PropertyChanged -= OnVmPropertyChanged;
        }
        _vm = DataContext as MainViewModel;
        if (_vm is null) return;
        _vm.Dialogs = this;
        _vm.ColumnsChanged += RebuildColumns;
        _vm.PropertyChanged += OnVmPropertyChanged;
        RebuildColumns();
        UpdateDetailsColumn();
    }

    private void OnVmPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.ShowDetails)) UpdateDetailsColumn();
    }

    private void UpdateDetailsColumn()
    {
        if (_vm is null) return;
        BodyGrid.ColumnDefinitions[4].Width = _vm.ShowDetails ? new GridLength(340) : new GridLength(0);
    }

    /// <summary>Recreates DataGrid columns for the selected table from its shared definition.</summary>
    private void RebuildColumns()
    {
        if (_vm?.SelectedTable is not { } table) return;
        var columns = _vm.VisibleColumns();
        Grid.Columns.Clear();
        var sample = table.AllRows.Take(300).ToList();
        foreach (var (index, column) in columns)
        {
            var maxChars = Math.Max(column.Header.Length, sample.Count == 0 ? 0 : sample.Max(r => r.Text[index].Length));
            var col = new DataGridTextColumn
            {
                Header = column.Header,
                Binding = new Binding($"Text[{index}]") { Mode = BindingMode.OneTime },
                CustomSortComparer = new GridRowComparer(index),
                SortMemberPath = $"Text[{index}]",
                CanUserSort = true,
                Width = new DataGridLength(Math.Clamp(maxChars * 7.4 + 30, 70, 420)),
            };
            if (column.Kind == CellKind.Number) col.CellStyleClasses.Add("num");
            Grid.Columns.Add(col);
        }
        Grid.FrozenColumnCount = columns.Count > 1 && table.Definition.FrozenColumns > 0 ? 1 : 0;
    }

    private void OnGridSelectionChanged(object? sender, SelectionChangedEventArgs e) =>
        _vm?.OnRowSelected(Grid.SelectedItem as GridRow);

    private void OnColumnsButtonClick(object? sender, RoutedEventArgs e)
    {
        if (_vm?.SelectedTable is not { } table) return;
        var hidden = _vm.HiddenColumns.GetValueOrDefault(table.Name) ?? [];
        var panel = new StackPanel { Spacing = 2, Margin = new Thickness(4) };
        var hideEmpty = new CheckBox { Content = "Hide columns with no data", IsChecked = _vm.HideEmptyColumns };
        hideEmpty.IsCheckedChanged += (_, _) => _vm.HideEmptyColumns = hideEmpty.IsChecked == true;
        panel.Children.Add(hideEmpty);
        panel.Children.Add(new Separator());
        for (var i = 0; i < table.Definition.Columns.Count; i++)
        {
            var header = table.Definition.Columns[i].Header;
            var hasData = i < table.ColumnHasData.Length && table.ColumnHasData[i];
            var cb = new CheckBox
            {
                Content = hasData ? header : $"{header}  (empty)",
                IsChecked = !hidden.Contains(header),
                Opacity = hasData ? 1 : 0.6,
            };
            cb.IsCheckedChanged += (_, _) => _vm.SetColumnHidden(header, cb.IsChecked != true);
            panel.Children.Add(cb);
        }
        var flyout = new Flyout
        {
            Content = new ScrollViewer { Content = panel, MaxHeight = 520, MinWidth = 260 },
            Placement = PlacementMode.BottomEdgeAlignedRight,
        };
        flyout.ShowAt(ColumnsButton);
    }

    private async void OnCopyRows(object? sender, RoutedEventArgs e)
    {
        if (_vm is null) return;
        var columns = _vm.VisibleColumns();
        var sb = new StringBuilder();
        sb.AppendLine(string.Join('\t', columns.Select(c => c.Column.Header)));
        foreach (var row in Grid.SelectedItems.OfType<GridRow>())
            sb.AppendLine(string.Join('\t', columns.Select(c => row.Text[c.Index])));
        await SetClipboardTextAsync(sb.ToString());
    }

    private async void OnCopyCell(object? sender, RoutedEventArgs e)
    {
        if (Grid.SelectedItem is not GridRow row || Grid.CurrentColumn is not DataGridTextColumn col) return;
        if (col.Binding is Binding { Path: { } path } && path.StartsWith("Text[") &&
            int.TryParse(path[5..^1], out var idx))
            await SetClipboardTextAsync(row.Text[idx]);
    }

    // ---------------- IDialogService ----------------

    public async Task<ConnectDialogResult?> ShowConnectDialogAsync(ConnectDialogViewModel vm) =>
        await new ConnectDialog { DataContext = vm }.ShowDialog<ConnectDialogResult?>(TopWindow());

    public async Task<bool> ShowGroupDialogAsync(GroupEditViewModel vm) =>
        await new GroupDialog { DataContext = vm }.ShowDialog<bool>(TopWindow());

    public Task ShowHostsWindowAsync(HostsViewModel vm) =>
        new HostsWindow { DataContext = vm }.ShowDialog(this);

    public async Task<string?> PickSaveFileAsync(string title, string suggestedName, string extension, string? startDirectory = null)
    {
        var options = new FilePickerSaveOptions
        {
            Title = title,
            SuggestedFileName = suggestedName,
            DefaultExtension = extension,
            ShowOverwritePrompt = true,
            FileTypeChoices = [new FilePickerFileType(extension.ToUpperInvariant()) { Patterns = [$"*.{extension}"] }],
        };
        if (startDirectory is not null && Directory.Exists(startDirectory))
            options.SuggestedStartLocation = await StorageProvider.TryGetFolderFromPathAsync(startDirectory);
        var file = await StorageProvider.SaveFilePickerAsync(options);
        return file?.TryGetLocalPath();
    }

    public async Task<string?> PickOpenFileAsync(string title, IReadOnlyList<string> extensions)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = false,
            FileTypeFilter = [new FilePickerFileType(string.Join("/", extensions).ToUpperInvariant()) { Patterns = extensions.Select(x => $"*.{x}").ToList() }],
        });
        return files.FirstOrDefault()?.TryGetLocalPath();
    }

    public Task<bool> ConfirmAsync(string title, string message) => MessageDialog.Show(TopWindow(), title, message, confirm: true);

    public Task ShowMessageAsync(string title, string message) => MessageDialog.Show(TopWindow(), title, message, confirm: false);

    public async Task SetClipboardTextAsync(string text)
    {
        if (Clipboard is { } clipboard) await clipboard.SetTextAsync(text);
    }

    /// <summary>The window dialogs should be owned by (a nested dialog may be open).</summary>
    private Window TopWindow() =>
        OwnedWindows.LastOrDefault(w => w.IsVisible) is { } owned
            ? owned.OwnedWindows.LastOrDefault(w => w.IsVisible) ?? owned
            : this;
}
