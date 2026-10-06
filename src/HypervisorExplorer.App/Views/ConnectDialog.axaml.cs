using Avalonia.Controls;
using Avalonia.Interactivity;
using HypervisorExplorer.App.ViewModels;

namespace HypervisorExplorer.App.Views;

public partial class ConnectDialog : Window
{
    public ConnectDialog()
    {
        InitializeComponent();
        Opened += (_, _) => AddressBox.Focus();
    }

    private void OnOk(object? sender, RoutedEventArgs e)
    {
        if (DataContext is ConnectDialogViewModel vm && vm.TryBuild() is { } result)
            Close(result);
    }

    private void OnCancel(object? sender, RoutedEventArgs e) => Close(null);
}
