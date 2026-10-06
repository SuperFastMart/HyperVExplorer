using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using HypervisorExplorer.App.ViewModels;

namespace HypervisorExplorer.App.Views;

public partial class HostsWindow : Window
{
    public HostsWindow() => InitializeComponent();

    private void OnHostDoubleTapped(object? sender, TappedEventArgs e)
    {
        if (DataContext is HostsViewModel vm && vm.EditHostCommand.CanExecute(null))
            vm.EditHostCommand.Execute(null);
    }

    private void OnClose(object? sender, RoutedEventArgs e) => Close();
}
