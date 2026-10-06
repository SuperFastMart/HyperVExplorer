using Avalonia.Controls;
using Avalonia.Interactivity;
using HypervisorExplorer.App.ViewModels;

namespace HypervisorExplorer.App.Views;

public partial class GroupDialog : Window
{
    public GroupDialog() => InitializeComponent();

    private void OnOk(object? sender, RoutedEventArgs e)
    {
        if (DataContext is GroupEditViewModel vm && vm.Validate()) Close(true);
    }

    private void OnCancel(object? sender, RoutedEventArgs e) => Close(false);
}
