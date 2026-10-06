using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace HypervisorExplorer.App.Views;

/// <summary>Minimal message / confirmation box.</summary>
public static class MessageDialog
{
    public static async Task<bool> Show(Window owner, string title, string message, bool confirm)
    {
        var result = false;
        var window = new Window
        {
            Title = title,
            Width = 460,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false,
        };
        var ok = new Button { Content = confirm ? "OK" : "Close", IsDefault = true, Classes = { "primary" } };
        ok.Click += (_, _) => { result = true; window.Close(); };
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, HorizontalAlignment = HorizontalAlignment.Right };
        if (confirm)
        {
            var cancel = new Button { Content = "Cancel", IsCancel = true };
            cancel.Click += (_, _) => window.Close();
            buttons.Children.Add(cancel);
        }
        buttons.Children.Add(ok);
        window.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 14,
            Children =
            {
                new TextBlock { Text = title, FontSize = 16, FontWeight = FontWeight.SemiBold },
                new SelectableTextBlock { Text = message, TextWrapping = TextWrapping.Wrap },
                buttons,
            },
        };
        await window.ShowDialog(owner);
        return result;
    }
}
