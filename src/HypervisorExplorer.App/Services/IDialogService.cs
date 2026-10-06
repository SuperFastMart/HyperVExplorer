using HypervisorExplorer.App.ViewModels;

namespace HypervisorExplorer.App.Services;

/// <summary>Window-level interactions the view models need, implemented by the main window.</summary>
public interface IDialogService
{
    Task<ConnectDialogResult?> ShowConnectDialogAsync(ConnectDialogViewModel vm);
    Task<bool> ShowGroupDialogAsync(GroupEditViewModel vm);
    Task ShowHostsWindowAsync(HostsViewModel vm);
    Task<string?> PickSaveFileAsync(string title, string suggestedName, string extension, string? startDirectory = null);
    Task<string?> PickOpenFileAsync(string title, IReadOnlyList<string> extensions);
    Task<bool> ConfirmAsync(string title, string message);
    Task ShowMessageAsync(string title, string message);
    Task SetClipboardTextAsync(string text);
}
