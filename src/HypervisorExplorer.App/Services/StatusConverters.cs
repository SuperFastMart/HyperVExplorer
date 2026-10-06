using Avalonia.Data.Converters;

namespace HypervisorExplorer.App.Services;

public static class StatusConverters
{
    public static readonly IValueConverter IsOk =
        new FuncValueConverter<ConnectionStatus, bool>(s => s is ConnectionStatus.Connected or ConnectionStatus.Imported);

    public static readonly IValueConverter IsFailed =
        new FuncValueConverter<ConnectionStatus, bool>(s => s == ConnectionStatus.Failed);
}
