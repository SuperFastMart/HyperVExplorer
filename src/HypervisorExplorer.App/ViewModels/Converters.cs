using Avalonia.Data.Converters;

namespace HypervisorExplorer.App.ViewModels;

public static class Converters
{
    public static readonly IValueConverter GroupSecretWatermark =
        new FuncValueConverter<bool, string>(inherits => inherits ? "leave blank to use the group's credentials" : "");
}
