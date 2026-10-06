using System.Reflection;

namespace HypervisorExplorer.Core;

/// <summary>The running application's version, taken from the build (release builds stamp it from the git tag).</summary>
public static class AppInfo
{
    /// <summary>e.g. "3.0.2" (any "+commit" suffix the SDK adds is removed).</summary>
    public static string Version { get; } = Resolve();

    /// <summary>True for local/CI builds that weren't stamped from a release tag.</summary>
    public static bool IsDevelopmentBuild => Version.StartsWith("0.0.0", StringComparison.Ordinal);

    private static string Resolve()
    {
        var asm = Assembly.GetEntryAssembly() ?? typeof(AppInfo).Assembly;
        var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                   ?? asm.GetName().Version?.ToString(3)
                   ?? "0.0.0";
        var plus = info.IndexOf('+');
        return plus >= 0 ? info[..plus] : info;
    }

    /// <summary>Compares "3.0.10" style versions (an optional leading "v" and any "-suffix" are ignored).</summary>
    public static bool IsNewer(string candidate, string current)
    {
        static System.Version? Parse(string v)
        {
            v = v.Trim().TrimStart('v', 'V');
            var dash = v.IndexOf('-');
            if (dash >= 0) v = v[..dash];
            return System.Version.TryParse(v, out var parsed) ? parsed : null;
        }
        var a = Parse(candidate);
        var b = Parse(current);
        return a is not null && b is not null && a > b;
    }
}
