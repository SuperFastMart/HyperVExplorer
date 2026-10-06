using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text.Json;
using HypervisorExplorer.Core;

namespace HypervisorExplorer.App.Services;

public sealed record UpdateInfo(string Version, string ReleaseUrl);

/// <summary>
/// Checks GitHub for a newer release and applies it by running the same one-line installer people use to install
/// (install.sh / install.ps1). Downloads made that way aren't marked as "from the internet", so the updated unsigned
/// app opens without Gatekeeper/SmartScreen prompts.
/// </summary>
public static class UpdateService
{
    public const string Repo = "SuperFastMart/HyperVisorExplorer";
    public const string InstallShUrl = "https://raw.githubusercontent.com/" + Repo + "/main/install.sh";
    public const string InstallPs1Url = "https://raw.githubusercontent.com/" + Repo + "/main/install.ps1";

    /// <summary>Returns the newer release, or null when up to date (or the check failed).</summary>
    public static async Task<UpdateInfo?> CheckAsync(CancellationToken ct)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("HypervisorExplorer", AppInfo.Version));
        http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        using var resp = await http.GetAsync($"https://api.github.com/repos/{Repo}/releases/latest", ct).ConfigureAwait(false);
        if (!resp.IsSuccessStatusCode) return null;
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var tag = doc.RootElement.TryGetProperty("tag_name", out var t) ? t.GetString() : null;
        var url = doc.RootElement.TryGetProperty("html_url", out var u) ? u.GetString() : null;
        if (string.IsNullOrEmpty(tag)) return null;
        return AppInfo.IsNewer(tag, AppInfo.Version) ? new UpdateInfo(tag.TrimStart('v'), url ?? "") : null;
    }

    /// <summary>
    /// Starts the installer detached, pointed at this copy's install location, waiting for this process to exit.
    /// The caller should then shut the app down; the installer replaces it and reopens it.
    /// </summary>
    public static void StartInstaller()
    {
        var pid = Environment.ProcessId;
        var exe = Environment.ProcessPath ?? throw new InvalidOperationException("Cannot determine the app's location.");
        if (OperatingSystem.IsWindows())
        {
            var dir = Path.GetDirectoryName(exe)!;
            var command = $"$env:HVE_WAIT_PID='{pid}'; $env:HVE_INSTALL_DIR='{dir.Replace("'", "''")}'; irm {InstallPs1Url} | iex";
            Process.Start(new ProcessStartInfo("powershell.exe")
            {
                ArgumentList = { "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command },
                UseShellExecute = false,
            });
        }
        else if (OperatingSystem.IsMacOS())
        {
            // exe = …/X.app/Contents/MacOS/<binary>: update the bundle in place, wherever it lives.
            var appDir = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(exe))))!;
            var log = Path.Combine(Path.GetTempPath(), "hypervisor-explorer-update.log");
            var script = $"(curl -fsSL '{InstallShUrl}' | HVE_WAIT_PID={pid} HVE_APP_DIR='{appDir.Replace("'", "'\\''")}' HVE_NO_CLI=1 bash) > '{log}' 2>&1 &";
            Process.Start(new ProcessStartInfo("/bin/bash") { ArgumentList = { "-c", script }, UseShellExecute = false });
        }
        else
        {
            throw new PlatformNotSupportedException("Automatic updates are available on Windows and macOS.");
        }
    }
}
