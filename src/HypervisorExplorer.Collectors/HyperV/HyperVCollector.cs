using System.Net.Sockets;
using System.Text.Json;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>Checks that a TCP port accepts connections. Returns null on success, else a short reason.</summary>
public delegate Task<string?> TcpProbe(string host, int port, TimeSpan timeout, CancellationToken cancellationToken);

/// <summary>
/// Hyper-V / Failover Cluster collector. Spawns PowerShell, which runs Collect-HyperV.ps1 on the host via
/// Invoke-Command (or locally) and returns one JSON document that <see cref="HyperVJsonMapper"/> maps.
/// </summary>
public sealed class HyperVCollector : IInventoryCollector
{
    private readonly IPowerShellRunner _runner;
    private readonly TcpProbe _probe;

    public HyperVCollector(IPowerShellRunner? runner = null, TcpProbe? tcpProbe = null)
    {
        _runner = runner ?? new ProcessPowerShellRunner();
        _probe = tcpProbe ?? DefaultTcpProbe;
    }

    public Platform Platform => Platform.HyperV;

    /// <summary>Overall limit for one collection (all cluster nodes included).</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromMinutes(10);

    /// <summary>Per-node limit for the parallel collection of other cluster nodes.</summary>
    public TimeSpan NodeTimeout { get; init; } = TimeSpan.FromMinutes(8);

    /// <summary>The text of Collect-HyperV.ps1, for "Save collection script..." (run by hand, then import).</summary>
    public static string GetCollectionScript() => HyperVScripts.CollectScript;

    public async Task<InventorySnapshot> CollectAsync(ConnectionRequest request, IProgress<string>? progress, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);
        var address = request.Address?.Trim() ?? "";
        if (address.Length == 0)
            throw new CollectionException(CollectionFailure.Other, "No host address was given.");

        if (!_runner.IsSupported)
            throw new CollectionException(CollectionFailure.Other,
                "Hyper-V collection requires Windows (it uses PowerShell remoting).", HyperVErrorMapper.NotWindowsHint);

        var local = IsLocalAddress(address);
        var useCurrentUser = request.CredentialKind == CredentialKind.CurrentUser || string.IsNullOrEmpty(request.Username);
        if (!local)
        {
            if (HyperVErrorMapper.IsIpAddress(address) && useCurrentUser)
                throw new CollectionException(CollectionFailure.AuthenticationFailed,
                    $"Cannot use current-user (Kerberos) authentication with the IP address {address}.",
                    HyperVErrorMapper.IpWithCurrentUserHint(address));

            var port = request.Port ?? (request.UseSsl ? 5986 : 5985);
            progress?.Report($"Checking WinRM port {port} on {address}...");
            string? probeError;
            try
            {
                probeError = await _probe(address, port, TimeSpan.FromSeconds(5), cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw new CollectionException(CollectionFailure.Cancelled, "Collection cancelled.");
            }
            if (probeError is not null)
                throw new CollectionException(CollectionFailure.Unreachable,
                    $"WinRM port (TCP {port}) is not reachable on '{address}': {probeError}",
                    request.UseSsl
                        ? HyperVErrorMapper.EnableRemotingHint + "\nFor HTTPS (5986) the host also needs an HTTPS listener: winrm quickconfig -transport:https"
                        : HyperVErrorMapper.EnableRemotingHint);
        }

        progress?.Report(local ? "Collecting the local Hyper-V host..." : $"Connecting to {address} over WinRM...");
        var invocation = new PowerShellInvocation(HyperVScripts.Bootstrap, BuildStandardInput(request, address, useCurrentUser, (int)NodeTimeout.TotalSeconds), Timeout);

        PowerShellResult result;
        try
        {
            result = await _runner.RunAsync(invocation, line =>
            {
                const string prefix = "PROGRESS:";
                if (line.StartsWith(prefix, StringComparison.Ordinal))
                    progress?.Report(line[prefix.Length..].Trim());
            }, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw new CollectionException(CollectionFailure.Cancelled, "Collection cancelled.");
        }
        catch (InvalidOperationException ex)
        {
            throw new CollectionException(CollectionFailure.Other, ex.Message,
                "Windows PowerShell (powershell.exe) or PowerShell 7 (pwsh) is required.", ex);
        }

        if (result.TimedOut)
            throw new CollectionException(CollectionFailure.Other,
                $"Hyper-V collection from '{address}' timed out after {Timeout.TotalMinutes:0} minutes.",
                "Very large hosts or unresponsive storage paths can stall Get-VHD. Try running Collect-HyperV.ps1 on the host and importing the JSON.");

        progress?.Report("Processing collected data...");
        return ParseEnvelope(result, request);
    }

    /// <summary>Extracts the envelope from PowerShell output and maps it (or the failure) — exposed for tests.</summary>
    public static InventorySnapshot ParseEnvelope(PowerShellResult result, ConnectionRequest request)
    {
        var envelope = FindEnvelope(result.StandardOutput);
        if (envelope is null)
        {
            var tail = string.Join(" ", result.StandardError.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(l => !l.StartsWith("PROGRESS:", StringComparison.Ordinal) && !l.StartsWith("#< CLIXML", StringComparison.Ordinal))
                .TakeLast(5));
            throw new CollectionException(CollectionFailure.Other,
                $"PowerShell exited (code {result.ExitCode}) without returning a result.{(tail.Length > 0 ? " " + tail : "")}");
        }

        using var doc = JsonDocument.Parse(envelope, new JsonDocumentOptions { MaxDepth = 256 });
        var root = doc.RootElement;
        var ok = root.TryGetProperty("ok", out var okEl) && okEl.ValueKind == JsonValueKind.True;
        if (!ok)
        {
            throw HyperVErrorMapper.FromEnvelope(
                GetString(root, "stage"), GetString(root, "error"), GetString(root, "category"), request);
        }
        if (!root.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Object)
            throw new CollectionException(CollectionFailure.Protocol, "The collection result contained no data.");
        return HyperVJsonMapper.Map(data, request);
    }

    private static string? GetString(JsonElement e, string name) =>
        e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    /// <summary>The last stdout line that is a JSON envelope.</summary>
    internal static string? FindEnvelope(string stdout)
    {
        string? found = null;
        foreach (var raw in stdout.Split('\n'))
        {
            var line = raw.Trim().TrimStart('﻿');
            if (line.StartsWith("{\"ok\"", StringComparison.Ordinal)) found = line;
        }
        return found;
    }

    private static string BuildStandardInput(ConnectionRequest request, string address, bool useCurrentUser, int nodeTimeoutSeconds)
    {
        var buffer = new MemoryStream();
        using (var w = new Utf8JsonWriter(buffer))
        {
            w.WriteStartObject();
            w.WriteString("computer", address);
            if (request.Port is { } port) w.WriteNumber("port", port);
            else w.WriteNull("port");
            w.WriteBoolean("useSsl", request.UseSsl);
            w.WriteBoolean("skipCertificateCheck", request.IgnoreCertificateErrors);
            if (!useCurrentUser)
            {
                w.WriteString("user", request.Username);
                w.WriteString("password", request.Secret ?? "");
            }
            w.WriteBoolean("expandCluster", request.ExpandCluster);
            w.WriteNumber("nodeTimeoutSeconds", nodeTimeoutSeconds);
            w.WriteString("script", HyperVScripts.CollectScript);
            w.WriteString("wrapper", HyperVScripts.WrapperScript);
            w.WriteEndObject();
        }
        // Base64 keeps stdin pure ASCII, so console code pages cannot mangle non-ASCII passwords.
        return Convert.ToBase64String(buffer.GetBuffer(), 0, (int)buffer.Length) + "\n";
    }

    /// <summary>localhost, ".", loopback, or this machine's own name: collected without WinRM.</summary>
    public static bool IsLocalAddress(string address)
    {
        var a = address.Trim();
        if (a == ".") return true;
        a = a.TrimEnd('.');
        if ( a.Equals("localhost", StringComparison.OrdinalIgnoreCase) || a is "127.0.0.1" or "::1" or "[::1]")
            return true;
        var machine = Environment.MachineName;
        return a.Equals(machine, StringComparison.OrdinalIgnoreCase)
            || a.StartsWith(machine + ".", StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<string?> DefaultTcpProbe(string host, int port, TimeSpan timeout, CancellationToken cancellationToken)
    {
        using var client = new TcpClient();
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(timeout);
        try
        {
            await client.ConnectAsync(host, port, cts.Token).ConfigureAwait(false);
            return null;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            return $"no response within {timeout.TotalSeconds:0} s";
        }
        catch (SocketException ex) when (ex.SocketErrorCode is SocketError.HostNotFound or SocketError.NoData or SocketError.TryAgain)
        {
            return "the host name could not be resolved (check the name / DNS)";
        }
        catch (SocketException ex) when (ex.SocketErrorCode == SocketError.ConnectionRefused)
        {
            return "connection refused";
        }
        catch (SocketException ex)
        {
            return ex.Message;
        }
    }
}
