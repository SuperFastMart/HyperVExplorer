using System.IO.Compression;
using System.Text;
using System.Text.Json;
using HypervisorExplorer.Core.Collection;

namespace HypervisorExplorer.Collectors.HyperV.WinRm;

/// <summary>Opens a WinRM transport to a host; injectable for tests.</summary>
public delegate WinRmTransport WinRmTransportFactory(string host, int port, bool useHttps, string user, string password, bool ignoreCertificateErrors);

/// <summary>
/// Hyper-V collection without local PowerShell (macOS/Linux): talks WinRM directly, runs Collect-HyperV.ps1
/// on each host through a remote shell, and fans out to Failover Cluster nodes from this machine.
/// Produces the same "multi" document the Windows launcher does, so <see cref="HyperVJsonMapper"/> maps it unchanged.
/// </summary>
public sealed class NativeHyperVCollection
{
    private const string ResultMarker = "<<<HVE:RESULT>>>";
    private const string ErrorMarker = "<<<HVE:ERROR>>>";

    /// <summary>
    /// Runs on the host via <c>powershell.exe -EncodedCommand</c>. Reads a base64 JSON request from stdin
    /// (credentials never appear here), then either discovers cluster nodes or runs the collection script.
    /// The result is gzip+base64 behind a marker so console code pages and line wrapping can't corrupt it.
    /// Windows PowerShell 5.1 compatible.
    /// </summary>
    internal const string RemoteBootstrap = """
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        function Out-Marked([string]$Tag, [string]$Text) {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
            $ms = New-Object System.IO.MemoryStream
            $gz = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
            $gz.Write($bytes, 0, $bytes.Length)
            $gz.Close()
            [Console]::Out.WriteLine($Tag + [Convert]::ToBase64String($ms.ToArray()))
            [Console]::Out.Flush()
        }
        try {
            $raw = [string]($input | Out-String)
            if (-not $raw.Trim()) { $raw = [Console]::In.ReadToEnd() }
            $b64 = $raw -replace '[^A-Za-z0-9+/=]', ''
            $req = ConvertFrom-Json -InputObject ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)))
            if ($req.mode -eq 'discover') {
                $me = [string]$env:COMPUTERNAME
                $fqdn = $me
                try { $fqdn = [System.Net.Dns]::GetHostEntry($me).HostName } catch { }
                $lines = @('PRIMARY|' + $me + '|' + $fqdn)
                if (Get-Command -Name Get-ClusterNode -ErrorAction SilentlyContinue) {
                    try {
                        foreach ($n in @(Get-ClusterNode -ErrorAction Stop)) {
                            $name = [string]$n.Name
                            $nf = $name
                            $ips = @()
                            try { $nf = [System.Net.Dns]::GetHostEntry($name).HostName } catch { }
                            try { $ips = @([System.Net.Dns]::GetHostAddresses($name) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.IPAddressToString }) } catch { }
                            $lines += ('NODE|' + $name + '|' + [string]$n.State + '|' + $nf + '|' + ($ips -join ','))
                        }
                    } catch { }
                }
                Out-Marked '<<<HVE:RESULT>>>' ($lines -join "`n")
            } else {
                $sb = [scriptblock]::Create([string]$req.script)
                $out = New-Object 'System.Collections.Generic.List[string]'
                & $sb $req.clusterPrimary 4>&1 | ForEach-Object {
                    if ($_ -is [System.Management.Automation.VerboseRecord]) {
                        if ($_.Message -like 'PROGRESS:*') { [Console]::Error.WriteLine($_.Message); [Console]::Error.Flush() }
                    } elseif ($_ -is [System.Management.Automation.WarningRecord] -or $_ -is [System.Management.Automation.DebugRecord]) {
                    } elseif ($null -ne $_) { $out.Add([string]$_) }
                }
                $json = $null
                foreach ($o in $out) { if ($o -and $o.TrimStart().StartsWith('{')) { $json = $o.Trim() } }
                if (-not $json) { throw 'The collection script returned no data.' }
                Out-Marked '<<<HVE:RESULT>>>' $json
            }
        } catch {
            Out-Marked '<<<HVE:ERROR>>>' ([string]$_.Exception.Message + '||' + [string]$_.FullyQualifiedErrorId)
        }
        """;

    private readonly WinRmTransportFactory _factory;

    public NativeHyperVCollection(WinRmTransportFactory? factory = null)
    {
        _factory = factory ?? ((host, port, https, user, password, ignore) =>
            new WinRmTransport(host, port, https, user, password, ignore, TimeSpan.FromMinutes(2)));
    }

    public int MaxParallelNodes { get; init; } = 8;

    /// <summary>Collects the host (and its cluster peers) and returns the "multi" JSON document.</summary>
    public async Task<string> CollectAsync(ConnectionRequest request, string address, int port, IProgress<string>? progress,
        TimeSpan nodeTimeout, CancellationToken ct)
    {
        var warnings = new List<string>();
        using var primary = Open(request, address, port);

        progress?.Report($"Connecting to {address} over WinRM ({(request.UseSsl ? "HTTPS" : "HTTP, NTLM encrypted")})...");
        var discovery = await RunAsync(primary, new { mode = "discover" }, null, ct).ConfigureAwait(false);
        var (primaryName, nodes) = ParseDiscovery(discovery);
        progress?.Report($"Connected to {primaryName}.");

        var others = new List<DiscoveredNode>();
        if (request.ExpandCluster)
        {
            foreach (var node in nodes)
            {
                if (node.Name.Equals(primaryName, StringComparison.OrdinalIgnoreCase)) continue;
                if (node.State is "Up" or "Paused") others.Add(node);
                else warnings.Add($"Cluster node {node.Name} is {node.State}; not collected.");
            }
        }
        if (others.Count > 0)
            progress?.Report($"Failover cluster detected; collecting {others.Count} other node(s) in parallel: {string.Join(", ", others.Select(o => o.Name))}");

        var request2 = new { mode = "collect", script = HyperVScripts.CollectScript, clusterPrimary = primaryName };
        using var allCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        var primaryTask = RunAsync(primary, request2, line => ReportProgress(progress, line), allCts.Token);

        using var gate = new SemaphoreSlim(MaxParallelNodes);
        var nodeTasks = others.Select(async node =>
        {
            try
            {
                await gate.WaitAsync(allCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                return null;
            }
            try
            {
                using var nodeCts = CancellationTokenSource.CreateLinkedTokenSource(allCts.Token);
                nodeCts.CancelAfter(nodeTimeout);
                return await CollectNodeAsync(request, node, port, request2, progress, nodeCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!allCts.IsCancellationRequested)
            {
                lock (warnings) warnings.Add($"Cluster node {node.Name} timed out after {nodeTimeout.TotalSeconds:0} s; not collected.");
                return null;
            }
            catch (Exception ex) when (ex is WinRmException or CollectionException)
            {
                lock (warnings) warnings.Add($"Cluster node collection failed: [{node.Name}] {ex.Message}");
                return null;
            }
            finally
            {
                gate.Release();
            }
        }).ToList();

        string primaryJson;
        try
        {
            primaryJson = await primaryTask.ConfigureAwait(false);
        }
        catch
        {
            // The primary host failed: stop the cluster peers rather than leave them running unobserved.
            await allCts.CancelAsync().ConfigureAwait(false);
            try { await Task.WhenAll(nodeTasks).ConfigureAwait(false); } catch (OperationCanceledException) { }
            throw;
        }
        var nodeJsons = new List<string> { primaryJson };
        foreach (var t in nodeTasks)
        {
            if (await t.ConfigureAwait(false) is { } json) nodeJsons.Add(json);
        }

        var sb = new StringBuilder();
        sb.Append("{\"schemaVersion\":1,\"kind\":\"multi\",\"requestedComputer\":").Append(JsonSerializer.Serialize(address))
          .Append(",\"primary\":").Append(JsonSerializer.Serialize(primaryName))
          .Append(",\"nodes\":[").Append(string.Join(",", nodeJsons))
          .Append("],\"warnings\":").Append(JsonSerializer.Serialize(warnings)).Append('}');
        return sb.ToString();
    }

    private async Task<string?> CollectNodeAsync(ConnectionRequest request, DiscoveredNode node, int port, object collectRequest,
        IProgress<string>? progress, CancellationToken ct)
    {
        // The cluster's own names may not resolve from this machine: try FQDN, short name, then each IPv4 address.
        Exception? last = null;
        foreach (var candidate in node.Candidates)
        {
            try
            {
                using var transport = Open(request, candidate, port);
                return await RunAsync(transport, collectRequest, line => ReportProgress(progress, line), ct).ConfigureAwait(false);
            }
            catch (WinRmException ex) when (ex.Kind == WinRmErrorKind.Unreachable)
            {
                last = ex;
            }
        }
        throw last ?? new WinRmException(WinRmErrorKind.Unreachable, $"No reachable address for cluster node {node.Name}.");
    }

    private WinRmTransport Open(ConnectionRequest request, string host, int port) =>
        _factory(host, port, request.UseSsl, request.Username ?? "", request.Secret ?? "", request.IgnoreCertificateErrors);

    private static void ReportProgress(IProgress<string>? progress, string line)
    {
        const string prefix = "PROGRESS:";
        if (line.StartsWith(prefix, StringComparison.Ordinal)) progress?.Report(line[prefix.Length..].Trim());
    }

    /// <summary>Runs the bootstrap with a request on stdin and returns the decoded result payload.</summary>
    private static async Task<string> RunAsync(WinRmTransport transport, object payload, Action<string>? onStderr, CancellationToken ct)
    {
        var json = JsonSerializer.Serialize(payload);
        var stdin = Encoding.ASCII.GetBytes(Convert.ToBase64String(Encoding.UTF8.GetBytes(json)) + "\r\n");
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(RemoteBootstrap));
        var shell = new WinRmShell(transport);
        var result = await shell.RunAsync("powershell.exe",
            "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand " + encoded,
            stdin, onStderr, ct).ConfigureAwait(false);
        return DecodeResult(result);
    }

    /// <summary>Finds the marked result (or error) in stdout. Exposed for tests.</summary>
    internal static string DecodeResult(RemoteCommandResult result)
    {
        foreach (var raw in result.StandardOutput.Split('\n'))
        {
            var line = raw.Trim();
            if (line.StartsWith(ResultMarker, StringComparison.Ordinal)) return Gunzip(line[ResultMarker.Length..]);
            if (line.StartsWith(ErrorMarker, StringComparison.Ordinal))
            {
                var parts = Gunzip(line[ErrorMarker.Length..]).Split("||", 2);
                throw new RemoteScriptException(parts[0], parts.Length > 1 ? parts[1] : "");
            }
        }
        var err = string.Join(" ", result.StandardError.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(l => !l.StartsWith("PROGRESS:", StringComparison.Ordinal) && !l.StartsWith("#< CLIXML", StringComparison.Ordinal)).TakeLast(5));
        throw new WinRmException(WinRmErrorKind.Protocol,
            $"PowerShell on the host exited (code {result.ExitCode}) without a result.{(err.Length > 0 ? " " + err : "")}");
    }

    internal static string Gunzip(string b64)
    {
        using var input = new MemoryStream(Convert.FromBase64String(b64));
        using var gz = new GZipStream(input, CompressionMode.Decompress);
        using var reader = new StreamReader(gz, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    internal sealed record DiscoveredNode(string Name, string State, string Fqdn, IReadOnlyList<string> Addresses)
    {
        public IEnumerable<string> Candidates =>
            new[] { Fqdn, Name }.Concat(Addresses).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase);
    }

    internal static (string Primary, List<DiscoveredNode> Nodes) ParseDiscovery(string text)
    {
        var primary = "";
        var nodes = new List<DiscoveredNode>();
        foreach (var line in text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var p = line.Split('|');
            if (p[0] == "PRIMARY" && p.Length >= 2) primary = p[1];
            else if (p[0] == "NODE" && p.Length >= 3)
                nodes.Add(new DiscoveredNode(p[1], p[2], p.Length > 3 ? p[3] : p[1],
                    p.Length > 4 ? p[4].Split(',', StringSplitOptions.RemoveEmptyEntries) : []));
        }
        if (primary.Length == 0) throw new WinRmException(WinRmErrorKind.Protocol, "The host did not report its computer name.");
        return (primary, nodes);
    }
}

/// <summary>A PowerShell error raised by the collection script on the host.</summary>
public sealed class RemoteScriptException(string message, string category) : Exception(message)
{
    public string Category { get; } = category;
}
