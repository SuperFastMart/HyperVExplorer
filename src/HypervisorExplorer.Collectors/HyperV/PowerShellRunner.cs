using System.ComponentModel;
using System.Diagnostics;
using System.Text;

namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>One PowerShell invocation: a script passed as -EncodedCommand plus text written to stdin.</summary>
/// <param name="Script">Script text; passed UTF-16LE base64 encoded on the command line. Must not contain secrets.</param>
/// <param name="StandardInput">Written to stdin, then stdin is closed. Carries the request (incl. credentials).</param>
/// <param name="Timeout">Kill the process tree after this long.</param>
public sealed record PowerShellInvocation(string Script, string StandardInput, TimeSpan Timeout);

/// <summary>Raw outcome of a PowerShell process.</summary>
public sealed record PowerShellResult(int ExitCode, string StandardOutput, string StandardError, bool TimedOut = false);

/// <summary>Runs PowerShell. Injectable so tests can feed canned envelopes.</summary>
public interface IPowerShellRunner
{
    /// <summary>False when PowerShell remoting cannot work here (non-Windows OS).</summary>
    bool IsSupported { get; }

    /// <summary>
    /// Runs the invocation. Each stderr line is passed to <paramref name="onStandardErrorLine"/> as it arrives.
    /// Cancellation kills the process tree and throws <see cref="OperationCanceledException"/>.
    /// </summary>
    Task<PowerShellResult> RunAsync(PowerShellInvocation invocation, Action<string>? onStandardErrorLine, CancellationToken cancellationToken);
}

/// <summary>
/// Spawns Windows PowerShell (<c>powershell.exe</c>, present on every Windows install) or, failing that, <c>pwsh</c>,
/// with <c>-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand</c>.
/// </summary>
public sealed class ProcessPowerShellRunner : IPowerShellRunner
{
    public bool IsSupported => OperatingSystem.IsWindows();

    /// <summary>Full path of the PowerShell executable to use, or null if none was found.</summary>
    public static string? FindPowerShell()
    {
        if (OperatingSystem.IsWindows())
        {
            var sysRoot = Environment.GetEnvironmentVariable("SystemRoot") ?? @"C:\Windows";
            var winPs = Path.Combine(sysRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            if (File.Exists(winPs)) return winPs;
        }
        var exe = OperatingSystem.IsWindows() ? "pwsh.exe" : "pwsh";
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                var candidate = Path.Combine(dir.Trim(), exe);
                if (File.Exists(candidate)) return candidate;
            }
            catch (ArgumentException)
            {
                // Malformed PATH entry.
            }
        }
        return null;
    }

    public async Task<PowerShellResult> RunAsync(PowerShellInvocation invocation, Action<string>? onStandardErrorLine, CancellationToken cancellationToken)
    {
        var exe = FindPowerShell() ?? throw new InvalidOperationException("Neither powershell.exe nor pwsh was found.");
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(invocation.Script));
        var utf8 = new UTF8Encoding(false);
        var psi = new ProcessStartInfo(exe)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardInputEncoding = utf8,
            StandardOutputEncoding = utf8,
            StandardErrorEncoding = utf8,
        };
        foreach (var arg in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-OutputFormat", "Text", "-EncodedCommand", encoded })
            psi.ArgumentList.Add(arg);

        using var process = new Process { StartInfo = psi, EnableRaisingEvents = true };
        var stderr = new StringBuilder();
        var stderrDone = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null)
            {
                stderrDone.TrySetResult();
                return;
            }
            lock (stderr) stderr.AppendLine(e.Data);
            try { onStandardErrorLine?.Invoke(e.Data); }
            catch { /* progress sinks must not break collection */ }
        };

        try
        {
            process.Start();
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException($"Could not start {exe}: {ex.Message}", ex);
        }
        process.BeginErrorReadLine();

        using var timeoutCts = new CancellationTokenSource(invocation.Timeout);
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);
        try
        {
            var stdoutTask = process.StandardOutput.ReadToEndAsync(linked.Token);
            await process.StandardInput.WriteAsync(invocation.StandardInput.AsMemory(), linked.Token).ConfigureAwait(false);
            await process.StandardInput.FlushAsync(linked.Token).ConfigureAwait(false);
            process.StandardInput.Close();

            var stdout = await stdoutTask.ConfigureAwait(false);
            await process.WaitForExitAsync(linked.Token).ConfigureAwait(false);
            try
            {
                await stderrDone.Task.WaitAsync(TimeSpan.FromSeconds(5), CancellationToken.None).ConfigureAwait(false);
            }
            catch (TimeoutException)
            {
                // A grandchild still holds stderr open; keep what arrived.
            }
            string err;
            lock (stderr) err = stderr.ToString();
            return new PowerShellResult(process.ExitCode, stdout, err);
        }
        catch (Exception ex) when (ex is OperationCanceledException or IOException)
        {
            Kill(process);
            if (cancellationToken.IsCancellationRequested) throw new OperationCanceledException(cancellationToken);
            if (timeoutCts.IsCancellationRequested)
            {
                string err;
                lock (stderr) err = stderr.ToString();
                return new PowerShellResult(-1, "", err, TimedOut: true);
            }
            throw;
        }
    }

    private static void Kill(Process process)
    {
        try
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        }
        catch (InvalidOperationException)
        {
            // Already exited.
        }
        catch (Win32Exception)
        {
            // Access denied / exiting.
        }
    }
}
