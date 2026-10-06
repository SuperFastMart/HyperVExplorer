using System.Text;
using System.Text.Json;
using HypervisorExplorer.Collectors.HyperV;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Tests.HyperV;

public class HyperVCollectorTests
{
    private sealed class FakeRunner(Func<PowerShellInvocation, Action<string>?, CancellationToken, Task<PowerShellResult>> run, bool supported = true)
        : IPowerShellRunner
    {
        public List<PowerShellInvocation> Calls { get; } = [];
        public bool IsSupported => supported;

        public Task<PowerShellResult> RunAsync(PowerShellInvocation invocation, Action<string>? onStandardErrorLine, CancellationToken cancellationToken)
        {
            Calls.Add(invocation);
            return run(invocation, onStandardErrorLine, cancellationToken);
        }

        public static FakeRunner Returning(string stdout, string stderr = "", int exitCode = 0) =>
            new((_, _, _) => Task.FromResult(new PowerShellResult(exitCode, stdout, stderr)));
    }

    private sealed class ListProgress : IProgress<string>
    {
        public List<string> Items { get; } = [];
        public void Report(string value) => Items.Add(value);
    }

    private static readonly TcpProbe PortOpen = (_, _, _, _) => Task.FromResult<string?>(null);

    private static ConnectionRequest Req(string address = "hv01.contoso.local", CredentialKind kind = CredentialKind.UsernamePassword,
        string? user = @"CONTOSO\admin", string? secret = "S3cr€t!", bool ssl = false) => new()
    {
        Platform = Platform.HyperV,
        Address = address,
        CredentialKind = kind,
        Username = kind == CredentialKind.CurrentUser ? null : user,
        Secret = kind == CredentialKind.CurrentUser ? null : secret,
        UseSsl = ssl,
    };

    private static string Failure(string stage, string error, string category = "") =>
        JsonSerializer.Serialize(new { ok = false, stage, error, category });

    [Fact]
    public async Task Success_ParsesEnvelope_ForwardsProgress_AndKeepsSecretsOffTheCommandLine()
    {
        var runner = new FakeRunner((inv, onErr, _) =>
        {
            onErr?.Invoke("PROGRESS: HV01: VM 1/2: SQL01");
            onErr?.Invoke("#< CLIXML");
            var stdout = "VERBOSE: noise\r\n" + HyperVFixtures.Envelope(HyperVFixtures.ClusterData) + "\r\n";
            return Task.FromResult(new PowerShellResult(0, stdout, ""));
        });
        var progress = new ListProgress();
        var snap = await new HyperVCollector(runner, PortOpen).CollectAsync(Req(), progress, CancellationToken.None);

        Assert.Equal(3, snap.VirtualMachines.Count);
        Assert.Contains("HV01: VM 1/2: SQL01", progress.Items);
        Assert.DoesNotContain(progress.Items, p => p.Contains("CLIXML"));

        var call = Assert.Single(runner.Calls);
        Assert.DoesNotContain("S3cr", call.Script);
        Assert.Equal(TimeSpan.FromMinutes(10), call.Timeout);
        // stdin is base64 (pure ASCII) UTF-8 JSON carrying the request and both scripts.
        Assert.All(call.StandardInput.Trim(), c => Assert.True(c < 128));
        using var req = JsonDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(call.StandardInput.Trim())));
        var root = req.RootElement;
        Assert.Equal("hv01.contoso.local", root.GetProperty("computer").GetString());
        Assert.Equal(@"CONTOSO\admin", root.GetProperty("user").GetString());
        Assert.Equal("S3cr€t!", root.GetProperty("password").GetString());
        Assert.True(root.GetProperty("expandCluster").GetBoolean());
        Assert.False(root.GetProperty("useSsl").GetBoolean());
        Assert.Contains("schemaVersion", root.GetProperty("script").GetString());
        Assert.Contains("Invoke-Command", root.GetProperty("wrapper").GetString());
    }

    [Fact]
    public async Task CurrentUser_SendsNoCredentials()
    {
        var runner = FakeRunner.Returning(HyperVFixtures.Envelope(HyperVFixtures.StandaloneHost));
        await new HyperVCollector(runner, PortOpen).CollectAsync(Req(kind: CredentialKind.CurrentUser), null, CancellationToken.None);
        using var req = JsonDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(runner.Calls[0].StandardInput.Trim())));
        Assert.False(req.RootElement.TryGetProperty("user", out _));
        Assert.False(req.RootElement.TryGetProperty("password", out _));
    }

    [Theory]
    [InlineData("connect", "Connecting to remote server hv01 failed with the following error message : Access is denied. For more information, see the about_Remote_Troubleshooting Help topic.", "AccessDenied,PSSessionStateBroken", CollectionFailure.AuthenticationFailed)]
    [InlineData("auth", "The user name or password is incorrect.", "", CollectionFailure.AuthenticationFailed)]
    [InlineData("connect", "WinRM cannot process the request. The following error with errorcode 0x80090311 occurred while using Kerberos authentication: There are currently no logon servers available to service the logon request.", "", CollectionFailure.AuthenticationFailed)]
    [InlineData("connect", "Connecting to remote server hv01 failed with the following error message : WinRM cannot complete the operation. Verify that the specified computer name is valid, that the computer is accessible over the network, and that a firewall exception for the WinRM service is enabled and allows access from this computer.", "WinRMOperationTimeout,PSSessionStateBroken", CollectionFailure.Unreachable)]
    [InlineData("connect", "The WinRM client cannot complete the operation within the time specified. 0x80338012", "", CollectionFailure.Unreachable)]
    [InlineData("collect", "HYPERV_MODULE_MISSING: The Hyper-V PowerShell module is not available on FS01. Is this a Hyper-V host?", "", CollectionFailure.Protocol)]
    [InlineData("collect", "HYPERV_ACCESS: Get-VMHost failed on HV01: You do not have the required permission to complete this task. Contact the administrator of the authorization policy for the computer 'HV01'.", "", CollectionFailure.PermissionDenied)]
    [InlineData("collect", "Something unexpected happened", "", CollectionFailure.Other)]
    [InlineData("connect", "Something unexpected happened", "", CollectionFailure.Unreachable)]
    public async Task FailureEnvelope_MapsToCollectionFailure(string stage, string error, string category, CollectionFailure expected)
    {
        var runner = FakeRunner.Returning(Failure(stage, error, category));
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req(), null, CancellationToken.None));
        Assert.Equal(expected, ex.Kind);
        if (expected != CollectionFailure.Other) Assert.False(string.IsNullOrWhiteSpace(ex.Hint));
        Assert.DoesNotContain("HYPERV_", ex.Message);
    }

    [Fact]
    public async Task TrustedHostsError_GivesExactCommand()
    {
        const string error = "Connecting to remote server 10.1.2.3 failed with the following error message : The WinRM client cannot process the request. " +
            "Default authentication may be used with an IP address under the following conditions: the transport is HTTPS or the destination is in the TrustedHosts list, " +
            "and explicit credentials are provided. Use winrm.cmd to configure TrustedHosts.";
        var runner = FakeRunner.Returning(Failure("connect", error, "CannotUseIPAddress,PSSessionStateBroken"));
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req("10.1.2.3"), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.AuthenticationFailed, ex.Kind);
        Assert.Contains(@"Set-Item WSMan:\localhost\Client\TrustedHosts -Value '10.1.2.3' -Concatenate -Force", ex.Hint);
    }

    [Fact]
    public async Task IpAddressWithCurrentUser_FailsBeforeSpawning()
    {
        var runner = FakeRunner.Returning("");
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req("10.1.2.3", CredentialKind.CurrentUser), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.AuthenticationFailed, ex.Kind);
        Assert.Contains("Kerberos", ex.Hint);
        Assert.Contains(@"Set-Item WSMan:\localhost\Client\TrustedHosts -Value '10.1.2.3' -Concatenate -Force", ex.Hint);
        Assert.Empty(runner.Calls);
    }

    [Theory]
    [InlineData(false, 5985)]
    [InlineData(true, 5986)]
    public async Task ClosedWinRmPort_IsUnreachableWithEnablePsRemotingHint(bool ssl, int expectedPort)
    {
        int? probedPort = null;
        TcpProbe closed = (_, port, _, _) => { probedPort = port; return Task.FromResult<string?>("connection refused"); };
        var runner = FakeRunner.Returning("");
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, closed).CollectAsync(Req(ssl: ssl), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.Unreachable, ex.Kind);
        Assert.Equal(expectedPort, probedPort);
        Assert.Contains("Enable-PSRemoting -Force", ex.Hint);
        Assert.Empty(runner.Calls);
    }

    [Fact]
    public async Task LocalHost_SkipsPortProbe()
    {
        var probed = false;
        TcpProbe probe = (_, _, _, _) => { probed = true; return Task.FromResult<string?>("closed"); };
        var runner = FakeRunner.Returning(HyperVFixtures.Envelope(HyperVFixtures.StandaloneHost));
        var snap = await new HyperVCollector(runner, probe).CollectAsync(Req("localhost", CredentialKind.CurrentUser), null, CancellationToken.None);
        Assert.False(probed);
        Assert.Single(snap.Hosts);
        Assert.Equal("localhost", snap.Source.Address);
    }

    [Fact]
    public async Task NonWindows_ThrowsWithImportHint()
    {
        var runner = new FakeRunner((_, _, _) => throw new InvalidOperationException("should not run"), supported: false);
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req(), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.Other, ex.Kind);
        Assert.Contains("requires Windows", ex.Message);
        Assert.Contains("Collect-HyperV.ps1", ex.Hint);
    }

    [Fact]
    public async Task NoEnvelope_ReportsStderrTail()
    {
        var runner = FakeRunner.Returning("", "PROGRESS: x\nThe term 'foo' is not recognized\n", exitCode: 1);
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req(), null, CancellationToken.None));
        Assert.Equal(CollectionFailure.Other, ex.Kind);
        Assert.Contains("not recognized", ex.Message);
        Assert.DoesNotContain("PROGRESS", ex.Message);
    }

    [Fact]
    public async Task Timeout_IsReported()
    {
        var runner = new FakeRunner((_, _, _) => Task.FromResult(new PowerShellResult(-1, "", "", TimedOut: true)));
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req(), null, CancellationToken.None));
        Assert.Contains("timed out", ex.Message);
    }

    [Fact]
    public async Task Cancellation_IsReportedAsCancelled()
    {
        using var cts = new CancellationTokenSource();
        var runner = new FakeRunner(async (_, _, ct) =>
        {
            await cts.CancelAsync();
            ct.ThrowIfCancellationRequested();
            return new PowerShellResult(0, "", "");
        });
        var ex = await Assert.ThrowsAsync<CollectionException>(() =>
            new HyperVCollector(runner, PortOpen).CollectAsync(Req(), null, cts.Token));
        Assert.Equal(CollectionFailure.Cancelled, ex.Kind);
    }

    [Theory]
    [InlineData("localhost", true)]
    [InlineData(".", true)]
    [InlineData("127.0.0.1", true)]
    [InlineData("::1", true)]
    [InlineData("hv01.contoso.local", false)]
    [InlineData("10.0.0.5", false)]
    public void IsLocalAddress(string address, bool expected) => Assert.Equal(expected, HyperVCollector.IsLocalAddress(address));

    [Fact]
    public void IsLocalAddress_MachineName() =>
        Assert.True(HyperVCollector.IsLocalAddress(Environment.MachineName.ToLowerInvariant()));
}
