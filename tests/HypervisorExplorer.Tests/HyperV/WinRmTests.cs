using System.IO.Compression;
using System.Text;
using HypervisorExplorer.Collectors.HyperV.WinRm;

namespace HypervisorExplorer.Tests.HyperV;

/// <summary>
/// Unit tests for the built-in WinRM client. The full NTLM + sealing + WinRS pipeline is exercised against an
/// independent implementation (pyspnego) by tools/FakeWinRm/fake_winrm.py.
/// </summary>
public class WinRmTests
{
    [Fact]
    public void Frame_and_unframe_round_trip_the_signature_and_payload()
    {
        var wrapped = Enumerable.Range(0, 16 + 300).Select(i => (byte)i).ToArray();
        var framed = WinRmTransport.Frame(wrapped, originalLength: 300);
        var text = Encoding.ASCII.GetString(framed);

        Assert.StartsWith("--Encrypted Boundary\r\n\tContent-Type: application/HTTP-SPNEGO-session-encrypted\r\n", text);
        Assert.Contains("\tOriginalContent: type=application/soap+xml;charset=UTF-8;Length=300\r\n", text);
        Assert.EndsWith("--Encrypted Boundary--\r\n", text);

        var back = WinRmTransport.Unframe(framed, out var expected);
        Assert.Equal(300, expected);
        Assert.Equal(wrapped, back);
    }

    [Theory]
    [InlineData(@"CONTOSO\admin", "admin", "CONTOSO")]
    [InlineData("admin@contoso.com", "admin@contoso.com", "")]
    [InlineData("administrator", "administrator", "")]
    public void Credentials_are_split_into_domain_and_user(string input, string user, string domain)
    {
        var c = WinRmTransport.ParseCredential(input, "pw");
        Assert.Equal(user, c.UserName);
        Assert.Equal(domain, c.Domain);
    }

    private static string GzB64(string s)
    {
        var ms = new MemoryStream();
        using (var gz = new GZipStream(ms, CompressionLevel.Optimal, leaveOpen: true))
            gz.Write(Encoding.UTF8.GetBytes(s));
        return Convert.ToBase64String(ms.ToArray());
    }

    [Fact]
    public void Decodes_marked_result_amid_other_output()
    {
        var stdout = "noise\r\n<<<HVE:RESULT>>>" + GzB64("{\"host\":{}}") + "\r\nmore noise\r\n";
        Assert.Equal("{\"host\":{}}", NativeHyperVCollection.DecodeResult(new RemoteCommandResult(0, stdout, "")));
    }

    [Fact]
    public void Marked_error_becomes_remote_script_exception()
    {
        var stdout = "<<<HVE:ERROR>>>" + GzB64("The term 'Get-VM' is not recognized||CommandNotFoundException") + "\r\n";
        var ex = Assert.Throws<RemoteScriptException>(() => NativeHyperVCollection.DecodeResult(new RemoteCommandResult(0, stdout, "")));
        Assert.Contains("Get-VM", ex.Message);
        Assert.Equal("CommandNotFoundException", ex.Category);
    }

    [Fact]
    public void Missing_result_reports_exit_code_and_stderr()
    {
        var ex = Assert.Throws<WinRmException>(() => NativeHyperVCollection.DecodeResult(
            new RemoteCommandResult(1, "", "PROGRESS: x\r\npowershell.exe : execution blocked\r\n")));
        Assert.Contains("code 1", ex.Message);
        Assert.Contains("execution blocked", ex.Message);
        Assert.DoesNotContain("PROGRESS", ex.Message);
    }

    [Fact]
    public void Discovery_lists_primary_and_cluster_node_addresses()
    {
        var (primary, nodes) = NativeHyperVCollection.ParseDiscovery(
            "PRIMARY|HV01|hv01.contoso.local\nNODE|HV01|Up|hv01.contoso.local|10.0.0.1\nNODE|HV02|Paused|hv02.contoso.local|10.0.0.2,10.1.0.2\nNODE|HV03|Down|hv03.contoso.local|");
        Assert.Equal("HV01", primary);
        Assert.Equal(3, nodes.Count);
        Assert.Equal(["hv02.contoso.local", "HV02", "10.0.0.2", "10.1.0.2"], nodes[1].Candidates);
        Assert.Empty(nodes[2].Addresses);
    }

    [Fact]
    public void Timeout_fault_is_recognised_and_access_denied_is_classified()
    {
        const string timedOut = """
            <s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:w="http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd"><s:Body><s:Fault>
            <s:Code><s:Value>s:Receiver</s:Value><s:Subcode><s:Value>w:TimedOut</s:Value></s:Subcode></s:Code>
            <s:Reason><s:Text xml:lang="en-US">The WS-Management service cannot complete the operation.</s:Text></s:Reason>
            <s:Detail><f:WSManFault xmlns:f="http://schemas.microsoft.com/wbem/wsman/1/wsmanfault" Code="2150858793" Machine="hv01"><f:Message>timed out</f:Message></f:WSManFault></s:Detail>
            </s:Fault></s:Body></s:Envelope>
            """;
        Assert.True(WsManFault.IsTimeout(WsManFault.Parse(timedOut)));

        var denied = WsManFault.Parse(timedOut.Replace("2150858793", "5").Replace("w:TimedOut", "w:AccessDenied").Replace("timed out", "Access is denied."));
        Assert.Equal(WinRmErrorKind.AccessDenied, denied.Kind);
        Assert.False(WsManFault.IsTimeout(denied));
    }

    [Fact]
    public void Remote_bootstrap_fits_command_line_and_is_windows_powershell_compatible()
    {
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(NativeHyperVCollection.RemoteBootstrap));
        Assert.True(encoded.Length < 30000, $"EncodedCommand is {encoded.Length} chars");
        foreach (var token in new[] { "??", "?.", "-AsHashtable", "-Parallel" })
            Assert.DoesNotContain(token, NativeHyperVCollection.RemoteBootstrap);
    }
}
