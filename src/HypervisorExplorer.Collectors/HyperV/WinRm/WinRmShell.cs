using System.Security;
using System.Text;
using System.Xml.Linq;

namespace HypervisorExplorer.Collectors.HyperV.WinRm;

/// <summary>Parses WS-Management SOAP faults into <see cref="WinRmException"/>.</summary>
internal static class WsManFault
{
    /// <summary>WS-Man OperationTimeout expired with no output yet; the client should simply poll again.</summary>
    public const string TimedOutCode = "2150858793";

    public static WinRmException Parse(string soap)
    {
        string message;
        string code = "";
        try
        {
            var doc = XDocument.Parse(soap);
            var reason = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Text")?.Value?.Trim();
            var detail = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Message")?.Value?.Trim();
            code = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "WSManFault")?.Attribute("Code")?.Value ?? "";
            var sub = doc.Descendants().Where(e => e.Name.LocalName == "Value").Select(e => e.Value).LastOrDefault() ?? "";
            if (sub.EndsWith(":TimedOut", StringComparison.Ordinal)) code = TimedOutCode;
            message = !string.IsNullOrEmpty(detail) ? detail : reason ?? "Unknown WS-Management fault.";
        }
        catch (System.Xml.XmlException)
        {
            message = soap.Length > 300 ? soap[..300] : soap;
        }

        var kind = message.Contains("Access is denied", StringComparison.OrdinalIgnoreCase) || code == "5"
            ? WinRmErrorKind.AccessDenied
            : WinRmErrorKind.Protocol;
        return new WinRmException(kind, message) { Data = { ["WsManCode"] = code } };
    }

    public static bool IsTimeout(WinRmException ex) => (ex.Data["WsManCode"] as string) == TimedOutCode;
}

/// <summary>Result of running a command in a remote shell.</summary>
public sealed record RemoteCommandResult(int ExitCode, string StandardOutput, string StandardError);

/// <summary>
/// Windows Remote Shell (WinRS) over WS-Management: creates a cmd shell, runs one command with stdin,
/// streams stdout/stderr, and cleans up. This is the same mechanism as <c>winrs.exe</c>, pywinrm and Ansible.
/// </summary>
public sealed class WinRmShell(WinRmTransport transport)
{
    private const string ShellUri = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/cmd";
    private const string NsShell = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell";
    private const int MaxEnvelopeBytes = 153600; // 150 KB: accepted by every WinRM version since 2008
    private const int StdinChunkBytes = 64 * 1024;

    /// <summary>Runs <paramref name="command"/> with <paramref name="arguments"/>, feeding <paramref name="stdin"/>.</summary>
    public async Task<RemoteCommandResult> RunAsync(string command, string arguments, byte[]? stdin,
        Action<string>? onStderrLine, CancellationToken ct)
    {
        var shellId = await CreateShellAsync(ct).ConfigureAwait(false);
        try
        {
            var commandId = await StartCommandAsync(shellId, command, arguments, ct).ConfigureAwait(false);
            if (stdin is { Length: > 0 })
            {
                for (var offset = 0; offset < stdin.Length; offset += StdinChunkBytes)
                {
                    var count = Math.Min(StdinChunkBytes, stdin.Length - offset);
                    var last = offset + count >= stdin.Length;
                    await SendStdinAsync(shellId, commandId, stdin.AsMemory(offset, count), last, ct).ConfigureAwait(false);
                }
            }
            else
            {
                await SendStdinAsync(shellId, commandId, ReadOnlyMemory<byte>.Empty, true, ct).ConfigureAwait(false);
            }

            var stdout = new MemoryStream();
            var stderr = new MemoryStream();
            var stderrLine = new StringBuilder();
            var exitCode = 0;
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                string response;
                try
                {
                    response = await transport.SendAsync(Envelope(
                        "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive", shellId, "PT20S", null,
                        $"<rsp:Receive><rsp:DesiredStream CommandId=\"{commandId}\">stdout stderr</rsp:DesiredStream></rsp:Receive>"),
                        ct).ConfigureAwait(false);
                }
                catch (WinRmException ex) when (WsManFault.IsTimeout(ex))
                {
                    continue; // nothing new within the operation timeout
                }

                var doc = XDocument.Parse(response);
                foreach (var stream in doc.Descendants(XName.Get("Stream", NsShell)))
                {
                    if (string.IsNullOrEmpty(stream.Value)) continue;
                    var bytes = Convert.FromBase64String(stream.Value);
                    if ((string?)stream.Attribute("Name") == "stderr")
                    {
                        stderr.Write(bytes);
                        if (onStderrLine is not null) EmitLines(stderrLine, Encoding.UTF8.GetString(bytes), onStderrLine);
                    }
                    else
                    {
                        stdout.Write(bytes);
                    }
                }

                var state = doc.Descendants(XName.Get("CommandState", NsShell)).FirstOrDefault();
                if (state is not null && ((string?)state.Attribute("State") ?? "").EndsWith("/Done", StringComparison.Ordinal))
                {
                    int.TryParse(state.Element(XName.Get("ExitCode", NsShell))?.Value, out exitCode);
                    break;
                }
            }
            if (stderrLine.Length > 0) onStderrLine?.Invoke(stderrLine.ToString());

            await TrySignalTerminateAsync(shellId, commandId).ConfigureAwait(false);
            return new RemoteCommandResult(exitCode, Encoding.UTF8.GetString(stdout.ToArray()), Encoding.UTF8.GetString(stderr.ToArray()));
        }
        finally
        {
            await TryDeleteShellAsync(shellId).ConfigureAwait(false);
        }
    }

    private static void EmitLines(StringBuilder pending, string text, Action<string> onLine)
    {
        pending.Append(text);
        while (true)
        {
            var s = pending.ToString();
            var nl = s.IndexOf('\n');
            if (nl < 0) return;
            onLine(s[..nl].TrimEnd('\r'));
            pending.Remove(0, nl + 1);
        }
    }

    private async Task<string> CreateShellAsync(CancellationToken ct)
    {
        var response = await transport.SendAsync(Envelope(
            "http://schemas.xmlsoap.org/ws/2004/09/transfer/Create", null, "PT60S",
            "<w:OptionSet><w:Option Name=\"WINRS_NOPROFILE\">TRUE</w:Option><w:Option Name=\"WINRS_CODEPAGE\">65001</w:Option></w:OptionSet>",
            "<rsp:Shell><rsp:InputStreams>stdin</rsp:InputStreams><rsp:OutputStreams>stdout stderr</rsp:OutputStreams></rsp:Shell>"),
            ct).ConfigureAwait(false);
        var doc = XDocument.Parse(response);
        var id = doc.Descendants(XName.Get("ShellId", NsShell)).FirstOrDefault()?.Value
                 ?? doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Selector" && (string?)e.Attribute("Name") == "ShellId")?.Value;
        return id ?? throw new WinRmException(WinRmErrorKind.Protocol, "WinRM did not return a shell id.");
    }

    private async Task<string> StartCommandAsync(string shellId, string command, string arguments, CancellationToken ct)
    {
        var response = await transport.SendAsync(Envelope(
            "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command", shellId, "PT60S",
            "<w:OptionSet><w:Option Name=\"WINRS_CONSOLEMODE_STDIN\">FALSE</w:Option><w:Option Name=\"WINRS_SKIP_CMD_SHELL\">TRUE</w:Option></w:OptionSet>",
            $"<rsp:CommandLine><rsp:Command>{SecurityElement.Escape(command)}</rsp:Command><rsp:Arguments>{SecurityElement.Escape(arguments)}</rsp:Arguments></rsp:CommandLine>"),
            ct).ConfigureAwait(false);
        var id = XDocument.Parse(response).Descendants(XName.Get("CommandId", NsShell)).FirstOrDefault()?.Value;
        return id ?? throw new WinRmException(WinRmErrorKind.Protocol, "WinRM did not return a command id.");
    }

    private Task SendStdinAsync(string shellId, string commandId, ReadOnlyMemory<byte> data, bool end, CancellationToken ct) =>
        transport.SendAsync(Envelope(
            "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Send", shellId, "PT60S", null,
            $"<rsp:Send><rsp:Stream Name=\"stdin\" CommandId=\"{commandId}\"{(end ? " End=\"true\"" : "")}>{Convert.ToBase64String(data.Span)}</rsp:Stream></rsp:Send>"),
            ct);

    private async Task TrySignalTerminateAsync(string shellId, string commandId)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            await transport.SendAsync(Envelope(
                "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Signal", shellId, "PT10S", null,
                $"<rsp:Signal CommandId=\"{commandId}\"><rsp:Code>http://schemas.microsoft.com/wbem/wsman/1/windows/shell/signal/terminate</rsp:Code></rsp:Signal>"),
                cts.Token).ConfigureAwait(false);
        }
        catch (Exception)
        {
            // best effort
        }
    }

    private async Task TryDeleteShellAsync(string shellId)
    {
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            await transport.SendAsync(Envelope("http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete", shellId, "PT10S", null, ""),
                cts.Token).ConfigureAwait(false);
        }
        catch (Exception)
        {
            // best effort: the shell expires on the host after its idle timeout
        }
    }

    private string Envelope(string action, string? shellId, string timeout, string? optionSet, string body)
    {
        var sb = new StringBuilder(1024 + body.Length);
        sb.Append("<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"")
          .Append(" xmlns:a=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"")
          .Append(" xmlns:w=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"")
          .Append(" xmlns:p=\"http://schemas.microsoft.com/wbem/wsman/1/wsman.xsd\"")
          .Append($" xmlns:rsp=\"{NsShell}\">")
          .Append("<s:Header>")
          .Append($"<a:To>{SecurityElement.Escape(transport.Endpoint.ToString())}</a:To>")
          .Append("<a:ReplyTo><a:Address s:mustUnderstand=\"true\">http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</a:Address></a:ReplyTo>")
          .Append($"<w:MaxEnvelopeSize s:mustUnderstand=\"true\">{MaxEnvelopeBytes}</w:MaxEnvelopeSize>")
          .Append($"<a:MessageID>uuid:{Guid.NewGuid().ToString().ToUpperInvariant()}</a:MessageID>")
          .Append("<w:Locale xml:lang=\"en-US\" s:mustUnderstand=\"false\"/>")
          .Append("<p:DataLocale xml:lang=\"en-US\" s:mustUnderstand=\"false\"/>")
          .Append($"<w:OperationTimeout>{timeout}</w:OperationTimeout>")
          .Append($"<w:ResourceURI s:mustUnderstand=\"true\">{ShellUri}</w:ResourceURI>")
          .Append($"<a:Action s:mustUnderstand=\"true\">{action}</a:Action>");
        if (shellId is not null)
            sb.Append($"<w:SelectorSet><w:Selector Name=\"ShellId\">{SecurityElement.Escape(shellId)}</w:Selector></w:SelectorSet>");
        if (optionSet is not null) sb.Append(optionSet);
        sb.Append("</s:Header><s:Body>").Append(body).Append("</s:Body></s:Envelope>");
        return sb.ToString();
    }
}
