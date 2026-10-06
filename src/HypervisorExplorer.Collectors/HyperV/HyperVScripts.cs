namespace HypervisorExplorer.Collectors.HyperV;

/// <summary>Access to the embedded PowerShell scripts.</summary>
public static class HyperVScripts
{
    public const string CollectScriptName = "Collect-HyperV.ps1";
    public const string WrapperScriptName = "Invoke-Collection.ps1";

    private static readonly Lazy<string> Collect = new(() => Load(CollectScriptName));
    private static readonly Lazy<string> Wrapper = new(() => Load(WrapperScriptName));

    /// <summary>The remote collection script (runs on the Hyper-V host; can be run by hand).</summary>
    public static string CollectScript => Collect.Value;

    /// <summary>The local launcher script (Invoke-Command, cluster fan-out, envelope).</summary>
    public static string WrapperScript => Wrapper.Value;

    /// <summary>
    /// Tiny script passed via -EncodedCommand. It reads the base64 UTF-8 JSON request from stdin (so secrets never
    /// touch the command line and the large scripts don't hit the 32K command-line limit) and runs the wrapper.
    /// </summary>
    public const string Bootstrap = """
        $ErrorActionPreference = 'Stop'
        try {
            $raw = [string]($input | Out-String)
            if (-not $raw.Trim()) { $raw = [Console]::In.ReadToEnd() }
            $b64 = $raw -replace '[^A-Za-z0-9+/=]', ''
            $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($b64))
            $req = ConvertFrom-Json -InputObject $json
            $wrapper = [scriptblock]::Create([string]$req.wrapper)
        } catch {
            $o = [ordered]@{ ok = $false; stage = 'connect'; error = ('Launcher failed: ' + $_.Exception.Message); category = [string]$_.FullyQualifiedErrorId }
            [Console]::Out.WriteLine((ConvertTo-Json -InputObject $o -Compress))
            exit 1
        }
        & $wrapper -Request $req
        """;

    private static string Load(string fileName)
    {
        var asm = typeof(HyperVScripts).Assembly;
        var name = asm.GetManifestResourceNames().FirstOrDefault(n => n.EndsWith("." + fileName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Embedded script {fileName} not found.");
        using var stream = asm.GetManifestResourceStream(name)!;
        using var reader = new StreamReader(stream, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }
}
