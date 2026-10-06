using HypervisorExplorer.Collectors;
using HypervisorExplorer.Collectors.HyperV;
using HypervisorExplorer.Core.Collection;
using HypervisorExplorer.Core.Config;
using HypervisorExplorer.Core.Export;
using HypervisorExplorer.Core.Model;

// Command-line collection and export, for scheduled runs (Task Scheduler / cron) and scripting.
return await Cli.RunAsync(args);

internal static class Cli
{
    private const string Usage = """
        Hypervisor Explorer CLI

        Usage:
          hvexplorer collect [sources...] -o <file> [options]
          hvexplorer convert <snapshot.json | hyperv-collection.json> -o <file>
          hvexplorer script -o Collect-HyperV.ps1

        Sources (repeatable):
          --vmware <host[:port]>       vCenter or ESXi (vSphere API)
          --proxmox <host[:port]>      Proxmox VE node (whole cluster is collected)
          --hyperv <host>              Hyper-V host (Windows only; cluster nodes are expanded)
          --saved                      every host saved in the GUI (uses stored credentials)
          --saved-group <name>         saved hosts in one group

        Credentials (apply to the sources that follow them):
          -u, --user <name>            username, or Proxmox token id (user@realm!token)
          --password-env <VAR>         read the password / token secret from an environment variable
          --password-file <path>       read it from the first line of a file
          --token                      treat credentials as a Proxmox API token
          --current-user               Hyper-V: use Windows integrated auth
          --group <label>              site label used as the RVTools 'Datacenter'
          --strict-tls                 reject untrusted certificates (default: accept self-signed)

        Output:
          -o, --out <file>             .xlsx (RVTools workbook), .json (snapshot), .html (report), .zip (CSV per table)
          --parallel <n>               concurrent collections (default 4)

        Exit codes: 0 success, 1 usage error, 2 some sources failed (output still written), 3 all failed.
        """;

    public static async Task<int> RunAsync(string[] args)
    {
        if (args.Length == 0 || args[0] is "-h" or "--help" or "help")
        {
            Console.WriteLine(Usage);
            return args.Length == 0 ? 1 : 0;
        }

        try
        {
            return args[0] switch
            {
                "collect" => await CollectAsync(args[1..]),
                "convert" => Convert(args[1..]),
                "script" => Script(args[1..]),
                _ => Fail($"Unknown command '{args[0]}'."),
            };
        }
        catch (ArgumentException ex)
        {
            return Fail(ex.Message);
        }
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine("error: " + message);
        Console.Error.WriteLine("Run 'hvexplorer --help' for usage.");
        return 1;
    }

    private static async Task<int> CollectAsync(string[] args)
    {
        var requests = new List<ConnectionRequest>();
        string? output = null;
        string? user = null, secret = null, group = null;
        var token = false;
        var currentUser = false;
        var strictTls = false;
        var parallel = 4;
        ConfigStore? config = null;

        ConnectionRequest Make(Platform platform, string target)
        {
            var (address, port) = SplitHostPort(target);
            var kind = currentUser ? CredentialKind.CurrentUser
                : token ? CredentialKind.ApiToken
                : platform == Platform.HyperV && user is null ? CredentialKind.CurrentUser
                : CredentialKind.UsernamePassword;
            return new ConnectionRequest
            {
                Platform = platform, Address = address, Port = port, CredentialKind = kind,
                Username = user, Secret = secret, Group = group, IgnoreCertificateErrors = !strictTls,
            };
        }

        for (var i = 0; i < args.Length; i++)
        {
            string Next() => i + 1 < args.Length ? args[++i] : throw new ArgumentException($"{args[i]} needs a value.");
            switch (args[i])
            {
                case "--vmware": requests.Add(Make(Platform.VMware, Next())); break;
                case "--proxmox": requests.Add(Make(Platform.Proxmox, Next())); break;
                case "--hyperv": requests.Add(Make(Platform.HyperV, Next())); break;
                case "-u" or "--user": user = Next(); break;
                case "--password-env":
                    var name = Next();
                    secret = Environment.GetEnvironmentVariable(name) ?? throw new ArgumentException($"Environment variable {name} is not set.");
                    break;
                case "--password-file": secret = File.ReadLines(Next()).FirstOrDefault()?.TrimEnd() ?? ""; break;
                case "--token": token = true; break;
                case "--current-user": currentUser = true; break;
                case "--group": group = Next(); break;
                case "--strict-tls": strictTls = true; break;
                case "-o" or "--out": output = Next(); break;
                case "--parallel": parallel = int.Parse(Next()); break;
                case "--saved":
                    config ??= new ConfigStore();
                    requests.AddRange(config.Config.Hosts.Select(config.BuildRequest));
                    break;
                case "--saved-group":
                    config ??= new ConfigStore();
                    var g = Next();
                    var grp = config.Config.Groups.FirstOrDefault(x => string.Equals(x.Name, g, StringComparison.OrdinalIgnoreCase))
                              ?? throw new ArgumentException($"No saved group named '{g}'.");
                    requests.AddRange(config.Config.Hosts.Where(h => h.GroupId == grp.Id).Select(config.BuildRequest));
                    break;
                default: throw new ArgumentException($"Unknown option '{args[i]}'.");
            }
        }

        if (requests.Count == 0) throw new ArgumentException("No sources given.");
        if (output is null) throw new ArgumentException("No output file given (-o).");
        foreach (var r in requests.Where(r => !ConfigStore.HasUsableCredentials(r)))
            throw new ArgumentException($"{r.Address}: no credentials (use -u and --password-env, or save them in the GUI).");

        var store = new InventoryStore();
        var failures = 0;
        using var gate = new SemaphoreSlim(Math.Max(1, parallel));
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

        await Task.WhenAll(requests.Select(async request =>
        {
            await gate.WaitAsync(cts.Token);
            try
            {
                var progress = new Progress<string>(m => Log(request.Address, m));
                var snap = await CollectorRegistry.Create(request.Platform).CollectAsync(request, progress, cts.Token);
                snap.RekeySource(request.DisplayName);
                store.Upsert(snap);
                foreach (var w in snap.Warnings) Log(request.Address, "warning: " + w);
                Log(request.Address, $"OK: {snap.Hosts.Count} host(s), {snap.VirtualMachines.Count} VM(s)");
            }
            catch (CollectionException ex)
            {
                Interlocked.Increment(ref failures);
                Log(request.Address, $"FAILED ({ex.Kind}): {ex.Message}" + (ex.Hint is null ? "" : $"\n    hint: {ex.Hint}"));
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Interlocked.Increment(ref failures);
                Log(request.Address, "FAILED: " + ex.Message);
            }
            finally
            {
                gate.Release();
            }
        }));

        if (failures == requests.Count)
        {
            Console.Error.WriteLine("All sources failed; nothing written.");
            return 3;
        }
        Write(store.Current, store.Snapshots, output);
        return failures > 0 ? 2 : 0;
    }

    private static int Convert(string[] args)
    {
        string? input = null, output = null;
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i] is "-o" or "--out") output = i + 1 < args.Length ? args[++i] : null;
            else input = args[i];
        }
        if (input is null || output is null) throw new ArgumentException("convert needs an input file and -o <file>.");

        List<InventorySnapshot> snaps;
        try
        {
            snaps = InventoryJson.Load(input);
        }
        catch (Exception ex) when (ex is InvalidDataException or System.Text.Json.JsonException)
        {
            // Not a snapshot: try a Hyper-V collection file produced by Collect-HyperV.ps1.
            var request = new ConnectionRequest { Platform = Platform.HyperV, Address = Path.GetFileNameWithoutExtension(input) };
            snaps = [HyperVJsonMapper.MapFile(input, request)];
        }
        Write(Inventory.FromSnapshots(snaps), snaps, output);
        return 0;
    }

    private static int Script(string[] args)
    {
        var output = args.SkipWhile(a => a is not ("-o" or "--out")).Skip(1).FirstOrDefault() ?? "Collect-HyperV.ps1";
        File.WriteAllText(output, HyperVCollector.GetCollectionScript(), new System.Text.UTF8Encoding(true));
        Console.WriteLine($"Wrote {output}");
        return 0;
    }

    private static void Write(Inventory inventory, IReadOnlyList<InventorySnapshot> snapshots, string output)
    {
        switch (Path.GetExtension(output).ToLowerInvariant())
        {
            case ".xlsx": InventoryExporter.ExportXlsx(inventory, output); break;
            case ".json": InventoryJson.Save(snapshots, output); break;
            case ".html" or ".htm": HtmlReport.Write(inventory, output); break;
            case ".zip": InventoryExporter.ExportCsvZip(inventory, output); break;
            default: throw new ArgumentException("Output must end in .xlsx, .json, .html or .zip.");
        }
        Console.WriteLine($"Wrote {output}: {inventory.Sources.Count} source(s), {inventory.Hosts.Count} host(s), {inventory.VirtualMachines.Count} VM(s).");
    }

    private static (string Address, int? Port) SplitHostPort(string target)
    {
        var t = target.Trim();
        if (t.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) t = t[8..];
        t = t.TrimEnd('/');
        var colon = t.LastIndexOf(':');
        return colon > 0 && t.Count(c => c == ':') == 1 && int.TryParse(t[(colon + 1)..], out var port)
            ? (t[..colon], port)
            : (t, null);
    }

    private static readonly object LogGate = new();

    private static void Log(string source, string message)
    {
        lock (LogGate) Console.Error.WriteLine($"{DateTime.Now:HH:mm:ss} [{source}] {message}");
    }
}
