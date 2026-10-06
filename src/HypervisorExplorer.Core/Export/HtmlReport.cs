using System.Net;
using System.Text;
using HypervisorExplorer.Core.Model;
using HypervisorExplorer.Core.Tables;

namespace HypervisorExplorer.Core.Export;

/// <summary>Single-file HTML estate report: summary figures plus searchable, sortable tables.</summary>
public static class HtmlReport
{
    public static void Write(Inventory inv, string path) => File.WriteAllText(path, Render(inv), new UTF8Encoding(false));

    public static string Render(Inventory inv)
    {
        var vms = inv.VirtualMachines;
        var on = vms.Count(v => v.PowerState == PowerState.PoweredOn);
        var sb = new StringBuilder();
        sb.Append("""
            <!doctype html><html lang="en"><head><meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>Hypervisor Estate Report</title>
            <style>
            :root{--bg:#f7f8fa;--fg:#1d2330;--muted:#5d6678;--card:#fff;--line:#e2e5eb;--accent:#2f6fde;--warn:#b7791f;--err:#c53030}
            @media (prefers-color-scheme:dark){:root{--bg:#161a22;--fg:#e6e9ef;--muted:#9aa3b5;--card:#1f2430;--line:#2e3442;--accent:#7aa7ff;--warn:#f6c560;--err:#ff7b7b}}
            *{box-sizing:border-box}body{margin:0;font:14px/1.45 system-ui,-apple-system,Segoe UI,sans-serif;background:var(--bg);color:var(--fg)}
            header{padding:24px 32px 8px}h1{margin:0;font-size:22px}header p{margin:4px 0 0;color:var(--muted)}
            main{padding:8px 32px 48px}.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:16px 0 24px}
            .card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px}.card b{display:block;font-size:22px}.card span{color:var(--muted);font-size:12px}
            section{background:var(--card);border:1px solid var(--line);border-radius:8px;margin:0 0 20px;overflow:hidden}
            section h2{font-size:15px;margin:0;padding:12px 14px;border-bottom:1px solid var(--line);display:flex;gap:12px;align-items:center}
            section h2 small{color:var(--muted);font-weight:400}section h2 input{margin-left:auto;padding:5px 8px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--fg);min-width:200px}
            .wrap{overflow:auto;max-height:70vh}table{border-collapse:collapse;width:100%;font-size:12.5px}
            th,td{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
            th{position:sticky;top:0;background:var(--card);cursor:pointer;user-select:none;color:var(--muted);font-weight:600}
            td.n{text-align:right;font-variant-numeric:tabular-nums}.Warning{color:var(--warn)}.Error{color:var(--err)}
            @media (max-width:640px){header,main{padding-left:16px;padding-right:16px}section h2{flex-wrap:wrap}section h2 input{min-width:0;width:100%;margin-left:0}}
            </style></head><body>
            """);
        sb.Append($"<header><h1>Hypervisor Estate Report</h1><p>Generated {DateTime.Now:yyyy-MM-dd HH:mm} from {inv.Sources.Count} source(s): {E(string.Join(", ", inv.Sources.Select(s => s.Address)))}</p></header><main>");

        sb.Append("<div class=\"cards\">");
        Card(sb, inv.Hosts.Count.ToString("N0"), "Hosts");
        Card(sb, inv.Clusters.Count.ToString("N0"), "Clusters");
        Card(sb, vms.Count.ToString("N0"), $"VMs / containers ({on:N0} on)");
        Card(sb, vms.Sum(v => v.CpuCount).ToString("N0"), "vCPUs allocated");
        Card(sb, $"{vms.Sum(v => v.MemoryMiB) / 1024.0:N0} GiB", "vRAM allocated");
        Card(sb, $"{vms.Sum(v => v.ProvisionedBytes) / 1099511627776.0:N1} TiB", "Disk provisioned");
        Card(sb, inv.Datastores.Count.ToString("N0"), "Datastores");
        Card(sb, inv.Health.Count(h => h.Severity != HealthSeverity.Info).ToString("N0"), "Health warnings");
        sb.Append("</div>");

        var byPlatform = inv.Hosts.GroupBy(h => h.Platform)
            .Select(g => new
            {
                Platform = RvToolsTables.PlatformName(g.Key),
                Hosts = g.Count(),
                Cores = g.Sum(h => h.CpuCores ?? 0),
                MemGiB = g.Sum(h => (h.MemoryBytes ?? 0) / 1073741824.0),
                Vms = vms.Count(v => v.Platform == g.Key),
                VmsOn = vms.Count(v => v.Platform == g.Key && v.PowerState == PowerState.PoweredOn),
                VCpu = vms.Where(v => v.Platform == g.Key).Sum(v => v.CpuCount),
            }).ToList();
        sb.Append("<section><h2>Platforms</h2><div class=\"wrap\"><table><thead><tr><th>Platform</th><th>Hosts</th><th>Physical cores</th><th>Host RAM (GiB)</th><th>VMs</th><th>Powered on</th><th>vCPU</th><th>vCPU : core</th></tr></thead><tbody>");
        foreach (var p in byPlatform)
        {
            sb.Append($"<tr><td>{E(p.Platform)}</td><td class=n>{p.Hosts}</td><td class=n>{p.Cores}</td><td class=n>{p.MemGiB:N0}</td><td class=n>{p.Vms}</td><td class=n>{p.VmsOn}</td><td class=n>{p.VCpu}</td><td class=n>{(p.Cores > 0 ? (p.VCpu / (double)p.Cores).ToString("0.00") : "")}</td></tr>");
        }
        sb.Append("</tbody></table></div></section>");

        TableSection(sb, "Health", "hvHealth", RvToolsTables.Get("vHealth"), inv, severityColumn: "Message type");
        TableSection(sb, "Virtual machines", "hvOverview", ExtendedTables.All.First(t => t.Name == "hvOverview"), inv);
        TableSection(sb, "Hosts", "vHost", RvToolsTables.Get("vHost"), inv, onlyColumns:
            ["Host", "Datacenter", "Cluster", "CPU Model", "# CPU", "# Cores", "# Memory", "Memory usage %", "# VMs total", "# vCPUs", "vCPUs per Core", "ESX Version", "Vendor", "Model", "Serial number"]);
        TableSection(sb, "Clusters", "vCluster", RvToolsTables.Get("vCluster"), inv, onlyColumns:
            ["Name", "OverallStatus", "NumHosts", "numEffectiveHosts", "NumCpuCores", "TotalMemory", "HA enabled", "VI SDK Server"]);
        TableSection(sb, "Datastores", "vDatastore", RvToolsTables.Get("vDatastore"), inv, onlyColumns:
            ["Name", "Type", "Cluster name", "# VMs total", "Capacity MiB", "Provisioned MiB", "In Use MiB", "Free MiB", "Free %", "Hosts"]);
        TableSection(sb, "Snapshots", "vSnapshot", RvToolsTables.Get("vSnapshot"), inv, onlyColumns:
            ["VM", "Name", "Description", "Date / time", "Size MiB (total)", "Host", "VI SDK Server"]);

        sb.Append("</main><script>");
        sb.Append("""
            document.querySelectorAll('input[data-t]').forEach(i=>i.addEventListener('input',()=>{const q=i.value.toLowerCase();document.querySelectorAll('#'+i.dataset.t+' tbody tr').forEach(r=>r.style.display=r.textContent.toLowerCase().includes(q)?'':'none')}));
            document.querySelectorAll('th').forEach(th=>th.addEventListener('click',()=>{const t=th.closest('table'),b=t.tBodies[0],i=[...th.parentNode.children].indexOf(th),asc=th.dataset.s!=='a';th.parentNode.querySelectorAll('th').forEach(x=>delete x.dataset.s);th.dataset.s=asc?'a':'d';
            const v=r=>{const s=r.children[i]?.textContent??'',n=parseFloat(s.replace(/,/g,''));return isNaN(n)||!/^[-\d.,]+$/.test(s)?s.toLowerCase():n};
            [...b.rows].sort((x,y)=>{const a=v(x),c=v(y);return (a>c?1:a<c?-1:0)*(asc?1:-1)}).forEach(r=>b.appendChild(r))}));
            """);
        sb.Append("</script></body></html>");
        return sb.ToString();
    }

    private static void Card(StringBuilder sb, string value, string label) =>
        sb.Append($"<div class=\"card\"><b>{E(value)}</b><span>{E(label)}</span></div>");

    private static void TableSection(StringBuilder sb, string title, string id, TableDefinition table, Inventory inv,
        string[]? onlyColumns = null, string? severityColumn = null)
    {
        var rows = table.Rows(inv);
        if (rows.Count == 0) return;
        var cols = onlyColumns is null
            ? table.Columns.ToList()
            : onlyColumns.Select(h => table.Columns.First(c => c.Header == h)).ToList();
        var sevIdx = severityColumn is null ? -1 : cols.FindIndex(c => c.Header == severityColumn);

        sb.Append($"<section><h2>{E(title)} <small>{rows.Count:N0}</small><input type=\"search\" placeholder=\"Filter…\" data-t=\"{id}\" aria-label=\"Filter {E(title)}\"></h2><div class=\"wrap\"><table id=\"{id}\"><thead><tr>");
        foreach (var c in cols) sb.Append($"<th>{E(c.Header)}</th>");
        sb.Append("</tr></thead><tbody>");
        foreach (var row in rows)
        {
            var values = cols.Select(c => c.GetValue(row)).ToList();
            var cls = sevIdx >= 0 ? $" class=\"{E(TableColumn.Format(values[sevIdx]))}\"" : "";
            sb.Append($"<tr{cls}>");
            foreach (var v in values)
            {
                var numeric = v is int or long or double or float or decimal;
                sb.Append(numeric ? $"<td class=n>{E(TableColumn.Format(v))}</td>" : $"<td>{E(TableColumn.Format(v))}</td>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table></div></section>");
    }

    private static string E(string s) => WebUtility.HtmlEncode(s);
}
