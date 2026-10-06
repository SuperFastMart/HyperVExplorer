// Renders the app headlessly with demo data and saves PNG screenshots (used for docs and visual checks).
// Usage: dotnet run --project tools/HypervisorExplorer.Screenshots -- <output-dir>
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using HypervisorExplorer.App;
using HypervisorExplorer.App.ViewModels;
using HypervisorExplorer.App.Views;
using HypervisorExplorer.Core.Config;

var outDir = args.Length > 0 ? args[0] : "screenshots";
Directory.CreateDirectory(outDir);
var configDir = Path.Combine(Path.GetTempPath(), "hve-screens-" + Guid.NewGuid().ToString("N"));

AppBuilder.Configure<App>()
    .UseSkia()
    .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false })
    .WithInterFont()
    .SetupWithoutStarting();

void Pump()
{
    for (var i = 0; i < 5; i++)
    {
        Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }
}

void Save(Window w, string name)
{
    Pump();
    var frame = w.CaptureRenderedFrame();
    var path = Path.Combine(outDir, name + ".png");
    frame?.Save(path, (int?)null);
    Console.WriteLine(path);
}

var vm = new MainViewModel(new ConfigStore(configDir));
var window = new MainWindow { DataContext = vm, Width = 1600, Height = 940 };
window.Show();
Save(window, "01-welcome");

vm.LoadDemoCommand.Execute(null);
Pump();
Save(window, "02-overview");

vm.SelectedTable = vm.VisibleTables.First(t => t.Name == "vInfo");
var firstVm = vm.SelectedTable.AllRows.First(r => r.Vm is not null);
vm.OnRowSelected(firstVm);
Save(window, "03-vinfo-details");

vm.SelectedNode = vm.TreeRoots[0].SelfAndDescendants().First(n => n.Kind == TreeNodeKind.Cluster && n.Title == "HV-CLU01");
vm.SelectedTable = vm.VisibleTables.First(t => t.Name == "hvClusterNodes");
Save(window, "04-cluster-nodes");

vm.SelectedNode = vm.TreeRoots[0];
vm.SelectedTable = vm.VisibleTables.First(t => t.Name == "vHost");
vm.ShowActivity = true;
Save(window, "05-vhost-activity");

vm.SourcesExpanded = false;
Save(window, "09-sources-collapsed");
vm.SourcesExpanded = true;

vm.SearchText = "sql";
for (var i = 0; i < 20; i++)
{
    Thread.Sleep(30); // let the search debounce timer elapse; timers fire while pumping the UI thread
    Pump();
}
vm.SelectedTable = vm.VisibleTables.First(t => t.Name == "vInfo");
Save(window, "06-search");

var dialog = new ConnectDialog { DataContext = new ConnectDialogViewModel([]) { SelectedPlatform = ConnectDialogViewModel.PlatformOptions[2] } };
dialog.Show();
Save(dialog, "07-connect-vmware");
dialog.Close();

var store = new ConfigStore(configDir);
store.Config.Groups.Add(new HostGroup { Name = "Manchester", Username = "root@pam" });
foreach (var (addr, platform) in new[] { ("192.0.2.50", HypervisorExplorer.Core.Model.Platform.Proxmox), ("vcenter.example.com", HypervisorExplorer.Core.Model.Platform.VMware), ("hv-node01.example.com", HypervisorExplorer.Core.Model.Platform.HyperV) })
{
    store.Config.Hosts.Add(new SavedHost
    {
        Address = addr, Platform = platform, Username = "root", LastConnected = DateTimeOffset.Now,
        CredentialKind = HypervisorExplorer.Core.Collection.CredentialKind.UsernamePassword,
        GroupId = platform == HypervisorExplorer.Core.Model.Platform.Proxmox ? store.Config.Groups[0].Id : null,
    });
}
var hosts = new HostsWindow { DataContext = new HostsViewModel(store, window, _ => Task.CompletedTask) };
hosts.Show();
Save(hosts, "08-saved-hosts");
hosts.Close();

Directory.Delete(configDir, true);
