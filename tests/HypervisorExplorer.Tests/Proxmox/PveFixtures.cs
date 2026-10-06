namespace HypervisorExplorer.Tests.Proxmox;

/// <summary>
/// Realistic PVE 8.2 API responses for a 2-node cluster ("lab-cluster": pve1, pve2) with:
/// VM 100 "web01" on pve1 (running, guest agent, scsi0 on local-lvm + scsi1 qcow2 on NFS + ISO cdrom + efidisk,
/// 2 NICs one tagged VLAN 20, one snapshot with RAM, HA-managed), CT 101 "ct01" on pve2 (running, bind mount),
/// template 9000 "tpl-debian" on pve2, storages local (dir), local-lvm (lvmthin) and nfs-shared (shared NFS).
/// </summary>
public static class PveFixtures
{
    public const string Version = """{"version":"8.2.4","release":"8.2","repoid":"faa83925c9641325","console":"xtermjs"}""";

    public const string Ticket = """
        {"username":"root@pam","ticket":"PVE:root@pam:66A1B2C3::c2lnbmF0dXJl+/=","CSRFPreventionToken":"66A1B2C3:Y3NyZg","cap":{"vms":{}}}
        """;

    public const string ClusterStatus = """
        [
          {"type":"cluster","id":"cluster","name":"lab-cluster","nodes":2,"quorate":1,"version":5},
          {"type":"node","id":"node/pve1","name":"pve1","nodeid":1,"ip":"10.0.0.11","online":1,"local":1,"level":""},
          {"type":"node","id":"node/pve2","name":"pve2","nodeid":2,"ip":"10.0.0.12","online":1,"local":0,"level":""}
        ]
        """;

    public const string StandaloneStatus = """
        [{"type":"node","id":"node/pve1","name":"pve1","nodeid":0,"ip":"10.0.0.11","online":1,"local":1,"level":""}]
        """;

    public const string Resources = """
        [
          {"id":"node/pve1","type":"node","node":"pve1","status":"online","cpu":0.0531,"maxcpu":64,"mem":54000000000,"maxmem":270000000000,"disk":9000000000,"maxdisk":100000000000,"uptime":1209600,"level":"","cgroup-mode":2},
          {"id":"node/pve2","type":"node","node":"pve2","status":"online","cpu":0.01,"maxcpu":16,"mem":8000000000,"maxmem":68000000000,"disk":5000000000,"maxdisk":100000000000,"uptime":600000,"level":"","cgroup-mode":2},
          {"id":"qemu/100","type":"qemu","vmid":100,"name":"web01","node":"pve1","status":"running","template":0,"pool":"prod","tags":"web;prod","maxcpu":4,"maxmem":8589934592,"mem":3221225472,"maxdisk":34359738368,"disk":0,"uptime":86400,"hastate":"started","cpu":0.02,"netin":1,"netout":1,"diskread":1,"diskwrite":1},
          {"id":"lxc/101","type":"lxc","vmid":101,"name":"ct01","node":"pve2","status":"running","template":0,"tags":"infra","maxcpu":2,"maxmem":1073741824,"mem":268435456,"maxdisk":8589934592,"disk":1610612736,"uptime":3600,"cpu":0.001},
          {"id":"qemu/9000","type":"qemu","vmid":9000,"name":"tpl-debian","node":"pve2","status":"stopped","template":1,"maxcpu":1,"maxmem":2147483648,"maxdisk":10737418240,"disk":0,"uptime":0},
          {"id":"storage/pve1/local","type":"storage","storage":"local","node":"pve1","status":"available","plugintype":"dir","content":"iso,vztmpl,backup","shared":0,"disk":20000000000,"maxdisk":100000000000},
          {"id":"storage/pve1/local-lvm","type":"storage","storage":"local-lvm","node":"pve1","status":"available","plugintype":"lvmthin","content":"rootdir,images","shared":0,"disk":200000000000,"maxdisk":800000000000},
          {"id":"storage/pve1/nfs-shared","type":"storage","storage":"nfs-shared","node":"pve1","status":"available","plugintype":"nfs","content":"images,iso","shared":1,"disk":1000000000000,"maxdisk":4000000000000},
          {"id":"storage/pve2/local","type":"storage","storage":"local","node":"pve2","status":"available","plugintype":"dir","content":"iso,vztmpl,backup","shared":0,"disk":10000000000,"maxdisk":100000000000},
          {"id":"storage/pve2/local-lvm","type":"storage","storage":"local-lvm","node":"pve2","status":"available","plugintype":"lvmthin","content":"rootdir,images","shared":0,"disk":50000000000,"maxdisk":400000000000},
          {"id":"storage/pve2/nfs-shared","type":"storage","storage":"nfs-shared","node":"pve2","status":"available","plugintype":"nfs","content":"images,iso","shared":1,"disk":1000000000000,"maxdisk":4000000000000},
          {"id":"pool/prod","type":"pool","pool":"prod"}
        ]
        """;

    public const string StorageConfig = """
        [
          {"storage":"local","type":"dir","path":"/var/lib/vz","content":"iso,vztmpl,backup","digest":"x"},
          {"storage":"local-lvm","type":"lvmthin","vgname":"pve","thinpool":"data","content":"rootdir,images","digest":"x"},
          {"storage":"nfs-shared","type":"nfs","server":"10.0.0.5","export":"/export/pve","path":"/mnt/pve/nfs-shared","content":"images,iso","options":"vers=4.2","digest":"x"}
        ]
        """;

    public const string HaResources = """
        [{"sid":"vm:100","type":"vm","state":"started","group":"prefer-pve1","max_restart":1,"max_relocate":1,"digest":"x"}]
        """;

    public const string HaStatus = """
        [
          {"id":"quorum","type":"quorum","status":"OK","quorate":1,"node":"pve1"},
          {"id":"master","type":"master","node":"pve1","status":"pve1 (active, Tue Jun  4 12:00:00 2024)","timestamp":1717500000},
          {"id":"lrm:pve1","type":"lrm","node":"pve1","status":"pve1 (active, Tue Jun  4 12:00:00 2024)","timestamp":1717500000},
          {"id":"lrm:pve2","type":"lrm","node":"pve2","status":"pve2 (idle, Tue Jun  4 12:00:00 2024)","timestamp":1717500000},
          {"id":"service:vm:100","type":"service","sid":"vm:100","node":"pve1","state":"started","crm_state":"started","request_state":"started","max_restart":1,"max_relocate":1}
        ]
        """;

    public const string HaGroups = """
        [{"group":"prefer-pve1","type":"group","nodes":"pve2:1,pve1:2","restricted":0,"nofailback":0,"digest":"x"}]
        """;

    public const string Pve1Status = """
        {"cpu":0.0531,"cpuinfo":{"model":"Intel(R) Xeon(R) Silver 4314 CPU @ 2.40GHz","sockets":2,"cores":32,"cpus":64,"mhz":"2400.000","hvm":"1","user_hz":100,"flags":"fpu vme"},
         "memory":{"total":270000000000,"used":54000000000,"free":216000000000},"swap":{"total":0,"used":0,"free":0},
         "uptime":1209600,"loadavg":["0.10","0.20","0.30"],"pveversion":"pve-manager/8.2.4/faa83925c9641325",
         "kversion":"Linux 6.8.8-2-pve #1 SMP PREEMPT_DYNAMIC PMX 6.8.8-2 (2024-06-24T09:00Z) x86_64",
         "current-kernel":{"release":"6.8.8-2-pve","sysname":"Linux","machine":"x86_64","version":"#1 SMP PREEMPT_DYNAMIC PMX 6.8.8-2"},
         "boot-info":{"mode":"efi","secureboot":0},"rootfs":{"total":100000000000,"used":9000000000,"avail":91000000000,"free":91000000000},
         "ksm":{"shared":0},"wait":0.0}
        """;

    public const string Pve2Status = """
        {"cpu":0.01,"cpuinfo":{"model":"AMD EPYC 4344P 8-Core Processor","sockets":1,"cores":8,"cpus":16,"mhz":"3800.000","hvm":"1","user_hz":100},
         "memory":{"total":68000000000,"used":8000000000,"free":60000000000},"uptime":600000,
         "pveversion":"pve-manager/8.2.4/faa83925c9641325","kversion":"Linux 6.8.8-2-pve #1 SMP PREEMPT_DYNAMIC PMX 6.8.8-2 (2024-06-24T09:00Z) x86_64"}
        """;

    public const string Dns = """{"search":"lab.local","dns1":"10.0.0.2","dns2":"1.1.1.1"}""";

    public const string Time = """{"timezone":"Europe/London","time":1717500000,"localtime":1717503600}""";

    public const string Pve1Network = """
        [
          {"iface":"eno1","type":"eth","method":"manual","families":["inet"],"active":1,"exists":1,"priority":3},
          {"iface":"eno2","type":"eth","method":"manual","families":["inet"],"active":1,"exists":1,"priority":4},
          {"iface":"bond0","type":"bond","method":"manual","families":["inet"],"active":1,"slaves":"eno1 eno2","bond_mode":"802.3ad","bond_xmit_hash_policy":"layer2+3","priority":5},
          {"iface":"vmbr0","type":"bridge","method":"static","families":["inet"],"active":1,"autostart":1,"bridge_ports":"bond0","bridge_stp":"off","bridge_fd":"0","bridge_vlan_aware":1,"bridge_vids":"2-4094","address":"10.0.0.11","netmask":"24","cidr":"10.0.0.11/24","gateway":"10.0.0.1","mtu":"9000","priority":6},
          {"iface":"vmbr0.20","type":"vlan","method":"static","families":["inet"],"active":1,"autostart":1,"vlan-id":"20","vlan-raw-device":"vmbr0","address":"10.20.0.11","netmask":"24","cidr":"10.20.0.11/24","priority":7}
        ]
        """;

    public const string Pve2Network = """
        [
          {"iface":"eno1","type":"eth","method":"manual","families":["inet"],"active":1,"exists":1},
          {"iface":"vmbr0","type":"bridge","method":"static","families":["inet"],"active":1,"bridge_ports":"eno1","address":"10.0.0.12","netmask":"255.255.255.0","gateway":"10.0.0.1","cidr":"10.0.0.12/24"}
        ]
        """;

    public const string Pve1Pci = """
        [
          {"id":"0000:00:17.0","class":"0x010601","vendor":"0x8086","device":"0xa182","vendor_name":"Intel Corporation","device_name":"C620 Series Chipset SATA Controller","iommugroup":12},
          {"id":"0000:3b:00.0","class":"0x010802","vendor":"0x144d","device":"0xa80a","vendor_name":"Samsung Electronics Co Ltd","device_name":"NVMe SSD Controller PM9A1","iommugroup":40},
          {"id":"0000:18:00.0","class":"0x020000","vendor":"0x8086","device":"0x1572","vendor_name":"Intel Corporation","device_name":"Ethernet Controller X710","iommugroup":30}
        ]
        """;

    public const string Pve1LocalLvmContent = """
        [
          {"volid":"local-lvm:vm-100-disk-0","format":"raw","size":34359738368,"used":12884901888,"vmid":"100","content":"images","ctime":1704067200},
          {"volid":"local-lvm:vm-100-disk-1","format":"raw","size":4194304,"vmid":"100","content":"images","ctime":1704067200}
        ]
        """;

    public const string Pve2LocalLvmContent = """
        [
          {"volid":"local-lvm:vm-101-disk-0","format":"raw","size":8589934592,"vmid":"101","content":"rootdir"},
          {"volid":"local-lvm:base-9000-disk-0","format":"raw","size":10737418240,"vmid":"9000","content":"images"}
        ]
        """;

    public const string NfsContent = """
        [
          {"volid":"nfs-shared:100/vm-100-disk-0.qcow2","format":"qcow2","size":107374182400,"used":21474836480,"vmid":"100","content":"images"},
          {"volid":"nfs-shared:iso/debian-12.5.0-amd64-netinst.iso","format":"iso","size":659554304,"content":"iso"}
        ]
        """;

    public const string Vm100Config = """
        {"name":"web01","cores":2,"sockets":2,"cpu":"x86-64-v2-AES","memory":"8192","balloon":2048,"bios":"ovmf",
         "efidisk0":"local-lvm:vm-100-disk-1,efitype=4m,pre-enrolled-keys=1,size=4M","machine":"pc-q35-8.1","ostype":"l26",
         "description":"Web front end\nOwner: ops\n","smbios1":"uuid=7b0c3a62-1d4e-4c1a-9a55-6e1f2a3b4c5d",
         "vmgenid":"0f6a6d0e-2c1b-4a5e-8f3d-1a2b3c4d5e6f","onboot":1,"startup":"order=1,up=30","tags":"web;prod",
         "boot":"order=scsi0;ide2;net0","scsihw":"virtio-scsi-single",
         "scsi0":"local-lvm:vm-100-disk-0,iothread=1,discard=on,ssd=1,size=32G",
         "scsi1":"nfs-shared:100/vm-100-disk-0.qcow2,cache=writeback,size=100G",
         "ide2":"local:iso/debian-12.5.0-amd64-netinst.iso,media=cdrom,size=629M",
         "net0":"virtio=BC:24:11:AA:BB:01,bridge=vmbr0,firewall=1",
         "net1":"virtio=BC:24:11:AA:BB:02,bridge=vmbr0,tag=20",
         "agent":"1,fstrim_cloned_disks=1","numa":0,"hotplug":"disk,network,usb","usb0":"host=046d:c52b",
         "hostpci0":"0000:3b:00.0,pcie=1","meta":"creation-qemu=8.1.5,ctime=1704067200","digest":"0123456789abcdef"}
        """;

    public const string Vm100Status = """
        {"status":"running","qmpstatus":"running","uptime":86400,"cpus":4,"maxmem":8589934592,"mem":3221225472,
         "balloon":6442450944,"ballooninfo":{"actual":6442450944,"max_mem":8589934592,"total_mem":6291456000,"free_mem":3000000000},
         "agent":1,"ha":{"managed":1,"state":"started","group":"prefer-pve1"},"name":"web01","pid":12345,
         "running-qemu":"8.1.5","running-machine":"pc-q35-8.1+pve0","vmid":100,"cpu":0.02,"disk":0,"maxdisk":34359738368}
        """;

    public const string Vm100Snapshots = """
        [
          {"name":"pre-upgrade","description":"Before apt upgrade\n","snaptime":1717200000,"vmstate":1},
          {"name":"current","description":"You are here!","parent":"pre-upgrade","running":1,"digest":"x"}
        ]
        """;

    public const string AgentInfo = """{"result":{"version":"7.2.11","supported_commands":[{"name":"guest-ping","enabled":true}]}}""";

    public const string AgentOsInfo = """
        {"result":{"id":"debian","kernel-release":"6.1.0-21-amd64","kernel-version":"#1 SMP PREEMPT_DYNAMIC Debian 6.1.90-1","machine":"x86_64","name":"Debian GNU/Linux","pretty-name":"Debian GNU/Linux 12 (bookworm)","version":"12 (bookworm)","version-id":"12"}}
        """;

    public const string AgentHostName = """{"result":{"host-name":"web01.lab.local"}}""";

    public const string AgentNetwork = """
        {"result":[
          {"name":"lo","hardware-address":"00:00:00:00:00:00","ip-addresses":[{"ip-address":"127.0.0.1","ip-address-type":"ipv4","prefix":8},{"ip-address":"::1","ip-address-type":"ipv6","prefix":128}]},
          {"name":"ens18","hardware-address":"bc:24:11:aa:bb:01","ip-addresses":[{"ip-address":"10.0.0.50","ip-address-type":"ipv4","prefix":24},{"ip-address":"fe80::be24:11ff:feaa:bb01","ip-address-type":"ipv6","prefix":64}],"statistics":{"rx-bytes":1}},
          {"name":"ens19","hardware-address":"bc:24:11:aa:bb:02","ip-addresses":[{"ip-address":"10.20.0.50","ip-address-type":"ipv4","prefix":24}]}
        ]}
        """;

    public const string AgentFsInfo = """
        {"result":[
          {"name":"sda1","mountpoint":"/","type":"ext4","total-bytes":33501757440,"used-bytes":5368709120,"disk":[{"bus-type":"scsi","serial":"drive-scsi0"}]},
          {"name":"sdb1","mountpoint":"/srv","type":"xfs","total-bytes":107321753600,"used-bytes":1073741824,"disk":[]},
          {"name":"sda15","mountpoint":"/boot/efi","type":"vfat","total-bytes":129718272,"used-bytes":12124160,"disk":[]},
          {"name":"tmpfs","mountpoint":"/run","type":"tmpfs","total-bytes":0,"used-bytes":0,"disk":[]}
        ]}
        """;

    public const string Ct101Config = """
        {"hostname":"ct01","arch":"amd64","cores":2,"memory":1024,"swap":512,"ostype":"debian",
         "rootfs":"local-lvm:vm-101-disk-0,size=8G","mp0":"/mnt/bindmounts/shared,mp=/shared",
         "net0":"name=eth0,bridge=vmbr0,firewall=1,hwaddr=BC:24:11:CC:DD:01,ip=10.0.0.60/24,gw=10.0.0.1,ip6=dhcp,tag=30,type=veth",
         "unprivileged":1,"features":"nesting=1","onboot":0,"tags":"infra","description":"Infra container","digest":"x"}
        """;

    public const string Ct101Status = """
        {"status":"running","uptime":3600,"cpus":2,"maxmem":1073741824,"mem":268435456,"maxswap":536870912,"swap":0,
         "disk":1610612736,"maxdisk":8589934592,"ha":{"managed":0},"name":"ct01","vmid":"101","type":"lxc","cpu":0.001}
        """;

    public const string Ct101Snapshots = """[{"name":"current","description":"You are here!","running":1,"digest":"x"}]""";

    public const string Ct101Interfaces = """
        [
          {"name":"lo","hwaddr":"00:00:00:00:00:00","inet":"127.0.0.1/8","inet6":"::1/128"},
          {"name":"eth0","hwaddr":"bc:24:11:cc:dd:01","inet":"10.0.0.60/24","inet6":"fd00::60/64"}
        ]
        """;

    public const string Tpl9000Config = """
        {"name":"tpl-debian","cores":1,"memory":"2048","ostype":"l26","scsi0":"local-lvm:base-9000-disk-0,size=10G",
         "net0":"virtio=BC:24:11:EE:00:01,bridge=vmbr0","template":1,"scsihw":"virtio-scsi-pci","agent":"1","digest":"x"}
        """;

    /// <summary>A handler serving the full 2-node cluster. Tests override individual routes.</summary>
    public static FakePveHandler Cluster()
    {
        var h = new FakePveHandler()
            .Data("/version", Version)
            .Data("/access/ticket", Ticket)
            .Data("/cluster/status", ClusterStatus)
            .Data("/cluster/resources", Resources)
            .Data("/storage", StorageConfig)
            .Data("/cluster/ha/resources", HaResources)
            .Data("/cluster/ha/status/current", HaStatus)
            .Data("/cluster/ha/groups", HaGroups)
            .Data("/nodes/pve1/status", Pve1Status)
            .Data("/nodes/pve2/status", Pve2Status)
            .Data("/nodes/pve1/dns", Dns)
            .Data("/nodes/pve2/dns", Dns)
            .Data("/nodes/pve1/time", Time)
            .Data("/nodes/pve2/time", Time)
            .Data("/nodes/pve1/network", Pve1Network)
            .Data("/nodes/pve2/network", Pve2Network)
            .Data("/nodes/pve1/hardware/pci", Pve1Pci)
            .Status("/nodes/pve2/hardware/pci", System.Net.HttpStatusCode.Forbidden, "Permission check failed (/nodes/pve2, Sys.Modify)")
            .Data("/nodes/pve1/storage/local-lvm/content", Pve1LocalLvmContent)
            .Data("/nodes/pve2/storage/local-lvm/content", Pve2LocalLvmContent)
            .Data("/nodes/pve1/storage/nfs-shared/content", NfsContent)
            .Data("/nodes/pve2/storage/nfs-shared/content", NfsContent)
            .Data("/nodes/pve1/qemu/100/config", Vm100Config)
            .Data("/nodes/pve1/qemu/100/status/current", Vm100Status)
            .Data("/nodes/pve1/qemu/100/snapshot", Vm100Snapshots)
            .Data("/nodes/pve1/qemu/100/agent/info", AgentInfo)
            .Data("/nodes/pve1/qemu/100/agent/get-osinfo", AgentOsInfo)
            .Data("/nodes/pve1/qemu/100/agent/get-host-name", AgentHostName)
            .Data("/nodes/pve1/qemu/100/agent/network-get-interfaces", AgentNetwork)
            .Data("/nodes/pve1/qemu/100/agent/get-fsinfo", AgentFsInfo)
            .Data("/nodes/pve2/lxc/101/config", Ct101Config)
            .Data("/nodes/pve2/lxc/101/status/current", Ct101Status)
            .Data("/nodes/pve2/lxc/101/snapshot", Ct101Snapshots)
            .Data("/nodes/pve2/lxc/101/interfaces", Ct101Interfaces)
            .Data("/nodes/pve2/qemu/9000/config", Tpl9000Config);
        return h;
    }
}
