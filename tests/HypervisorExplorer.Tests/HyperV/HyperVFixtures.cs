namespace HypervisorExplorer.Tests.HyperV;

/// <summary>Realistic Collect-HyperV.ps1 / launcher output, as Windows PowerShell 5.1 would emit it.</summary>
internal static class HyperVFixtures
{
    /// <summary>Launcher "data" for a 2-node Failover Cluster (HV01 primary with cluster data, HV02 expanded).</summary>
    public const string ClusterData = """
    {
      "schemaVersion": 1, "kind": "multi", "requestedComputer": "hvclu01.contoso.local", "primary": "HV01",
      "warnings": ["Cluster node HV03 is Down; not collected."],
      "nodes": [
        {
          "schemaVersion": 1, "generator": "Collect-HyperV.ps1", "scriptVersion": "1.0.0",
          "collectedAt": "2026-10-06T10:00:00.0000000+01:00", "computerName": "HV01", "psVersion": "5.1.20348.2849",
          "host": {
            "name": "HV01", "dnsHostName": "hv01", "domain": "contoso.local", "partOfDomain": true,
            "manufacturer": "Dell Inc.", "model": "PowerEdge R750", "numberOfProcessors": 2, "totalPhysicalMemory": 549547945984,
            "logicalProcessorCount": 64, "memoryCapacity": 549755813888,
            "virtualHardDiskPath": "C:\\ClusterStorage\\Volume1\\Hyper-V\\Virtual Hard Disks",
            "virtualMachinePath": "C:\\ClusterStorage\\Volume1\\Hyper-V", "numaSpanningEnabled": true,
            "enableEnhancedSessionMode": false, "liveMigrationEnabled": true, "maximumLiveMigrations": 2,
            "processors": [
              { "socket": "CPU1", "name": "Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz", "manufacturer": "GenuineIntel", "maxClockSpeed": 2000, "numberOfCores": 16, "numberOfLogicalProcessors": 32, "loadPercentage": 10 },
              { "socket": "CPU2", "name": "Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz", "manufacturer": "GenuineIntel", "maxClockSpeed": 2000, "numberOfCores": 16, "numberOfLogicalProcessors": 32, "loadPercentage": 20 }
            ],
            "os": { "caption": "Microsoft Windows Server 2022 Datacenter", "version": "10.0.20348", "buildNumber": "20348",
                    "lastBootUpTime": "2026-09-01T03:15:00.0000000+01:00", "installDate": "2024-01-10T12:00:00.0000000+00:00",
                    "freePhysicalMemoryKb": 268435456, "totalVisibleMemoryKb": 536870912 },
            "bios": { "manufacturer": "Dell Inc.", "version": "1.10.2", "releaseDate": "2025-03-14T00:00:00.0000000+00:00", "serialNumber": "ABC1234" },
            "uuid": "4C4C4544-0042-3510-8051-B4C04F333332",
            "timeZone": { "id": "GMT Standard Time", "displayName": "(UTC+00:00) Dublin, Edinburgh, Lisbon, London", "baseUtcOffsetMinutes": 0 },
            "dnsServers": ["10.0.0.10", "10.0.0.11"], "dnsSearch": ["contoso.local"],
            "ntpSource": "dc01.contoso.local",
            "nics": [
              { "name": "SLOT 3 Port 1", "interfaceDescription": "Mellanox ConnectX-5 Adapter", "driverProvider": "Mellanox", "driverDescription": "Mellanox ConnectX-5 Adapter", "driverFileName": "mlx5.sys", "driverVersion": "3.0.25668.0", "linkSpeed": "25 Gbps", "speedBps": 25000000000, "fullDuplex": true, "macAddress": "B8-59-9F-00-00-01", "status": "Up", "mtu": 1514, "ifIndex": 5, "pci": "0000:3b:00.0" },
              { "name": "SLOT 3 Port 2", "interfaceDescription": "Mellanox ConnectX-5 Adapter #2", "driverProvider": "Mellanox", "driverDescription": "Mellanox ConnectX-5 Adapter", "driverFileName": "mlx5.sys", "driverVersion": "3.0.25668.0", "linkSpeed": "25 Gbps", "speedBps": 25000000000, "fullDuplex": true, "macAddress": "B8-59-9F-00-00-02", "status": "Up", "mtu": 1514, "ifIndex": 6, "pci": "0000:3b:00.1" },
              { "name": "NIC1", "interfaceDescription": "Broadcom NetXtreme Gigabit Ethernet", "driverFileName": "b57nd60a.sys", "speedBps": 0, "fullDuplex": true, "macAddress": "B0-26-28-00-00-01", "status": "Disconnected", "ifIndex": 7 }
            ],
            "lbfoTeams": [],
            "switches": [
              { "name": "SETswitch", "id": "0a1b2c3d-0000-0000-0000-000000000001", "switchType": "External",
                "netAdapterInterfaceDescriptions": ["Mellanox ConnectX-5 Adapter", "Mellanox ConnectX-5 Adapter #2"],
                "embeddedTeamingEnabled": true, "allowManagementOS": true, "bandwidthReservationMode": "Weight", "notes": null }
            ],
            "managementAdapters": [
              { "name": "Management", "switchName": "SETswitch", "macAddress": "00155D0A0001", "status": "Ok", "vlanMode": "Access", "accessVlanId": 10 },
              { "name": "LiveMigration", "switchName": "SETswitch", "macAddress": "00155D0A0002", "status": "Ok", "vlanMode": "Access", "accessVlanId": 30 }
            ],
            "ipInterfaces": [
              { "interfaceAlias": "vEthernet (Management)", "ifIndex": 12, "interfaceDescription": "Hyper-V Virtual Ethernet Adapter", "macAddress": "00-15-5D-0A-00-01",
                "dhcp": false, "mtu": 1500, "gateway": "10.0.10.1",
                "ipv4": [{ "address": "10.0.10.21", "prefixLength": 24, "origin": "Manual" }], "ipv6": [] },
              { "interfaceAlias": "vEthernet (LiveMigration)", "ifIndex": 13, "interfaceDescription": "Hyper-V Virtual Ethernet Adapter #2", "macAddress": "00-15-5D-0A-00-02",
                "dhcp": false, "mtu": 9000, "gateway": null,
                "ipv4": [{ "address": "10.0.30.21", "prefixLength": 24, "origin": "Manual" }], "ipv6": [] }
            ],
            "initiatorPorts": [
              { "instanceName": "PCI\\VEN_10DF&DEV_F400\\4&1a2b3c4d&0&0010_0", "nodeAddress": "20000090FA000001", "portAddress": "10000090FA000001", "connectionType": "Fibre Channel", "operationalStatus": "Operational" }
            ],
            "scsiControllers": [
              { "name": "Dell HBA355i Front", "driverName": "SmartPqi", "manufacturer": "Dell", "status": "OK", "pnpDeviceId": "PCI\\VEN_9005" }
            ],
            "volumes": [
              { "label": null, "fileSystem": "NTFS", "driveLetter": "C", "size": 479559348224, "sizeRemaining": 400000000000, "healthStatus": "Healthy", "paths": ["C:\\"] },
              { "label": "Data", "fileSystem": "ReFS", "driveLetter": "D", "size": 1999844147200, "sizeRemaining": 1500000000000, "healthStatus": "Healthy", "paths": ["D:\\"] }
            ]
          },
          "cluster": {
            "name": "HVCLU01", "id": "a1b2c3d4-1111-2222-3333-444455556666", "domain": "contoso.local", "functionalLevel": 11, "s2dEnabled": 0,
            "quorumType": "NodeAndFileShareMajority", "quorumWitness": "File Share Witness", "quorumWitnessType": "File Share Witness",
            "nodes": [
              { "name": "HV01", "id": "1", "state": "Up", "drainStatus": "NotInitiated", "nodeWeight": 1, "dynamicWeight": 1, "manufacturer": "Dell Inc.", "model": "PowerEdge R750", "serialNumber": "ABC1234" },
              { "name": "HV02", "id": "2", "state": "Up", "drainStatus": "NotInitiated", "nodeWeight": 1, "dynamicWeight": 1 },
              { "name": "HV03", "id": "3", "state": "Down", "drainStatus": "NotInitiated", "nodeWeight": 1, "dynamicWeight": 0 }
            ],
            "networkInterfaces": [
              { "name": "HV01 - vEthernet (Management)", "node": "HV01", "network": "Management", "address": "10.0.10.21", "state": "Up" },
              { "name": "HV01 - vEthernet (LiveMigration)", "node": "HV01", "network": "LiveMigration", "address": "10.0.30.21", "state": "Up" },
              { "name": "HV02 - vEthernet (LiveMigration)", "node": "HV02", "network": "LiveMigration", "address": "10.0.30.22", "state": "Up" },
              { "name": "HV02 - vEthernet (Management)", "node": "HV02", "network": "Management", "address": "10.0.10.22", "state": "Up" }
            ],
            "networks": [
              { "name": "Management", "id": "n1", "role": "ClusterAndClient", "roleValue": 3, "address": "10.0.10.0", "addressMask": "255.255.255.0", "ipv4PrefixLengths": ["24"], "state": "Up", "metric": 70384, "autoMetric": true },
              { "name": "LiveMigration", "id": "n2", "role": "Cluster", "roleValue": 1, "address": "10.0.30.0", "addressMask": "255.255.255.0", "ipv4PrefixLengths": ["24"], "state": "Up", "metric": 39840, "autoMetric": true },
              { "name": "iSCSI", "id": "n3", "role": "0", "address": "10.0.40.0", "addressMask": "255.255.255.0", "state": "Up", "metric": 79999, "autoMetric": false }
            ],
            "sharedVolumes": [
              { "name": "Cluster Disk 1", "id": "c1", "state": "Online", "ownerNode": "HV01", "friendlyVolumeName": "C:\\ClusterStorage\\Volume1",
                "fileSystem": "NTFS", "size": 4398046511104, "freeSpace": 2199023255552, "maintenanceMode": false, "redirectedAccess": false, "faultState": "NoFaults" },
              { "name": "Cluster Disk 2", "id": "c2", "state": "Online", "ownerNode": "HV02", "friendlyVolumeName": "C:\\ClusterStorage\\Volume10",
                "fileSystem": "ReFS", "size": 2199023255552, "freeSpace": 109951162777, "maintenanceMode": false, "redirectedAccess": true, "faultState": "NoFaults" }
            ],
            "vmGroups": [
              { "name": "SQL01", "id": "g1", "ownerNode": "HV01", "state": "Online", "priority": 3000, "preferredOwners": ["HV01", "HV02"],
                "autoFailbackType": 1, "failoverThreshold": 2, "failoverPeriod": 6, "vmId": "6f1d3a2e-8b4c-4d5e-9f00-112233445566" },
              { "name": "SCVMM APP01 Resources", "id": "g2", "ownerNode": "HV02", "state": "Online", "priority": 1000, "preferredOwners": [],
                "autoFailbackType": 0, "failoverThreshold": 1, "failoverPeriod": 6, "vmId": "{9A8B7C6D-0000-4000-8000-ABCDEF012345}" },
              { "name": "OLDVM", "id": "g3", "ownerNode": "HV03", "state": "Failed", "priority": 0, "preferredOwners": ["HV03"], "vmId": null }
            ]
          },
          "vms": [
            {
              "name": "SQL01", "id": "6f1d3a2e-8b4c-4d5e-9f00-112233445566", "state": "Running", "status": "Operating normally",
              "uptimeSeconds": 86400, "generation": 2, "version": "10.0", "processorCount": 8,
              "memoryStartup": 17179869184, "memoryAssigned": 21474836480, "memoryDemand": 19327352832,
              "memoryMinimum": 8589934592, "memoryMaximum": 34359738368, "dynamicMemoryEnabled": true,
              "path": "C:\\ClusterStorage\\Volume1\\SQL01", "configurationLocation": "C:\\ClusterStorage\\Volume1\\SQL01",
              "snapshotFileLocation": "C:\\ClusterStorage\\Volume1\\SQL01\\Snapshots", "smartPagingFilePath": "C:\\ClusterStorage\\Volume1\\SQL01",
              "notes": "Production SQL", "creationTime": "2024-02-01T09:30:00.0000000+00:00",
              "automaticStartAction": "StartIfRunning", "automaticStartDelay": 120, "automaticStopAction": "ShutDown",
              "isClustered": true, "replicationState": "Disabled", "replicationHealth": "NotApplicable", "checkpointType": "Production",
              "integrationServicesVersion": "10.0.20348", "integrationServicesState": "Up to date",
              "heartbeat": "OkApplicationsUnknown", "heartbeatStatus": "Ok",
              "integrationServices": [
                { "name": "Guest Service Interface", "enabled": false, "primaryOperationalStatus": null, "primaryStatusDescription": null },
                { "name": "Heartbeat", "enabled": true, "primaryOperationalStatus": "Ok", "primaryStatusDescription": "OK" },
                { "name": "Key-Value Pair Exchange", "enabled": true, "primaryOperationalStatus": "Ok", "primaryStatusDescription": "OK" },
                { "name": "Shutdown", "enabled": true, "primaryOperationalStatus": "Ok", "primaryStatusDescription": "OK" },
                { "name": "Time Synchronization", "enabled": true, "primaryOperationalStatus": "Ok", "primaryStatusDescription": "OK" },
                { "name": "VSS", "enabled": true, "primaryOperationalStatus": "Ok", "primaryStatusDescription": "OK" }
              ],
              "processor": { "count": 8, "reserve": 0, "maximum": 100, "relativeWeight": 100, "exposeVirtualizationExtensions": false, "compatibilityForMigrationEnabled": true },
              "firmware": { "secureBoot": true, "secureBootTemplate": "MicrosoftWindows", "bootOrder": ["HardDiskDrive", "DvdDrive", "VMNetworkAdapter", "File:Windows Boot Manager"] },
              "bios": null,
              "disks": [
                { "name": "Hard Drive", "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 0,
                  "path": "C:\\ClusterStorage\\Volume1\\SQL01\\Virtual Hard Disks\\SQL01_OS.vhdx", "diskNumber": null, "supportPersistentReservations": false,
                  "vhdFormat": "VHDX", "vhdType": "Dynamic", "size": 136365211648, "fileSize": 42949672960, "parentPath": null },
                { "name": "Hard Drive", "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 1,
                  "path": "c:\\clusterstorage\\volume10\\SQL01\\SQL01_Data_8C1F.avhdx", "diskNumber": null, "supportPersistentReservations": false,
                  "vhdFormat": "VHDX", "vhdType": "Differencing", "size": 536870912000, "fileSize": 1073741824,
                  "parentPath": "C:\\ClusterStorage\\Volume10\\SQL01\\SQL01_Data.vhdx" }
              ],
              "dvdDrives": [
                { "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 2, "path": "C:\\ClusterStorage\\Volume1\\ISO\\SQLServer2022.iso", "dvdMediaType": "ISO" }
              ],
              "networkAdapters": [
                { "name": "Network Adapter", "id": "Microsoft:6F1D3A2E\\A1", "switchName": "SETswitch", "macAddress": "00155D0A1001",
                  "ipAddresses": ["10.0.20.50", "fe80::1c2d:3e4f:5a6b:7c8d", "2001:db8:20::50"], "status": "Ok", "connected": true,
                  "dynamicMacAddressEnabled": false, "isLegacy": false, "vlanMode": "Access", "accessVlanId": 20, "nativeVlanId": 0, "allowedVlanIds": null }
              ],
              "snapshots": [
                { "name": "Before CU12", "id": "s1", "creationTime": "2026-09-20T22:00:00.0000000+01:00", "snapshotType": "Production",
                  "parentSnapshotName": null, "path": "C:\\ClusterStorage\\Volume1\\SQL01\\Snapshots" }
              ],
              "guest": { "osName": "Windows Server 2022 Datacenter", "osVersion": "10.0.20348", "osBuildNumber": "20348",
                         "fullyQualifiedDomainName": "sql01.contoso.local", "integrationServicesVersion": "10.0.20348.2849" },
              "biosGuid": "2b9f1e4a-6c7d-4e8f-9a0b-1c2d3e4f5a6b"
            },
            {
              "name": "WEB01", "id": "11111111-2222-3333-4444-555555555555", "state": "Off", "status": "Operating normally",
              "uptimeSeconds": 0, "generation": 1, "version": "9.0", "processorCount": 2,
              "memoryStartup": 4294967296, "memoryAssigned": 0, "memoryDemand": 0, "memoryMinimum": 536870912, "memoryMaximum": 1099511627776,
              "dynamicMemoryEnabled": false, "path": "D:\\Hyper-V\\WEB01", "configurationLocation": "D:\\Hyper-V\\WEB01",
              "snapshotFileLocation": "D:\\Hyper-V\\WEB01", "notes": null, "creationTime": "2023-05-05T10:00:00.0000000+00:00",
              "automaticStartAction": "Nothing", "automaticStartDelay": 0, "automaticStopAction": "Save", "isClustered": false,
              "replicationState": "Disabled", "checkpointType": "Standard", "integrationServicesVersion": "0.0", "heartbeat": null, "heartbeatStatus": null,
              "integrationServices": [{ "name": "Heartbeat", "enabled": true, "primaryOperationalStatus": null, "primaryStatusDescription": null }],
              "processor": { "count": 2, "reserve": 10, "maximum": 75, "relativeWeight": 200 },
              "firmware": null, "bios": { "startupOrder": ["CD", "IDE", "LegacyNetworkAdapter", "Floppy"] },
              "disks": [
                { "name": "Hard Drive", "controllerType": "IDE", "controllerNumber": 0, "controllerLocation": 0, "path": "D:\\Hyper-V\\WEB01\\WEB01.vhd",
                  "diskNumber": null, "supportPersistentReservations": false, "vhdFormat": "VHD", "vhdType": "Fixed", "size": 42949672960, "fileSize": 42949673472, "parentPath": null },
                { "name": "Hard Drive", "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 0, "path": "Disk 4 500.00 GB Bus 0 Lun 4 Target 0",
                  "diskNumber": 4, "supportPersistentReservations": false, "vhdFormat": null, "vhdType": null, "size": null, "fileSize": null, "parentPath": null }
              ],
              "dvdDrives": [{ "controllerType": "IDE", "controllerNumber": 1, "controllerLocation": 0, "path": null, "dvdMediaType": "None" }],
              "networkAdapters": [
                { "name": "Legacy Network Adapter", "switchName": null, "macAddress": "000000000000", "ipAddresses": [], "status": null, "connected": false,
                  "dynamicMacAddressEnabled": true, "isLegacy": true, "vlanMode": "Untagged", "accessVlanId": 0 }
              ],
              "snapshots": [], "guest": null, "biosGuid": null
            }
          ],
          "warnings": ["HV01: VM 'WEB01' integration services: Sample warning"]
        },
        {
          "schemaVersion": 1, "generator": "Collect-HyperV.ps1", "scriptVersion": "1.0.0",
          "collectedAt": "2026-10-06T10:00:05.0000000+01:00", "computerName": "HV02",
          "host": {
            "name": "HV02", "domain": "contoso.local", "manufacturer": "Dell Inc.", "model": "PowerEdge R750",
            "logicalProcessorCount": 48, "memoryCapacity": 412316860416,
            "processors": [
              { "name": "Intel(R) Xeon(R) Gold 5318Y CPU @ 2.10GHz", "maxClockSpeed": 2100, "numberOfCores": 24, "numberOfLogicalProcessors": 24, "loadPercentage": 50 }
            ],
            "os": { "caption": "Microsoft Windows Server 2022 Datacenter", "version": "10.0.20348", "buildNumber": "20348",
                    "lastBootUpTime": "2026-09-02T03:15:00.0000000+01:00", "freePhysicalMemoryKb": 100000000, "totalVisibleMemoryKb": 402653184 },
            "bios": { "manufacturer": "Dell Inc.", "version": "1.10.2", "serialNumber": "XYZ9876" },
            "uuid": "4C4C4544-0042-3510-8051-B4C04F333333",
            "dnsServers": ["10.0.0.10"], "ntpSource": "dc01.contoso.local,0x9",
            "nics": [], "switches": [], "managementAdapters": [], "ipInterfaces": [], "initiatorPorts": [], "scsiControllers": [],
            "volumes": [ { "label": null, "fileSystem": "NTFS", "driveLetter": "C", "size": 479559348224, "sizeRemaining": 20000000000, "healthStatus": "Warning", "paths": ["C:\\"] } ]
          },
          "cluster": null,
          "vms": [
            {
              "name": "APP01", "id": "9a8b7c6d-0000-4000-8000-abcdef012345", "state": "Running", "status": "Operating normally",
              "uptimeSeconds": 3600, "generation": 2, "version": "10.0", "processorCount": 4,
              "memoryStartup": 8589934592, "memoryAssigned": 8589934592, "memoryDemand": 4294967296, "dynamicMemoryEnabled": false,
              "path": "C:\\ClusterStorage\\Volume1\\APP01", "configurationLocation": "C:\\ClusterStorage\\Volume1\\APP01",
              "isClustered": true, "automaticStartAction": "Start", "automaticStartDelay": 0,
              "integrationServicesVersion": "10.0.17763", "heartbeat": "LostCommunication",
              "integrationServices": [],
              "firmware": { "secureBoot": false, "bootOrder": ["VMNetworkAdapter"] },
              "disks": [
                { "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 0, "path": "\\\\fs01\\vms\\APP01\\APP01.vhdx",
                  "vhdFormat": "VHDX", "vhdType": "Fixed", "size": 107374182400, "fileSize": 107374182400 },
                { "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 1, "path": "\\\\fs01\\vms\\APP01\\shared.vhds",
                  "vhdFormat": "VHDSet", "vhdType": "Dynamic", "size": 10737418240, "fileSize": 4194304, "supportPersistentReservations": true }
              ],
              "dvdDrives": [],
              "networkAdapters": {
                "name": "Network Adapter", "switchName": "SETswitch", "macAddress": "00155D0A2001", "ipAddresses": "10.0.20.60",
                "connected": true, "dynamicMacAddressEnabled": true, "isLegacy": false, "vlanMode": "Trunk", "accessVlanId": 0,
                "nativeVlanId": 1, "allowedVlanIds": "20-25"
              },
              "snapshots": []
            }
          ],
          "warnings": []
        }
      ]
    }
    """;

    /// <summary>A manual run of Collect-HyperV.ps1 on a standalone host (single-host document).</summary>
    public const string StandaloneHost = """
    {
      "schemaVersion": 1, "generator": "Collect-HyperV.ps1", "scriptVersion": "1.0.0",
      "collectedAt": "2026-10-05T08:00:00.0000000+00:00", "computerName": "HVSTD01", "psVersion": "5.1.17763.1",
      "host": {
        "name": "HVSTD01", "domain": "WORKGROUP", "partOfDomain": false, "manufacturer": "HPE", "model": "ProLiant DL360 Gen10",
        "logicalProcessorCount": 16, "memoryCapacity": 137438953472,
        "processors": { "name": "Intel(R) Xeon(R) Silver 4110 CPU @ 2.10GHz", "maxClockSpeed": 2100, "numberOfCores": 8, "numberOfLogicalProcessors": 16, "loadPercentage": 5 },
        "os": { "caption": "Microsoft Windows Server 2019 Standard", "version": "10.0.17763", "buildNumber": "17763",
                "lastBootUpTime": "2026-08-15T06:00:00.0000000+00:00", "freePhysicalMemoryKb": 67108864, "totalVisibleMemoryKb": 134217728 },
        "bios": { "manufacturer": "HPE", "version": "U32", "releaseDate": "2024-11-01T00:00:00.0000000+00:00", "serialNumber": "CZJ0000001" },
        "uuid": "30373237-3132-5A43-4A30-303030303031",
        "timeZone": { "id": "W. Europe Standard Time", "displayName": "(UTC+01:00) Amsterdam, Berlin", "baseUtcOffsetMinutes": 60 },
        "dnsServers": ["192.168.1.1"], "ntpSource": "Local CMOS Clock",
        "nics": [ { "name": "Embedded LOM 1 Port 1", "interfaceDescription": "HPE Ethernet 1Gb 4-port 331i Adapter", "driverFileName": "b57nd60a.sys",
                    "speedBps": 1000000000, "fullDuplex": true, "macAddress": "94-18-82-00-00-01", "status": "Up" } ],
        "lbfoTeams": [],
        "switches": [ { "name": "External", "switchType": "External", "netAdapterInterfaceDescriptions": "HPE Ethernet 1Gb 4-port 331i Adapter", "embeddedTeamingEnabled": false, "allowManagementOS": true } ],
        "managementAdapters": [ { "name": "External", "switchName": "External", "macAddress": "00155D010101", "vlanMode": "Untagged", "accessVlanId": 0 } ],
        "ipInterfaces": [ { "interfaceAlias": "vEthernet (External)", "macAddress": "00-15-5D-01-01-01", "dhcp": true, "mtu": 1500, "gateway": "192.168.1.1",
                            "ipv4": { "address": "192.168.1.20", "prefixLength": 24 }, "ipv6": [] } ],
        "volumes": [ { "label": "OS", "fileSystem": "NTFS", "driveLetter": "C", "size": 999653638144, "sizeRemaining": 49982681907, "healthStatus": "Healthy", "paths": ["C:\\"] } ]
      },
      "cluster": null,
      "vms": [
        {
          "name": "DC01", "id": "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", "state": "Saved", "status": "Operating normally", "generation": 2,
          "version": "9.0", "processorCount": 2, "memoryStartup": 4294967296, "dynamicMemoryEnabled": false,
          "configurationLocation": "C:\\ProgramData\\Microsoft\\Windows\\Hyper-V", "isClustered": false,
          "automaticStartAction": "StartIfRunning", "automaticStartDelay": 30, "integrationServicesVersion": "10.0.17763",
          "firmware": { "secureBoot": true, "secureBootTemplate": "MicrosoftWindows", "bootOrder": ["HardDiskDrive"] },
          "disks": [ { "controllerType": "SCSI", "controllerNumber": 0, "controllerLocation": 0,
                       "path": "C:\\Users\\Public\\Documents\\Hyper-V\\Virtual Hard Disks\\DC01.vhdx", "vhdFormat": "VHDX", "vhdType": "Dynamic",
                       "size": 64424509440, "fileSize": 16106127360 } ],
          "networkAdapters": [ { "name": "Network Adapter", "switchName": "External", "macAddress": "00155D010102", "ipAddresses": [], "isLegacy": false } ],
          "snapshots": [ { "name": "DC01 - (01/09/2026 - 10:00:00)", "creationTime": "2026-09-01T10:00:00.0000000+00:00", "snapshotType": "Standard" },
                         { "name": "after patch", "creationTime": "2026-09-02T10:00:00.0000000+00:00", "snapshotType": "Standard", "parentSnapshotName": "DC01 - (01/09/2026 - 10:00:00)" } ]
        }
      ],
      "warnings": []
    }
    """;

    /// <summary>Wraps launcher data in a success envelope, as written to stdout.</summary>
    public static string Envelope(string data) =>
        "{\"ok\":true,\"data\":" + System.Text.Json.JsonSerializer.Serialize(System.Text.Json.JsonDocument.Parse(data).RootElement) + "}";
}
