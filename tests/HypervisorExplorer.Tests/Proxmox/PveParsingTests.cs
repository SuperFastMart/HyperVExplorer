using HypervisorExplorer.Collectors.Proxmox;
using HypervisorExplorer.Core.Model;

namespace HypervisorExplorer.Tests.Proxmox;

public class PveParsingTests
{
    [Theory]
    [InlineData("32G", 34359738368L)]
    [InlineData("500M", 524288000L)]
    [InlineData("1T", 1099511627776L)]
    [InlineData("10737418240", 10737418240L)]
    [InlineData("4M", 4194304L)]
    [InlineData("512K", 524288L)]
    [InlineData("1.5G", 1610612736L)]
    [InlineData("629M", 659554304L)]
    public void ParseSize_HandlesUnitsAndBareBytes(string input, long expected) =>
        Assert.Equal(expected, PveParsing.ParseSize(input));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData("G")]
    public void ParseSize_ReturnsNullForGarbage(string? input) => Assert.Null(PveParsing.ParseSize(input));

    [Fact]
    public void PropertyString_SplitsPositionalAndPairs()
    {
        var p = PveProps.Parse("local-lvm:vm-100-disk-0,iothread=1,discard=on,size=32G");
        Assert.Equal("local-lvm:vm-100-disk-0", p.Positional);
        Assert.Equal("32G", p["size"]);
        Assert.Equal("on", p["discard"]);
        Assert.Equal(3, p.Pairs.Count);

        var nic = PveProps.Parse("virtio=BC:24:11:AA:BB:01,bridge=vmbr0,tag=20");
        Assert.Null(nic.Positional);
        Assert.Equal("BC:24:11:AA:BB:01", nic["virtio"]);
        Assert.Equal("20", nic["tag"]);
    }

    [Theory]
    [InlineData("local-lvm:vm-100-disk-0", "local-lvm", "vm-100-disk-0")]
    [InlineData("nfs-shared:100/vm-100-disk-0.qcow2", "nfs-shared", "100/vm-100-disk-0.qcow2")]
    [InlineData("/dev/disk/by-id/ata-SAMSUNG", null, "/dev/disk/by-id/ata-SAMSUNG")]
    [InlineData("none", null, "none")]
    public void SplitVolume(string volid, string? storage, string? volume)
    {
        var (s, v) = PveParsing.SplitVolume(volid);
        Assert.Equal(storage, s);
        Assert.Equal(volume, v);
    }

    [Theory]
    [InlineData("24", "255.255.255.0")]
    [InlineData("16", "255.255.0.0")]
    [InlineData("255.255.255.128", "255.255.255.128")]
    public void SubnetMask(string input, string expected) => Assert.Equal(expected, PveParsing.ToSubnetMask(input));

    [Theory]
    [InlineData("running", "running", PowerState.PoweredOn)]
    [InlineData("running", "paused", PowerState.Suspended)]
    [InlineData("stopped", null, PowerState.PoweredOff)]
    [InlineData("unknown", null, PowerState.Unknown)]
    public void PowerStateMapping(string status, string? qmp, PowerState expected) =>
        Assert.Equal(expected, PveParsing.MapPowerState(status, qmp));

    [Fact]
    public void Tags_SplitOnSemicolonCommaSpace() =>
        Assert.Equal(["web", "prod", "dmz"], PveParsing.SplitTags("web;prod,dmz web"));

    [Theory]
    [InlineData("root@pam!explorer", "secret", "root@pam!explorer", "secret")]
    [InlineData("PVEAPIToken=root@pam!explorer=abc-123", "", "root@pam!explorer", "abc-123")]
    [InlineData("root!explorer", "s", "root@pam!explorer", "s")]
    public void TokenNormalization(string user, string secret, string expectedId, string expectedSecret)
    {
        var (id, sec) = PveApiClient.NormalizeToken(user, secret);
        Assert.Equal(expectedId, id);
        Assert.Equal(expectedSecret, sec);
    }

    [Theory]
    [InlineData("pve1", null, "https://pve1:8006/api2/json")]
    [InlineData("10.0.0.5", 8443, "https://10.0.0.5:8443/api2/json")]
    [InlineData("https://pve1.lab.local:8006/", null, "https://pve1.lab.local:8006/api2/json")]
    [InlineData("pve1:9006", null, "https://pve1:9006/api2/json")]
    [InlineData("fe80::1", null, "https://[fe80::1]:8006/api2/json")]
    public void BaseUri(string address, int? port, string expected) =>
        Assert.Equal(expected, PveApiClient.BuildBaseUri(address, port).ToString());

    [Fact]
    public void QemuOsNames()
    {
        Assert.Equal("Linux 2.6+ kernel", PveParsing.QemuOsName("l26"));
        Assert.Equal("Microsoft Windows 11/2022/2025", PveParsing.QemuOsName("win11"));
        Assert.Null(PveParsing.QemuOsName(null));
    }
}
