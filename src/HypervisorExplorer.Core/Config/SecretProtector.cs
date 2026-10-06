using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;

namespace HypervisorExplorer.Core.Config;

/// <summary>Encrypts stored passwords / token secrets at rest.</summary>
public interface ISecretProtector
{
    string Protect(string plaintext);
    /// <summary>Returns null when the value cannot be decrypted (different user/machine, corrupted).</summary>
    string? Unprotect(string protectedValue);
}

/// <summary>
/// Windows: DPAPI bound to the current user (like the legacy PowerShell tool).
/// Elsewhere: AES-GCM with a per-user random key file readable only by the owner.
/// </summary>
public sealed class SecretProtector : ISecretProtector
{
    private const string DpapiPrefix = "dpapi:";
    private const string AesPrefix = "aes:";
    private static readonly byte[] Entropy = "HypervisorExplorer.v3"u8.ToArray();
    private readonly string _keyPath;
    private byte[]? _key;

    public SecretProtector(string configDirectory)
    {
        _keyPath = Path.Combine(configDirectory, "secret.key");
    }

    public string Protect(string plaintext)
    {
        if (OperatingSystem.IsWindows()) return DpapiPrefix + ProtectDpapi(plaintext);
        return AesPrefix + ProtectAes(plaintext);
    }

    public string? Unprotect(string protectedValue)
    {
        try
        {
            if (protectedValue.StartsWith(DpapiPrefix, StringComparison.Ordinal))
                return OperatingSystem.IsWindows() ? UnprotectDpapi(protectedValue[DpapiPrefix.Length..]) : null;
            if (protectedValue.StartsWith(AesPrefix, StringComparison.Ordinal))
                return UnprotectAes(protectedValue[AesPrefix.Length..]);
            return null;
        }
        catch (CryptographicException)
        {
            return null;
        }
        catch (FormatException)
        {
            return null;
        }
    }

    [SupportedOSPlatform("windows")]
    private static string ProtectDpapi(string plaintext) =>
        Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(plaintext), Entropy, DataProtectionScope.CurrentUser));

    [SupportedOSPlatform("windows")]
    private static string UnprotectDpapi(string b64) =>
        Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(b64), Entropy, DataProtectionScope.CurrentUser));

    private string ProtectAes(string plaintext)
    {
        var key = GetOrCreateKey();
        var nonce = RandomNumberGenerator.GetBytes(12);
        var plain = Encoding.UTF8.GetBytes(plaintext);
        var cipher = new byte[plain.Length];
        var tag = new byte[16];
        using var aes = new AesGcm(key, 16);
        aes.Encrypt(nonce, plain, cipher, tag);
        return Convert.ToBase64String([.. nonce, .. tag, .. cipher]);
    }

    private string? UnprotectAes(string b64)
    {
        if (!File.Exists(_keyPath)) return null;
        var key = GetOrCreateKey();
        var data = Convert.FromBase64String(b64);
        if (data.Length < 28) return null;
        var nonce = data.AsSpan(0, 12);
        var tag = data.AsSpan(12, 16);
        var cipher = data.AsSpan(28);
        var plain = new byte[cipher.Length];
        using var aes = new AesGcm(key, 16);
        aes.Decrypt(nonce, cipher, tag, plain);
        return Encoding.UTF8.GetString(plain);
    }

    private byte[] GetOrCreateKey()
    {
        if (_key is not null) return _key;
        if (File.Exists(_keyPath))
        {
            _key = Convert.FromBase64String(File.ReadAllText(_keyPath).Trim());
            return _key;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(_keyPath)!);
        _key = RandomNumberGenerator.GetBytes(32);
        File.WriteAllText(_keyPath, Convert.ToBase64String(_key));
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(_keyPath, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        return _key;
    }
}
