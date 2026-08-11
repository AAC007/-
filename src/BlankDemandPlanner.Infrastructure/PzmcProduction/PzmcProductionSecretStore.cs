using System.Security.Cryptography;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;

namespace BlankDemandPlanner.Infrastructure.PzmcProduction;

[SupportedOSPlatform("windows")]
internal static class PzmcProductionSecretStore
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("BlankDemandPlanner.PzmcProduction.v1");

    public static IReadOnlyDictionary<string, string> Read(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var encrypted = File.ReadAllBytes(path);
            var clear = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
            return JsonSerializer.Deserialize<Dictionary<string, string>>(clear)
                   ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        catch (CryptographicException)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        catch (JsonException)
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
