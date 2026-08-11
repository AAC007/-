namespace BlankDemandPlanner.Infrastructure.PzmcProduction;

internal sealed record PzmcProductionOptions(
    string BaseUrl,
    string ApiToken,
    string UserToken,
    string Login,
    string Password,
    int TimeoutSeconds,
    IReadOnlyList<string> CmoGroups,
    string CachePath,
    string PersonnelCachePath,
    int? MechanicalEntityId,
    int PersonnelRefreshMinutes)
{
    public bool IsConfigured =>
        Uri.TryCreate(BaseUrl, UriKind.Absolute, out _) &&
        !string.IsNullOrWhiteSpace(ApiToken) &&
        (!string.IsNullOrWhiteSpace(UserToken) ||
         (!string.IsNullOrWhiteSpace(Login) && !string.IsNullOrWhiteSpace(Password)));

    public static PzmcProductionOptions Load()
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var explicitPath = Environment.GetEnvironmentVariable("PZMC_PRODUCTION_ENV_FILE");
        var defaultPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BlankDemandPlanner",
            "PzmcProduction",
            ".env");
        var envPath = string.IsNullOrWhiteSpace(explicitPath) ? defaultPath : explicitPath;
        if (File.Exists(envPath))
        {
            foreach (var rawLine in File.ReadLines(envPath))
            {
                var line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith('#'))
                {
                    continue;
                }

                var separator = line.IndexOf('=');
                if (separator <= 0)
                {
                    continue;
                }

                values[line[..separator].Trim()] = Unquote(line[(separator + 1)..].Trim());
            }
        }

        string Read(string key, string fallback = "")
        {
            var environmentValue = Environment.GetEnvironmentVariable(key);
            return !string.IsNullOrWhiteSpace(environmentValue)
                ? environmentValue
                : values.GetValueOrDefault(key, fallback);
        }

        var dataDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BlankDemandPlanner",
            "PzmcProduction");
        Directory.CreateDirectory(dataDirectory);
        var secretFile = Environment.ExpandEnvironmentVariables(
            Read("PZMC_PRODUCTION_SECRET_FILE", Path.Combine(dataDirectory, "secrets.dat")));
        var secrets = OperatingSystem.IsWindows()
            ? PzmcProductionSecretStore.Read(secretFile)
            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        return new PzmcProductionOptions(
            Read("PZMC_PRODUCTION_API_URL", "http://192.168.88.179:82/"),
            Read("PZMC_PRODUCTION_API_TOKEN", secrets.GetValueOrDefault("api_token", string.Empty)),
            Read("PZMC_PRODUCTION_TOKEN", secrets.GetValueOrDefault("user_token", string.Empty)),
            Read("PZMC_PRODUCTION_LOGIN"),
            Read("PZMC_PRODUCTION_PASSWORD", secrets.GetValueOrDefault("password", string.Empty)),
            int.TryParse(Read("PZMC_PRODUCTION_TIMEOUT_SECONDS", "30"), out var timeout) ? Math.Clamp(timeout, 5, 180) : 30,
            Read("PZMC_PRODUCTION_CMO_GROUPS")
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
            Path.Combine(dataDirectory, "need-snapshot.json"),
            Path.Combine(dataDirectory, "personnel-snapshot.json"),
            int.TryParse(Read("PZMC_PRODUCTION_MECHANICAL_ENTITY_ID"), out var entityId) && entityId > 0 ? entityId : null,
            int.TryParse(Read("PZMC_PRODUCTION_PERSONNEL_REFRESH_MINUTES", "10"), out var refresh)
                ? Math.Clamp(refresh, 5, 15)
                : 10);
    }

    private static string Unquote(string value) =>
        value.Length >= 2 && ((value[0] == '"' && value[^1] == '"') || (value[0] == '\'' && value[^1] == '\''))
            ? value[1..^1]
            : value;
}
