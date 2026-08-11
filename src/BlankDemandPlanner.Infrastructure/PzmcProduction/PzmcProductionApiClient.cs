using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.PzmcProduction;

public sealed class PzmcProductionApiClient : IPzmcProductionApiClient, IDisposable
{
    private readonly PzmcProductionOptions options = PzmcProductionOptions.Load();
    private readonly HttpClient client;
    private readonly ILogger<PzmcProductionApiClient> logger;
    private readonly SemaphoreSlim authenticationLock = new(1, 1);
    private readonly JsonSerializerOptions jsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = true
    };
    private bool authenticated;

    public bool IsConfigured => options.IsConfigured;
    public IReadOnlyList<string> CmoGroups => options.CmoGroups;
    public int? MechanicalEntityId => options.MechanicalEntityId;
    public int PersonnelRefreshMinutes => options.PersonnelRefreshMinutes;

    public PzmcProductionApiClient(ILogger<PzmcProductionApiClient> logger)
    {
        this.logger = logger;
        var handler = new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };
        client = new HttpClient(handler)
        {
            BaseAddress = new Uri(options.BaseUrl, UriKind.Absolute),
            Timeout = TimeSpan.FromSeconds(options.TimeoutSeconds)
        };
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        client.DefaultRequestHeaders.AcceptEncoding.ParseAdd("gzip, deflate");
        client.DefaultRequestHeaders.TryAddWithoutValidation("client-name", Environment.MachineName);
        if (!string.IsNullOrWhiteSpace(options.ApiToken))
        {
            client.DefaultRequestHeaders.TryAddWithoutValidation("X-API-TOKEN", options.ApiToken);
        }
    }

    public async Task<PzmcProductionSnapshot> LoadSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken)
    {
        if (!forceRefresh && TryLoadCache(out var cached))
        {
            return cached;
        }

        try
        {
            if (!IsConfigured)
            {
                throw new InvalidOperationException(
                    "Интеграция с ПО «Производство» не настроена. Заполните внешний .env без размещения пароля в коде.");
            }

            await EnsureAuthenticatedAsync(cancellationToken);
            var needsTask = GetProductNeedsAsync(cancellationToken);
            var productsTask = GetProductsAsync(cancellationToken);
            var ipsTask = GetIpsObjectsAsync(cancellationToken);
            var skladsTask = GetWarehousesAsync(cancellationToken);
            var plansTask = GetPlannedReceiptsAsync(cancellationToken);
            var analogsTask = GetAnalogsAsync(cancellationToken);
            var groupsTask = GetProductGroupsAsync(cancellationToken);
            await Task.WhenAll(needsTask, productsTask, ipsTask, skladsTask, plansTask, analogsTask, groupsTask);

            var snapshot = new PzmcProductionSnapshot(
                await needsTask,
                await productsTask,
                await ipsTask,
                await skladsTask,
                await plansTask,
                await analogsTask,
                await groupsTask,
                DateTime.Now,
                "API ПО «Производство»");
            SaveCache(snapshot);
            return snapshot;
        }
        catch (Exception ex) when (TryLoadCache(out cached))
        {
            logger.LogWarning(ex, "PZMC Production API unavailable, cached need snapshot is used.");
            return cached with { Source = $"Локальный кэш от {cached.UpdatedAt:dd.MM.yyyy HH:mm}" };
        }
    }

    public async Task<PzmcProductionPersonnelSnapshot> LoadPersonnelSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken)
    {
        if (!forceRefresh && TryLoadPersonnelCache(out var cached))
        {
            return cached;
        }

        try
        {
            if (!IsConfigured)
            {
                throw new InvalidOperationException(
                    "Интеграция с ПО «Производство» не настроена. Заполните внешний .env без размещения пароля в коде.");
            }

            await EnsureAuthenticatedAsync(cancellationToken);
            var employeesTask = GetEmployeesAsync(cancellationToken);
            var statisticsTask = GetEmployeeStatisticsAsync(cancellationToken);
            var entitiesTask = GetEntitiesAsync(cancellationToken);
            await Task.WhenAll(employeesTask, statisticsTask, entitiesTask);
            var employees = await employeesTask;
            var skudLogs = await GetSkudLogUsersAsync(employees, cancellationToken);

            var snapshot = new PzmcProductionPersonnelSnapshot(
                employees,
                await statisticsTask,
                skudLogs,
                await entitiesTask,
                DateTime.Now,
                "API ПО «Производство»");
            SavePersonnelCache(snapshot);
            return snapshot;
        }
        catch (Exception ex) when (TryLoadPersonnelCache(out cached))
        {
            logger.LogWarning(ex, "PZMC Production API unavailable, cached personnel snapshot is used.");
            return cached with { Source = $"Локальный кэш персонала от {cached.UpdatedAt:dd.MM.yyyy HH:mm}" };
        }
    }

    public Task<IReadOnlyList<PzmcProductNeed>> GetProductNeedsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcProductNeed>("ProductNeed", cancellationToken);

    public Task<IReadOnlyList<PzmcProduct>> GetProductsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcProduct>("Product", cancellationToken);

    public Task<IReadOnlyList<PzmcSpecIpsObject>> GetIpsObjectsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcSpecIpsObject>("SpecIPSObject", cancellationToken);

    public Task<IReadOnlyList<PzmcSklad>> GetWarehousesAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcSklad>("Sklad", cancellationToken);

    public Task<IReadOnlyList<PzmcSkladPlanBase>> GetPlannedReceiptsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcSkladPlanBase>("SkladPlanBase", cancellationToken);

    public Task<IReadOnlyList<PzmcAnalog>> GetAnalogsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcAnalog>("Analog", cancellationToken);

    public Task<IReadOnlyList<PzmcProductGroup>> GetProductGroupsAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcProductGroup>("ProductGroup", cancellationToken);

    public Task<IReadOnlyList<PzmcEmployee>> GetEmployeesAsync(CancellationToken cancellationToken) =>
        GetListWithFallbackAsync<PzmcEmployee>(["Employee/dep/all", "Employee/all/now", "Employee"], cancellationToken);

    public Task<IReadOnlyList<PzmcEmployeeStatistic>> GetEmployeeStatisticsAsync(CancellationToken cancellationToken) =>
        GetListWithFallbackAsync<PzmcEmployeeStatistic>(
            [$"EmployeeStatistic/on_date/{DateTime.Today:yyyy-MM-dd}", "EmployeeStatistic"],
            cancellationToken);

    public Task<IReadOnlyList<PzmcSkudLogUser>> GetSkudLogUsersAsync(CancellationToken cancellationToken) =>
        GetListWithFallbackAsync<PzmcSkudLogUser>(
            [$"SkudLogUser/on_date/{DateTime.Today:yyyy-MM-dd}", "SkudLogUser/all/now", "SkudLogUser"],
            cancellationToken);

    public Task<IReadOnlyList<PzmcEntitie>> GetEntitiesAsync(CancellationToken cancellationToken) =>
        GetListAsync<PzmcEntitie>("Entitie", cancellationToken);

    private async Task EnsureAuthenticatedAsync(CancellationToken cancellationToken)
    {
        if (authenticated)
        {
            return;
        }

        await authenticationLock.WaitAsync(cancellationToken);
        try
        {
            if (authenticated)
            {
                return;
            }

            var token = options.UserToken;
            if (string.IsNullOrWhiteSpace(token))
            {
                using var response = await client.PostAsJsonAsync("auth/get_token", new
                {
                    login_name = options.Login,
                    login_password = options.Password
                }, jsonOptions, cancellationToken);
                await EnsureSuccessAsync(response, cancellationToken);
                var payload = await response.Content.ReadFromJsonAsync<PzmcTokenResponse>(jsonOptions, cancellationToken);
                token = payload?.TokenName ?? payload?.TokenApi ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Сервер «Производство» не вернул пользовательский токен.");
            }

            client.DefaultRequestHeaders.Remove("X-TOKEN");
            client.DefaultRequestHeaders.TryAddWithoutValidation("X-TOKEN", token);
            authenticated = true;
        }
        finally
        {
            authenticationLock.Release();
        }
    }

    private async Task<IReadOnlyList<T>> GetListAsync<T>(string endpoint, CancellationToken cancellationToken)
    {
        await EnsureAuthenticatedAsync(cancellationToken);
        using var response = await client.GetAsync(endpoint, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        await EnsureSuccessAsync(response, cancellationToken);
        return await response.Content.ReadFromJsonAsync<List<T>>(jsonOptions, cancellationToken) ?? [];
    }

    private async Task<IReadOnlyList<T>> GetListWithFallbackAsync<T>(IReadOnlyList<string> endpoints, CancellationToken cancellationToken)
    {
        Exception? lastError = null;
        foreach (var endpoint in endpoints)
        {
            try
            {
                return await GetListAsync<T>(endpoint, cancellationToken);
            }
            catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                lastError = ex;
            }
        }

        throw lastError ?? new InvalidOperationException("No API endpoint candidates were supplied.");
    }

    private async Task<IReadOnlyList<PzmcSkudLogUser>> GetSkudLogUsersAsync(
        IReadOnlyList<PzmcEmployee> employees,
        CancellationToken cancellationToken)
    {
        var month = DateTime.Today.ToString("yyyy-MM");
        var userIds = employees
            .Select(x => x.UserId > 0 ? x.UserId : PzmcDynamicJson.GetInt(x.Extra, "userId", "user_id") ?? 0)
            .Where(x => x > 0)
            .Distinct()
            .Take(250)
            .ToArray();
        var rows = new List<PzmcSkudLogUser>();
        foreach (var userId in userIds)
        {
            try
            {
                var logs = await GetListAsync<PzmcSkudLogUser>($"SkudLogUser/main_on_date/{month}/user_id/{userId}", cancellationToken);
                rows.AddRange(logs);
            }
            catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                logger.LogDebug(ex, "PZMC Production SKUD endpoint not found for user {UserId}.", userId);
                break;
            }
        }

        return rows;
    }

    private static async Task EnsureSuccessAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }

        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        var shortBody = body.Length > 500 ? body[..500] : body;
        throw new HttpRequestException(
            $"API «Производство» вернул {(int)response.StatusCode} {response.ReasonPhrase}: {shortBody}",
            null,
            response.StatusCode);
    }

    private bool TryLoadCache(out PzmcProductionSnapshot snapshot)
    {
        snapshot = null!;
        try
        {
            if (!File.Exists(options.CachePath))
            {
                return false;
            }

            snapshot = JsonSerializer.Deserialize<PzmcProductionSnapshot>(
                File.ReadAllText(options.CachePath),
                jsonOptions)!;
            return snapshot is not null;
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "PZMC Production need cache cannot be read.");
            return false;
        }
    }

    private void SaveCache(PzmcProductionSnapshot snapshot)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(options.CachePath)!);
            var temporaryPath = options.CachePath + ".tmp";
            File.WriteAllText(temporaryPath, JsonSerializer.Serialize(snapshot, jsonOptions));
            File.Move(temporaryPath, options.CachePath, true);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "PZMC Production need cache cannot be saved.");
        }
    }

    private bool TryLoadPersonnelCache(out PzmcProductionPersonnelSnapshot snapshot)
    {
        snapshot = null!;
        try
        {
            if (!File.Exists(options.PersonnelCachePath))
            {
                return false;
            }

            snapshot = JsonSerializer.Deserialize<PzmcProductionPersonnelSnapshot>(
                File.ReadAllText(options.PersonnelCachePath),
                jsonOptions)!;
            return snapshot is not null;
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "PZMC Production personnel cache cannot be read.");
            return false;
        }
    }

    private void SavePersonnelCache(PzmcProductionPersonnelSnapshot snapshot)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(options.PersonnelCachePath)!);
            var temporaryPath = options.PersonnelCachePath + ".tmp";
            File.WriteAllText(temporaryPath, JsonSerializer.Serialize(snapshot, jsonOptions));
            File.Move(temporaryPath, options.PersonnelCachePath, true);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "PZMC Production personnel cache cannot be saved.");
        }
    }

    public void Dispose()
    {
        client.Dispose();
        authenticationLock.Dispose();
    }

    private sealed class PzmcTokenResponse
    {
        [JsonPropertyName("token_name")] public string TokenName { get; set; } = string.Empty;
        [JsonPropertyName("token_api")] public string TokenApi { get; set; } = string.Empty;
    }
}
