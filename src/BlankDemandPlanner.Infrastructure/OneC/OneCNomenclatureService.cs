using System.Diagnostics;
using System.Text.Json;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.OneC;

public sealed class OneCNomenclatureService(ILogger<OneCNomenclatureService> logger) : IOneCNomenclatureService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };
    private static readonly string DefaultPythonPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Documents",
        "Codex",
        "1C_Diagnostics",
        ".venv",
        "Scripts",
        "python.exe");

    public async Task<IReadOnlyList<OneCNomenclatureItem>> SearchAsync(string query, int limit, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return [];
        }

        var root = FindWorkspaceRoot();
        var scriptPath = Path.Combine(root, "onec_integration", "tools", "search_nomenclature.py");
        var outPath = Path.Combine(root, "onec_integration", "reports", $"nomenclature_lookup_{Guid.NewGuid():N}.json");
        var pythonPath = File.Exists(DefaultPythonPath) ? DefaultPythonPath : "python";
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("Не найден скрипт поиска номенклатуры 1С.", scriptPath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add(scriptPath);
        startInfo.ArgumentList.Add("--query");
        startInfo.ArgumentList.Add(query.Trim());
        startInfo.ArgumentList.Add("--limit");
        startInfo.ArgumentList.Add(Math.Clamp(limit, 1, 500).ToString());
        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить поиск номенклатуры 1С.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала файл поиска номенклатуры. Код завершения: {process.ExitCode}. {stderr}".Trim());
        }

        var payload = JsonSerializer.Deserialize<OneCNomenclaturePayload>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат поиска номенклатуры 1С.");
        if (!payload.Ok || process.ExitCode != 0)
        {
            var errors = payload.Errors.Count == 0 ? stderr : string.Join("; ", payload.Errors);
            throw new InvalidOperationException($"Поиск номенклатуры 1С завершился с ошибкой. {errors}".Trim());
        }

        logger.LogInformation("1C nomenclature search for {Query}: {Count} rows, stdout {Stdout}", query, payload.Items.Count, stdout);
        return payload.Items
            .Where(x => !string.IsNullOrWhiteSpace(x.Code) || !string.IsNullOrWhiteSpace(x.Article))
            .Select(x => new OneCNomenclatureItem(
                x.Code?.Trim() ?? string.Empty,
                x.Article?.Trim() ?? string.Empty,
                x.Name?.Trim() ?? string.Empty,
                x.Unit?.Trim() ?? string.Empty))
            .ToList();
    }

    public async Task<IReadOnlyList<OneCNomenclatureItem>> ResolveByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken)
    {
        var requestedCodes = codes
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(500)
            .ToArray();
        if (requestedCodes.Length == 0)
        {
            return [];
        }

        var root = FindWorkspaceRoot();
        var scriptPath = Path.Combine(root, "onec_integration", "tools", "resolve_nomenclature_codes.py");
        var outPath = Path.Combine(root, "onec_integration", "reports", $"nomenclature_codes_lookup_{Guid.NewGuid():N}.json");
        var pythonPath = File.Exists(DefaultPythonPath) ? DefaultPythonPath : "python";
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("Не найден скрипт пакетного получения номенклатуры 1С.", scriptPath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add(scriptPath);
        startInfo.ArgumentList.Add("--code");
        foreach (var code in requestedCodes)
        {
            startInfo.ArgumentList.Add(code);
        }

        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить пакетное получение номенклатуры 1С.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала файл пакетного получения номенклатуры. Код завершения: {process.ExitCode}. {stderr}".Trim());
        }

        var payload = JsonSerializer.Deserialize<OneCNomenclaturePayload>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат пакетного получения номенклатуры 1С.");
        if (!payload.Ok || process.ExitCode != 0)
        {
            var errors = payload.Errors.Count == 0 ? stderr : string.Join("; ", payload.Errors);
            throw new InvalidOperationException($"Пакетное получение номенклатуры 1С завершилось с ошибкой. {errors}".Trim());
        }

        logger.LogInformation("1C nomenclature resolve by codes: requested {Requested}, returned {Count}, stdout {Stdout}", requestedCodes.Length, payload.Items.Count, stdout);
        return payload.Items
            .Where(x => !string.IsNullOrWhiteSpace(x.Code) || !string.IsNullOrWhiteSpace(x.Article))
            .Select(x => new OneCNomenclatureItem(
                x.Code?.Trim() ?? string.Empty,
                x.Article?.Trim() ?? string.Empty,
                x.Name?.Trim() ?? string.Empty,
                x.Unit?.Trim() ?? string.Empty))
            .ToList();
    }

    public async Task<IReadOnlyList<OneCNomenclaturePrice>> ResolvePricesByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken)
    {
        var requestedCodes = codes
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(500)
            .ToArray();
        if (requestedCodes.Length == 0)
        {
            return [];
        }

        var root = FindWorkspaceRoot();
        var scriptPath = Path.Combine(root, "onec_integration", "tools", "resolve_nomenclature_prices.py");
        var outPath = Path.Combine(root, "onec_integration", "reports", $"nomenclature_prices_lookup_{Guid.NewGuid():N}.json");
        var pythonPath = File.Exists(DefaultPythonPath) ? DefaultPythonPath : "python";
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("Не найден скрипт пакетного получения цен номенклатуры 1С.", scriptPath);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = pythonPath,
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add(scriptPath);
        startInfo.ArgumentList.Add("--code");
        foreach (var code in requestedCodes)
        {
            startInfo.ArgumentList.Add(code);
        }

        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить пакетное получение цен номенклатуры 1С.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала файл пакетного получения цен. Код завершения: {process.ExitCode}. {stderr}".Trim());
        }

        var payload = JsonSerializer.Deserialize<OneCNomenclaturePricePayload>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат пакетного получения цен 1С.");
        if (!payload.Ok || process.ExitCode != 0)
        {
            var errors = payload.Errors.Count == 0 ? stderr : string.Join("; ", payload.Errors);
            throw new InvalidOperationException($"Пакетное получение цен 1С завершилось с ошибкой. {errors}".Trim());
        }

        logger.LogInformation("1C nomenclature price resolve by codes: requested {Requested}, returned {Count}, stdout {Stdout}", requestedCodes.Length, payload.Items.Count, stdout);
        return payload.Items
            .Where(x => (!string.IsNullOrWhiteSpace(x.Code) || !string.IsNullOrWhiteSpace(x.Article)) && x.Price > 0)
            .Select(x => new OneCNomenclaturePrice(
                x.Code?.Trim() ?? string.Empty,
                x.Article?.Trim() ?? string.Empty,
                x.Price,
                x.Currency?.Trim() ?? string.Empty,
                x.PriceType?.Trim() ?? string.Empty))
            .ToList();
    }

    private static string FindWorkspaceRoot()
    {
        foreach (var basePath in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            var directory = new DirectoryInfo(basePath);
            for (var i = 0; directory is not null && i < 10; i++, directory = directory.Parent)
            {
                if (Directory.Exists(Path.Combine(directory.FullName, "onec_integration")))
                {
                    return directory.FullName;
                }
            }
        }

        return Environment.CurrentDirectory;
    }

    private sealed record OneCNomenclaturePayload(bool Ok, string? GeneratedAt, string? Database, string? Query, List<OneCNomenclaturePayloadItem> Items, List<string> Errors);
    private sealed record OneCNomenclaturePayloadItem(string? Code, string? Article, string? Name, string? Unit);
    private sealed record OneCNomenclaturePricePayload(bool Ok, string? GeneratedAt, string? Database, List<OneCNomenclaturePricePayloadItem> Items, List<string> Errors);
    private sealed record OneCNomenclaturePricePayloadItem(string? Code, string? Article, decimal Price, string? Currency, string? PriceType);
}
