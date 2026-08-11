using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.OneC;

public sealed class OneCProductionLaunchService(ILogger<OneCProductionLaunchService> logger) : IOneCProductionLaunchService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };
    private static readonly string DefaultPythonPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Documents",
        "Codex",
        "1C_Diagnostics",
        ".venv",
        "Scripts",
        "python.exe");

    public async Task<OneCProductionLaunchResult> CreateAssemblyAsync(OneCProductionLaunchRequest request, CancellationToken cancellationToken)
    {
        if (request.Quantity <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(request), "Количество запуска должно быть больше 0.");
        }

        if (request.Components.Count == 0)
        {
            throw new InvalidOperationException("Для запуска детали не найдена активная заготовка.");
        }

        var root = FindWorkspaceRoot();
        var scriptPath = Path.Combine(root, "onec_integration", "tools", "create_production_assembly_odata.py");
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("Не найден скрипт создания комплектации 1С.", scriptPath);
        }

        var reportsDirectory = Path.Combine(root, "onec_integration", "reports");
        Directory.CreateDirectory(reportsDirectory);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
        var safeIps = SanitizeFilePart(request.Ips);
        var payloadPath = Path.Combine(reportsDirectory, $"production_launch_{safeIps}_{stamp}_payload.json");
        var outPath = Path.Combine(reportsDirectory, $"production_launch_{safeIps}_{stamp}_result.json");
        await File.WriteAllTextAsync(payloadPath, JsonSerializer.Serialize(ToPayload(request), JsonOptions), cancellationToken);

        var pythonPath = File.Exists(DefaultPythonPath) ? DefaultPythonPath : "python";
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
        startInfo.ArgumentList.Add("--payload");
        startInfo.ArgumentList.Add(payloadPath);
        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);
        startInfo.ArgumentList.Add("--allow-write");

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить создание комплектации 1С.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала результат комплектации. Код завершения: {process.ExitCode}. {stderr}".Trim());
        }

        var result = JsonSerializer.Deserialize<OneCProductionLaunchPayloadResult>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат создания комплектации 1С.");
        if (!result.Ok || process.ExitCode != 0)
        {
            var errors = result.Errors.Count == 0 ? stderr : string.Join("; ", result.Errors);
            throw new InvalidOperationException($"Создание комплектации 1С завершилось с ошибкой. {errors}".Trim());
        }

        logger.LogInformation("1C production launch created: IPS {Ips}, qty {Quantity}, number {Number}, stdout {Stdout}",
            request.Ips,
            request.Quantity,
            result.Number,
            stdout);
        return new OneCProductionLaunchResult(
            result.Ok,
            result.Number ?? string.Empty,
            result.RefKey ?? string.Empty,
            result.Posted,
            result.Comment ?? string.Empty,
            outPath,
            result.Errors);
    }

    private static OneCProductionLaunchPayload ToPayload(OneCProductionLaunchRequest request) =>
        new(
            request.Ips,
            StockCodeNormalizer.NormalizeForComparison(request.Ips),
            request.Designation,
            request.PartName,
            request.Quantity,
            request.Comment,
            request.Components
                .Select(x => new OneCProductionLaunchPayloadComponent(
                    x.OneCCode,
                    x.Name,
                    x.Quantity,
                    UiTextUnit(x.Unit)))
                .ToList());

    private static string UiTextUnit(BlankDemandPlanner.Core.Enums.MeasurementUnit unit) =>
        unit switch
        {
            BlankDemandPlanner.Core.Enums.MeasurementUnit.Meter => "пог. м",
            BlankDemandPlanner.Core.Enums.MeasurementUnit.Kilogram => "кг",
            _ => "шт"
        };

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

    private static string SanitizeFilePart(string value)
    {
        var invalid = Path.GetInvalidFileNameChars().ToHashSet();
        var text = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim('_');
        return string.IsNullOrWhiteSpace(text) ? "item" : text;
    }

    private sealed record OneCProductionLaunchPayload(
        string Ips,
        string OneCPartCode,
        string Designation,
        string PartName,
        decimal Quantity,
        string Comment,
        IReadOnlyList<OneCProductionLaunchPayloadComponent> Components);

    private sealed record OneCProductionLaunchPayloadComponent(string OneCCode, string Name, decimal Quantity, string Unit);

    private sealed record OneCProductionLaunchPayloadResult(bool Ok, string? Number, string? RefKey, bool Posted, string? Comment, List<string> Errors);
}
