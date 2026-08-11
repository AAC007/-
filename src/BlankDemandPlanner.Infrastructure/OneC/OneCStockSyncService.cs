using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.OneC;

public sealed class OneCStockSyncService(
    BlankDemandPlannerDbContext dbContext,
    ILogger<OneCStockSyncService> logger) : IOneCStockSyncService
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

    public async Task<OneCStockSyncReport> SyncAsync(CancellationToken cancellationToken)
    {
        var root = FindWorkspaceRoot();
        var pythonPath = File.Exists(DefaultPythonPath) ? DefaultPythonPath : "python";
        var odataScriptPath = Path.Combine(root, "onec_integration", "tools", "sync_stock_sections_odata.py");
        var odataOutPath = Path.Combine(root, "onec_integration", "reports", "stock_sections_sync_odata.json");
        var comScriptPath = Path.Combine(root, "onec_integration", "tools", "sync_stock_sections.py");
        var comOutPath = Path.Combine(root, "onec_integration", "reports", "stock_sections_sync.json");

        if (!File.Exists(comScriptPath))
        {
            throw new FileNotFoundException("Не найден резервный COM-скрипт синхронизации остатков 1С.", comScriptPath);
        }

        logger.LogInformation("1C stock sync start: WIP warehouses {WipWarehouses}, production warehouses {ProductionWarehouses}",
            string.Join("; ", StockWarehouseRules.WipWarehouseFilters),
            string.Join("; ", StockWarehouseRules.ProductionWarehouseFilters));
        OneCStockSyncPayload payload;
        ScriptRunResult run;
        if (File.Exists(odataScriptPath))
        {
            try
            {
                using var odataCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                odataCts.CancelAfter(TimeSpan.FromSeconds(20));
                run = await RunSyncScriptAsync(pythonPath, root, odataScriptPath, odataOutPath, odataCts.Token);
                payload = await ReadPayloadAsync(odataOutPath, run, cancellationToken);
                await SaveSnapshotAsync(payload, "OData", cancellationToken);
                var odataReport = BuildReport(payload);
                logger.LogInformation("1C stock sync finished via OData: rows {Rows}, stdout {Stdout}", odataReport.ReadRows, run.Stdout);
                return odataReport;
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "1C OData stock sync failed, fallback to COM/Python script.");
            }
        }

        run = await RunSyncScriptAsync(pythonPath, root, comScriptPath, comOutPath, cancellationToken);
        payload = await ReadPayloadAsync(comOutPath, run, cancellationToken);
        await SaveSnapshotAsync(payload, "COM", cancellationToken);
        var report = BuildReport(payload);
        logger.LogInformation("1C stock sync finished via COM: rows {Rows}, stdout {Stdout}", report.ReadRows, run.Stdout);
        return report;
    }

    private static OneCStockSyncReport BuildReport(OneCStockSyncPayload payload) =>
        new(
            payload.Rows.Count,
            payload.Rows.Count,
            payload.Rows.Count(x => x.Role == "production"),
            payload.Rows.Count(x => x.Role == "wip"),
            ParseDate(payload.GeneratedAt) ?? DateTime.UtcNow,
            []);

    private static async Task<ScriptRunResult> RunSyncScriptAsync(string pythonPath, string root, string scriptPath, string outPath, CancellationToken cancellationToken)
    {
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
        startInfo.ArgumentList.Add("--wip-warehouse");
        foreach (var warehouse in StockWarehouseRules.WipWarehouseFilters)
        {
            startInfo.ArgumentList.Add(warehouse);
        }

        startInfo.ArgumentList.Add("--production-warehouse");
        foreach (var warehouse in StockWarehouseRules.ProductionWarehouseFilters)
        {
            startInfo.ArgumentList.Add(warehouse);
        }

        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить синхронизацию остатков 1С.");
        try
        {
            var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
            var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);
            var stdout = await stdoutTask;
            var stderr = await stderrTask;
            return new ScriptRunResult(process.ExitCode, stdout, stderr);
        }
        catch (OperationCanceledException) when (!process.HasExited)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync(CancellationToken.None);
            throw new TimeoutException("Быстрый канал 1С не уложился в лимит времени и будет заменен резервным COM-запросом.");
        }
    }

    private static async Task<OneCStockSyncPayload> ReadPayloadAsync(string outPath, ScriptRunResult run, CancellationToken cancellationToken)
    {
        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала файл остатков. Код завершения: {run.ExitCode}. {run.Stderr}".Trim());
        }

        var payload = JsonSerializer.Deserialize<OneCStockSyncPayload>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат синхронизации остатков 1С.");
        if (!payload.Ok || payload.Errors.Count > 0 || run.ExitCode != 0)
        {
            var errors = payload.Errors.Count == 0
                ? run.Stderr
                : string.Join("; ", payload.Errors.Select(x => $"{x.Role}/{x.Warehouse}: {x.Error}"));
            throw new InvalidOperationException($"Синхронизация остатков 1С завершилась с ошибкой. {errors}".Trim());
        }

        return payload;
    }

    private async Task SaveSnapshotAsync(OneCStockSyncPayload payload, string provider, CancellationToken cancellationToken)
    {
        var codeKeys = payload.Rows
            .Select(x => StockCodeNormalizer.NormalizeForComparison(x.Code))
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var aliases = await dbContext.BlankAliases
            .Where(x => x.IsActive)
            .ToListAsync(cancellationToken);
        var aliasesByCode = aliases
            .Where(x => codeKeys.Contains(StockCodeNormalizer.NormalizeForComparison(x.OneCCode)))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(a => a.UpdatedAt).First(), StringComparer.OrdinalIgnoreCase);
        var snapshot = new StockSnapshot
        {
            SnapshotDate = ParseDate(payload.GeneratedAt) ?? DateTime.UtcNow,
            ImportedAt = DateTime.UtcNow,
            SourceFile = $"1С {provider} НЗП/ЦМО: {StockWarehouseRules.WipWarehouseSummary}; склад: {StockWarehouseRules.ProductionWarehouseSummary}"
        };
        dbContext.StockSnapshots.Add(snapshot);

        foreach (var row in payload.Rows.Where(x => TryParseDecimal(x.Quantity, out _)))
        {
            var code = row.Code ?? string.Empty;
            aliasesByCode.TryGetValue(StockCodeNormalizer.NormalizeForComparison(code), out var alias);
            dbContext.StockItems.Add(new StockItem
            {
                StockSnapshot = snapshot,
                BlankAlias = alias,
                OneCCode = code,
                SourceName = FirstNotEmpty(row.Name, row.Article, code),
                Quantity = TryParseDecimal(row.Quantity, out var quantity) ? quantity : 0m,
                Unit = ParseUnit(row.Unit),
                Warehouse = FirstNotEmpty(row.Warehouse, row.WarehouseCode, row.Role)
            });
        }

        await dbContext.SaveChangesAsync(cancellationToken);
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

    private static MeasurementUnit ParseUnit(string? value)
    {
        var text = (value ?? string.Empty).ToUpperInvariant();
        if ((text.Contains("М", StringComparison.Ordinal) || text.Contains("M", StringComparison.Ordinal)) &&
            !text.Contains("ММ", StringComparison.Ordinal) &&
            !text.Contains("MM", StringComparison.Ordinal))
        {
            return MeasurementUnit.Meter;
        }

        if (text.Contains("КГ", StringComparison.Ordinal) || text.Contains("KG", StringComparison.Ordinal))
        {
            return MeasurementUnit.Kilogram;
        }

        return MeasurementUnit.Piece;
    }

    private static DateTime? ParseDate(string? value) =>
        DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var result) ? result.ToUniversalTime() : null;

    private static bool TryParseDecimal(string? value, out decimal result) =>
        decimal.TryParse((value ?? string.Empty).Replace(",", ".", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture, out result);

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;

    private sealed record OneCStockSyncPayload(bool Ok, string? GeneratedAt, List<OneCStockSyncRow> Rows, List<OneCStockSyncError> Errors);
    private sealed record OneCStockSyncRow(string? Role, string? Code, string? Article, string? Name, string? Unit, string? WarehouseCode, string? Warehouse, string? Quantity);
    private sealed record OneCStockSyncError(string? Role, string? Warehouse, string? Error);
    private sealed record ScriptRunResult(int ExitCode, string Stdout, string Stderr);
}
