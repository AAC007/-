using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Infrastructure.OneC;

public sealed class OneCGoodsTransferService(ILogger<OneCGoodsTransferService> logger) : IOneCGoodsTransferService
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

    public async Task<OneCGoodsTransferResult> CreateTransferAsync(OneCGoodsTransferRequest request, CancellationToken cancellationToken)
    {
        if (request.Items.Count == 0)
        {
            throw new InvalidOperationException("Для услуг на стороне выберите хотя бы одну деталь из НЗП.");
        }

        if (request.Items.Any(x => x.Quantity <= 0))
        {
            throw new InvalidOperationException("Количество перемещения должно быть больше 0.");
        }

        var root = FindWorkspaceRoot();
        var scriptPath = Path.Combine(root, "onec_integration", "tools", "create_goods_transfer_odata.py");
        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException("Не найден скрипт создания перемещения 1С.", scriptPath);
        }

        var reportsDirectory = Path.Combine(root, "onec_integration", "reports");
        Directory.CreateDirectory(reportsDirectory);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
        var outPath = Path.Combine(reportsDirectory, $"external_service_transfer_{stamp}_result.json");

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
        foreach (var item in request.Items)
        {
            startInfo.ArgumentList.Add("--item");
            startInfo.ArgumentList.Add(item.OneCCode);
            startInfo.ArgumentList.Add(item.Quantity.ToString(CultureInfo.InvariantCulture));
        }

        startInfo.ArgumentList.Add("--comment");
        startInfo.ArgumentList.Add(string.IsNullOrWhiteSpace(request.Comment)
            ? $"Услуги на стороне: {request.ServiceType}"
            : request.Comment.Trim());
        startInfo.ArgumentList.Add("--out");
        startInfo.ArgumentList.Add(outPath);
        startInfo.ArgumentList.Add("--allow-write");

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить создание перемещения 1С.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);
        var stdout = await stdoutTask;
        var stderr = await stderrTask;

        if (!File.Exists(outPath))
        {
            throw new InvalidOperationException($"1С не сформировала результат перемещения. Код завершения: {process.ExitCode}. {stderr}".Trim());
        }

        var result = JsonSerializer.Deserialize<OneCGoodsTransferPayloadResult>(await File.ReadAllTextAsync(outPath, cancellationToken), JsonOptions)
            ?? throw new InvalidOperationException("Не удалось прочитать результат создания перемещения 1С.");
        if (!result.Ok || process.ExitCode != 0)
        {
            var errors = result.Errors.Count == 0 ? stderr : string.Join("; ", result.Errors);
            throw new InvalidOperationException($"Создание перемещения 1С завершилось с ошибкой. {errors}".Trim());
        }

        logger.LogInformation("1C goods transfer created for external service {ServiceType}: rows {Rows}, number {Number}, stdout {Stdout}",
            request.ServiceType,
            request.Items.Count,
            result.Number,
            stdout);
        return new OneCGoodsTransferResult(
            result.Ok,
            result.Number ?? string.Empty,
            result.RefKey ?? string.Empty,
            result.Posted,
            result.Comment ?? string.Empty,
            outPath,
            result.Errors);
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

    private sealed record OneCGoodsTransferPayloadResult(bool Ok, string? Number, string? RefKey, bool Posted, string? Comment, List<string> Errors);
}
