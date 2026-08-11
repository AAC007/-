using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;

namespace BlankDemandPlanner.UI.ViewModels;

public sealed partial class WorkshopReportViewModel(
    BlankDemandPlannerDbContext dbContext,
    IPzmcPersonnelAvailabilityService personnelAvailabilityService) : ObservableObject
{
    private const string DefaultAudience = "Генеральный директор / Технический директор / Совет директоров";

    public ObservableCollection<WorkshopReportKpiRow> KpiRows { get; } = [];
    public ObservableCollection<WorkshopReportOutputRow> OutputRows { get; } = [];

    [ObservableProperty] private DateTime? periodStart = new(DateTime.Today.Year, DateTime.Today.Month, 1);
    [ObservableProperty] private DateTime? periodEnd = DateTime.Today;
    [ObservableProperty] private string audience = DefaultAudience;
    [ObservableProperty] private string plannedDowntimeText = string.Empty;
    [ObservableProperty] private string unplannedDowntimeText = string.Empty;
    [ObservableProperty] private string repairsText = string.Empty;
    [ObservableProperty] private string incidentsText = string.Empty;
    [ObservableProperty] private string payrollText = string.Empty;
    [ObservableProperty] private string materialCostText = string.Empty;
    [ObservableProperty] private string reportText = "Выберите период и нажмите «Сформировать отчет».";
    [ObservableProperty] private string statusText = "Отчет готов к формированию.";
    [ObservableProperty] private bool isBusy;

    [RelayCommand]
    public async Task LoadAsync()
    {
        if (KpiRows.Count == 0)
        {
            await GenerateReportCoreAsync(refreshOneC: false);
        }
    }

    [RelayCommand]
    private Task GenerateReportAsync() => GenerateReportCoreAsync(refreshOneC: false);

    [RelayCommand]
    private Task RefreshOneCAndGenerateReportAsync() => GenerateReportCoreAsync(refreshOneC: true);

    private async Task GenerateReportCoreAsync(bool refreshOneC)
    {
        if (IsBusy)
        {
            return;
        }

        var start = (PeriodStart ?? new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)).Date;
        var end = (PeriodEnd ?? DateTime.Today).Date;
        if (end < start)
        {
            StatusText = "Дата окончания периода не может быть раньше даты начала.";
            return;
        }

        IsBusy = true;
        try
        {
            StatusText = refreshOneC
                ? "Обновление read-only отчета 1С за выбранный период..."
                : "Формирование отчета по локальным данным...";
            if (refreshOneC)
            {
                await RefreshOneCReportAsync(start, end);
            }

            var facts = await BuildFactsAsync(start, end);
            Replace(KpiRows, BuildKpis(facts));
            Replace(OutputRows, facts.OutputRows.OrderByDescending(x => x.Quantity).ThenBy(x => x.Nomenclature));
            ReportText = BuildReportText(facts);
            StatusText = $"Отчет сформирован: {FormatDate(start)} - {FormatDate(end)}; факт строк выпуска: {facts.ActualOutputRows}.";
        }
        catch (Exception ex)
        {
            StatusText = $"Не удалось сформировать отчет: {ex.GetBaseException().Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task<WorkshopReportFacts> BuildFactsAsync(DateTime start, DateTime end)
    {
        var endExclusive = end.AddDays(1);
        var latestBatchId = await dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .ThenByDescending(x => x.Id)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();

        var demandRows = latestBatchId is null
            ? []
            : await dbContext.DemandItems.AsNoTracking()
                .Where(x => x.DemandBatchId == latestBatchId.Value &&
                    x.Quantity > 0 &&
                    (x.DemandDate == null || (x.DemandDate.Value.Date >= start && x.DemandDate.Value.Date <= end)))
                .ToListAsync();
        var plannedQuantity = demandRows.Sum(x => x.Quantity);
        var plannedUniqueIps = demandRows.Select(x => x.Ips).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count();

        var scheduleRows = await dbContext.ProductionScheduleEntries.AsNoTracking()
            .Where(x => x.PlannedStart < endExclusive && x.PlannedEnd >= start)
            .ToListAsync();
        var equipment = await dbContext.ProductionEquipment.AsNoTracking().ToListAsync();
        var localEmployees = await dbContext.ProductionEmployees.AsNoTracking().ToListAsync();

        var oneC = LoadOneCReport(start, end);
        PzmcPersonnelAvailabilityResult? personnel = null;
        try
        {
            personnel = await personnelAvailabilityService.GetAvailabilityAsync(forceRefresh: false, CancellationToken.None);
        }
        catch
        {
            // The report must remain available when the production API is offline.
        }

        var activeEquipment = equipment.Count(x => x.IsActive);
        var periodDays = Math.Max(1, (end - start).Days + 1);
        var plannedEquipmentMinutes = Math.Max(1, activeEquipment * periodDays * 9 * 60);
        var usedEquipmentMinutes = scheduleRows.Sum(x => Math.Max(0, (x.PlannedEnd - x.PlannedStart).TotalMinutes));
        var oee = Math.Min(100m, (decimal)usedEquipmentMinutes / plannedEquipmentMinutes * 100m);

        var activePersonnel = localEmployees.Count(x => x.IsAvailable);
        var factPersonnel = personnel?.Rows.Count(x => x.AvailableToday) ?? activePersonnel;
        var personnelStatus = personnel?.Status ?? "Данные персонала ПО «Производство» недоступны; использован локальный справочник.";

        var actualOutputQuantity = oneC.OutputRows.Sum(x => x.Quantity);
        var completionPercent = plannedQuantity <= 0 ? 0m : actualOutputQuantity / plannedQuantity * 100m;
        var onTimeRisk = scheduleRows.Count(x => x.PlannedEnd > x.DueDate);
        var completedOperations = oneC.AssemblyRows;

        return new WorkshopReportFacts(
            start,
            end,
            Audience.Trim().Length == 0 ? DefaultAudience : Audience.Trim(),
            plannedQuantity,
            plannedUniqueIps,
            demandRows.Count,
            actualOutputQuantity,
            oneC.OutputRows.Select(x => x.Code).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            oneC.OutputRows.Count,
            completionPercent,
            oee,
            activeEquipment,
            onTimeRisk,
            completedOperations,
            localEmployees.Count,
            factPersonnel,
            personnelStatus,
            oneC.Source,
            oneC.Status,
            oneC.OutputRows,
            PlannedDowntimeText,
            UnplannedDowntimeText,
            RepairsText,
            IncidentsText,
            PayrollText,
            MaterialCostText);
    }

    private static IReadOnlyList<WorkshopReportKpiRow> BuildKpis(WorkshopReportFacts facts) =>
    [
        new("План выпуска", $"{FormatDecimal(facts.PlannedQuantity)} шт", $"{facts.PlannedUniqueIps} уникальных IPS; строк потребности: {facts.PlannedRows}"),
        new("Фактический выпуск", $"{FormatDecimal(facts.ActualOutputQuantity)} шт", $"{facts.ActualUniqueItems} номенклатур; строк перемещений: {facts.ActualOutputRows}"),
        new("Выполнение плана", $"{facts.CompletionPercent:0.#}%", "Факт перемещений на секцию 1 и покраску / потребность периода"),
        new("OEE / загрузка", $"{facts.OeePercent:0.#}%", $"По рассчитанному плану: оборудование {facts.ActiveEquipment} ед.; рисков сроков: {facts.DeadlineRisks}"),
        new("Комплектации", $"{facts.AssemblyRows} строк", "Списания материалов из 44 секции по документам комплектации"),
        new("Персонал", $"{facts.FactPersonnel}/{facts.PlanPersonnel} чел.", facts.PersonnelStatus)
    ];

    private static string BuildReportText(WorkshopReportFacts facts)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"ОТЧЕТ О ДЕЯТЕЛЬНОСТИ МЕХАНИЧЕСКОГО ЦЕХА");
        builder.AppendLine($"Период: {FormatDate(facts.Start)} - {FormatDate(facts.End)}");
        builder.AppendLine($"Целевая аудитория: {facts.Audience}");
        builder.AppendLine();
        builder.AppendLine("1. ПРОИЗВОДСТВЕННЫЕ ПОКАЗАТЕЛИ");
        builder.AppendLine($"План выпуска: {FormatDecimal(facts.PlannedQuantity)} шт; номенклатура: {facts.PlannedUniqueIps} IPS; строк потребности: {facts.PlannedRows}.");
        builder.AppendLine($"Фактический выпуск: {FormatDecimal(facts.ActualOutputQuantity)} шт; номенклатура: {facts.ActualUniqueItems}; строк перемещений на секцию 1 и покраску: {facts.ActualOutputRows}.");
        builder.AppendLine($"Процент выполнения плана: {facts.CompletionPercent:0.#}%.");
        builder.AppendLine($"OEE / коэффициент использования оборудования по рассчитанному плану: {facts.OeePercent:0.#}%.");
        builder.AppendLine($"Lean-оценка: рисков нарушения сроков по плану: {facts.DeadlineRisks}; строк комплектации материалов из 44 секции: {facts.AssemblyRows}.");
        builder.AppendLine();
        builder.AppendLine("2. ТЕХНИЧЕСКОЕ СОСТОЯНИЕ И РЕМОНТЫ");
        builder.AppendLine($"Простои плановые: {ValueOrDash(facts.PlannedDowntime)}.");
        builder.AppendLine($"Простои внеплановые: {ValueOrDash(facts.UnplannedDowntime)}.");
        builder.AppendLine($"Выполненные ТО и ремонты: {ValueOrDash(facts.Repairs)}.");
        builder.AppendLine($"Аварийные ситуации и меры: {ValueOrDash(facts.Incidents)}.");
        builder.AppendLine();
        builder.AppendLine("3. ПЕРСОНАЛ");
        builder.AppendLine($"Численность персонала: план {facts.PlanPersonnel} чел.; факт {facts.FactPersonnel} чел.");
        builder.AppendLine($"Источник факта персонала: {facts.PersonnelStatus}");
        builder.AppendLine();
        builder.AppendLine("4. БЮДЖЕТ И РЕСУРСЫ");
        builder.AppendLine($"Расход ФОТ: {ValueOrDash(facts.Payroll)}.");
        builder.AppendLine($"Расход материалов и запчастей: {ValueOrDash(facts.MaterialCost)}.");
        builder.AppendLine();
        builder.AppendLine("5. ВЫВОДЫ НАЧАЛЬНИКА ЦЕХА");
        builder.AppendLine(BuildLeanConclusion(facts));
        builder.AppendLine();
        builder.AppendLine($"Источник фактического выпуска: {facts.OneCStatus}");
        return builder.ToString();
    }

    private static string BuildLeanConclusion(WorkshopReportFacts facts)
    {
        if (facts.PlannedQuantity <= 0)
        {
            return "Потребность за период не загружена или не попала в выбранный диапазон. Для управленческого анализа необходимо обновить потребность и повторить расчет.";
        }

        if (facts.CompletionPercent >= 95m && facts.DeadlineRisks == 0)
        {
            return "Производственный поток стабилен: выполнение плана близко к целевому, критических рисков по срокам в рассчитанном плане нет. Рекомендуется удерживать ежедневный контроль выпуска и не наращивать НЗП сверх потребности.";
        }

        if (facts.CompletionPercent < 80m)
        {
            return "Выполнение плана ниже целевого уровня. Приоритет: снять ограничения по узким местам оборудования/персонала, ежедневно разбирать отклонения по выпуску и переводить дефицитные позиции в короткий сменный план.";
        }

        return "План выполняется частично. Требуется точечная работа по сроковым рискам, приоритетное закрытие номенклатуры с ближайшей датой потребности и контроль перемещений фактического выпуска в 1С.";
    }

    private async Task RefreshOneCReportAsync(DateTime start, DateTime end)
    {
        var root = FindWorkspaceRoot();
        var script = Path.Combine(root, "onec_integration", "tools", "build_onec_ut_reports_odata.py");
        if (!File.Exists(script))
        {
            throw new FileNotFoundException("Не найден скрипт read-only отчета 1С.", script);
        }

        var python = @"C:\Users\dpd\Documents\Codex\1C_Diagnostics\.venv\Scripts\python.exe";
        if (!File.Exists(python))
        {
            python = "python";
        }

        var reportsDirectory = Path.Combine(root, "onec_integration", "reports");
        var outPath = Path.Combine(reportsDirectory, "onec_ut_period_reports.json");
        Directory.CreateDirectory(reportsDirectory);
        var startInfo = new ProcessStartInfo
        {
            FileName = python,
            Arguments = $"\"{script}\" --start {start:yyyy-MM-dd} --end {end:yyyy-MM-dd} --out \"{outPath}\"",
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Не удалось запустить read-only отчет 1С.");
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(120));
        await process.WaitForExitAsync(cts.Token);
        var output = await outputTask;
        var error = await errorTask;
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Отчет 1С завершился с ошибкой: {FirstNotEmpty(error, output, $"ExitCode={process.ExitCode}")}");
        }
    }

    private static OneCWorkshopReport LoadOneCReport(DateTime start, DateTime end)
    {
        var file = FindOneCReportFile(start, end);
        if (file is null)
        {
            return new OneCWorkshopReport([], 0, "Нет локального JSON отчета 1С за выбранный период.", "1С: данные не загружены.");
        }

        try
        {
            using var stream = File.OpenRead(file.FullName);
            using var document = JsonDocument.Parse(stream);
            var root = document.RootElement;
            var transfers = root.GetProperty("reports").GetProperty("transfers_from_44_to_production_or_paint");
            var rows = new List<WorkshopReportOutputRow>();
            if (transfers.TryGetProperty("summary", out var summary) && summary.ValueKind == JsonValueKind.Array)
            {
                foreach (var row in summary.EnumerateArray())
                {
                    rows.Add(new WorkshopReportOutputRow(
                        UiText.Clean(GetString(row, "target_warehouse")),
                        UiText.Clean(GetString(row, "code")),
                        UiText.Clean(GetString(row, "nomenclature")),
                        UiText.Clean(GetString(row, "unit")),
                        GetDecimal(row, "quantity"),
                        GetInt(row, "document_count")));
                }
            }

            var assemblyRows = 0;
            if (root.GetProperty("reports").TryGetProperty("assembly_materials_44_raw_to_details", out var assemblies) &&
                assemblies.TryGetProperty("rows", out var assemblyRowsElement) &&
                assemblyRowsElement.ValueKind == JsonValueKind.Array)
            {
                assemblyRows = assemblyRowsElement.GetArrayLength();
            }

            return new OneCWorkshopReport(
                rows,
                assemblyRows,
                file.FullName,
                $"1С: {file.Name}; сформирован {File.GetLastWriteTime(file.FullName):dd.MM.yyyy HH:mm}");
        }
        catch (Exception ex)
        {
            return new OneCWorkshopReport([], 0, file.FullName, $"1С: JSON найден, но не прочитан: {ex.GetBaseException().Message}");
        }
    }

    private static FileInfo? FindOneCReportFile(DateTime start, DateTime end)
    {
        var dir = new DirectoryInfo(Path.Combine(FindWorkspaceRoot(), "onec_integration", "reports"));
        if (!dir.Exists)
        {
            return null;
        }

        var exactName = $"onec_ut_reports_{start:yyyyMMdd}_{end:yyyyMMdd}.json";
        var singleName = $"onec_ut_reports_{start:yyyyMMdd}.json";
        var candidates = dir.EnumerateFiles("onec_ut*_reports*.json")
            .Concat(dir.EnumerateFiles("onec_ut_period_reports.json"))
            .OrderByDescending(x => x.LastWriteTime)
            .ToList();
        return candidates.FirstOrDefault(x => string.Equals(x.Name, exactName, StringComparison.OrdinalIgnoreCase)) ??
            (start == end ? candidates.FirstOrDefault(x => string.Equals(x.Name, singleName, StringComparison.OrdinalIgnoreCase)) : null) ??
            candidates.FirstOrDefault(x => PeriodMatches(x, start, end));
    }

    private static bool PeriodMatches(FileInfo file, DateTime start, DateTime end)
    {
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllText(file.FullName));
            var period = document.RootElement.GetProperty("period");
            return DateTime.TryParse(GetString(period, "start"), out var jsonStart) &&
                DateTime.TryParse(GetString(period, "end"), out var jsonEnd) &&
                jsonStart.Date == start.Date &&
                jsonEnd.Date == end.Date;
        }
        catch
        {
            return false;
        }
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

        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var known = Path.Combine(documents, "Заказ заготовок для ЦМО");
        return Directory.Exists(Path.Combine(known, "onec_integration"))
            ? known
            : Environment.CurrentDirectory;
    }

    private static string GetString(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null
            ? value.ToString()
            : string.Empty;

    private static decimal GetDecimal(JsonElement element, string name)
    {
        var text = GetString(element, name).Replace(',', '.');
        return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0m;
    }

    private static int GetInt(JsonElement element, string name)
    {
        var text = GetString(element, name);
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;
    }

    private static string FormatDate(DateTime value) => value.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU"));
    private static string FormatDecimal(decimal value) => value.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    private static string ValueOrDash(string? value) => string.IsNullOrWhiteSpace(value) ? "не заполнено" : value.Trim();
    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

    private static void Replace<T>(ObservableCollection<T> target, IEnumerable<T> source)
    {
        target.Clear();
        foreach (var item in source)
        {
            target.Add(item);
        }
    }
}

public sealed record WorkshopReportKpiRow(string Indicator, string Value, string Comment);

public sealed record WorkshopReportOutputRow(
    string TargetWarehouse,
    string Code,
    string Nomenclature,
    string Unit,
    decimal Quantity,
    int DocumentCount);

internal sealed record OneCWorkshopReport(
    IReadOnlyList<WorkshopReportOutputRow> OutputRows,
    int AssemblyRows,
    string Source,
    string Status);

internal sealed record WorkshopReportFacts(
    DateTime Start,
    DateTime End,
    string Audience,
    decimal PlannedQuantity,
    int PlannedUniqueIps,
    int PlannedRows,
    decimal ActualOutputQuantity,
    int ActualUniqueItems,
    int ActualOutputRows,
    decimal CompletionPercent,
    decimal OeePercent,
    int ActiveEquipment,
    int DeadlineRisks,
    int AssemblyRows,
    int PlanPersonnel,
    int FactPersonnel,
    string PersonnelStatus,
    string OneCSource,
    string OneCStatus,
    IReadOnlyList<WorkshopReportOutputRow> OutputRows,
    string PlannedDowntime,
    string UnplannedDowntime,
    string Repairs,
    string Incidents,
    string Payroll,
    string MaterialCost);
