using System.Collections;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.UI.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using OfficeOpenXml;

namespace BlankDemandPlanner.UI.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly BlankDemandPlannerDbContext _dbContext;
    private readonly IExcelImportService _excelImportService;
    private readonly IOneCStockSyncService _oneCStockSyncService;
    private readonly IBlankDemandCalculationService _calculationService;
    private readonly IFileDialogService _fileDialogService;
    private readonly ILogger<MainViewModel> _logger;

    public ObservableCollection<NavigationItem> NavigationItems { get; } = [];

    [ObservableProperty] private NavigationItem? selectedNavigationItem;
    [ObservableProperty] private object? currentPage;
    [ObservableProperty] private string? globalSearch;
    [ObservableProperty] private string undoStatusText = "Отменить";
    [ObservableProperty] private string stockSyncStatusText = "Остатки 1С: не обновлялись";
    [ObservableProperty] private bool isStockSyncRunning;

    public DashboardViewModel Dashboard { get; }
    public DemandViewModel Demand { get; }
    public LibraryViewModel Library { get; }
    public NormalizationViewModel Normalization { get; }
    public StockViewModel Stock { get; }
    public CalculationViewModel Calculation { get; }
    public BlankSelectionViewModel BlankSelection { get; }
    public HistoryViewModel History { get; }
    public SettingsViewModel Settings { get; }
    public MskViewModel Msk { get; }

    public MainViewModel(
        BlankDemandPlannerDbContext dbContext,
        IExcelImportService excelImportService,
        IOneCStockSyncService oneCStockSyncService,
        IBlankDemandCalculationService calculationService,
        IBlankNormalizationService normalizationService,
        IReportExportService reportExportService,
        IFileDialogService fileDialogService,
        ILogger<MainViewModel> logger)
    {
        _dbContext = dbContext;
        _excelImportService = excelImportService;
        _oneCStockSyncService = oneCStockSyncService;
        _calculationService = calculationService;
        _fileDialogService = fileDialogService;
        _logger = logger;

        Dashboard = new DashboardViewModel(dbContext);
        Demand = new DemandViewModel(dbContext, excelImportService);
        Library = new LibraryViewModel(dbContext, normalizationService, excelImportService, fileDialogService, reportExportService);
        Normalization = new NormalizationViewModel(dbContext, excelImportService, fileDialogService);
        Stock = new StockViewModel(dbContext);
        Calculation = new CalculationViewModel(dbContext, calculationService, reportExportService, fileDialogService);
        BlankSelection = new BlankSelectionViewModel(dbContext);
        History = new HistoryViewModel();
        Settings = new SettingsViewModel(dbContext);
        Msk = new MskViewModel(dbContext, logger, new IpsBridgeDrawingService(), autoOpenDrawings: true);
        UndoCenter.Changed += (_, _) => UndoStatusText = UndoCenter.StatusText;

        NavigationItems.Add(new NavigationItem("Главная", Dashboard));
        NavigationItems.Add(new NavigationItem("Потребность", Demand));
        NavigationItems.Add(new NavigationItem("Расчет материалов", Calculation));
        NavigationItems.Add(new NavigationItem("Библиотека", Library));
        NavigationItems.Add(new NavigationItem("НСИ", Normalization));
        NavigationItems.Add(new NavigationItem("Подбор заготовок", BlankSelection));
        NavigationItems.Add(new NavigationItem("МСК", Msk));
        NavigationItems.Add(new NavigationItem("История", History));
        NavigationItems.Add(new NavigationItem("Настройки", Settings));
        SelectedNavigationItem = NavigationItems[0];
        _ = Dashboard.LoadAsync();
        _ = Settings.LoadAsync();
        _ = SyncStockFromOneCCoreAsync(showMessage: false);
    }

    partial void OnSelectedNavigationItemChanged(NavigationItem? value)
    {
        CurrentPage = value?.Page;
        if (value?.Page == Library)
        {
            _ = Library.LoadAsync();
        }
        else if (value?.Page == Demand)
        {
            _ = Demand.LoadAsync();
        }
        else if (value?.Page == Dashboard)
        {
            _ = Dashboard.LoadAsync();
        }
        else if (value?.Page == Calculation)
        {
            _ = Calculation.LoadLastRunAsync();
        }
        else if (value?.Page == BlankSelection)
        {
            _ = BlankSelection.LoadAsync();
        }
        else if (value?.Page == Msk)
        {
            _ = Msk.LoadAsync();
        }
        else if (value?.Page == Normalization)
        {
            _ = Normalization.LoadAsync();
        }
        else if (value?.Page == Stock)
        {
            _ = Stock.LoadAsync();
        }
        else if (value?.Page == History)
        {
            _ = History.LoadAsync();
        }
        else if (value?.Page == Settings)
        {
            _ = Settings.LoadAsync();
        }
    }

    partial void OnGlobalSearchChanged(string? value)
    {
        Library.Search = value ?? string.Empty;
        if (CurrentPage == Library)
        {
            _ = Library.LoadAsync();
        }
        else if (CurrentPage == Normalization)
        {
            Normalization.Search = value ?? string.Empty;
            _ = Normalization.LoadAsync();
        }
    }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        await Dashboard.LoadAsync();
        await Demand.LoadAsync();
        await Library.LoadAsync();
        await Normalization.LoadAsync();
        await Stock.LoadAsync();
        await Calculation.LoadLastRunAsync();
        await BlankSelection.LoadAsync();
        await Msk.LoadAsync();
        await History.LoadAsync();
        await Settings.LoadAsync();
    }

    [RelayCommand]
    private async Task ImportDemandAsync()
    {
        var file = _fileDialogService.OpenExcelFile();
        if (file is null)
        {
            return;
        }

        try
        {
            var beforeIps = (await _dbContext.Parts.AsNoTracking().Select(x => x.Ips).ToListAsync())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var report = await _excelImportService.ImportDemandAsync(file, null, CancellationToken.None);
            _dbContext.ChangeTracker.Clear();
            var afterIps = await _dbContext.Parts.AsNoTracking().Select(x => x.Ips).ToListAsync();
            var newPartCount = afterIps.Count(x => !beforeIps.Contains(x));

            SelectedNavigationItem = NavigationItems.First(x => x.Page == Demand);
            CurrentPage = Demand;
            await Demand.LoadAsync();
            await Dashboard.LoadAsync();
            await Library.LoadAsync();
            await BlankSelection.LoadAsync();

            if (newPartCount > 0)
            {
                Demand.StatusText = $"{Demand.StatusText}; добавлено новых деталей: {newPartCount}";
            }

            MessageBox.Show($"Импорт потребности из ПП завершен.\nПрочитано строк: {report.ReadRows}\nДобавлено: {report.AddedRows}\nОбновлено: {report.UpdatedRows}\nПропущено: {report.SkippedRows}\nОшибок: {report.ErrorRows}\nДобавлено новых деталей в библиотеку: {newPartCount}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Import demand failed");
            MessageBox.Show(ex.Message, "Ошибка импорта", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async Task ImportStockAsync() => await ImportAsync("остатков по 1С", file => _excelImportService.ImportStockAsync(file, null, CancellationToken.None));

    [RelayCommand]
    private async Task SyncStockFromOneCAsync() => await SyncStockFromOneCCoreAsync(showMessage: true);

    private async Task SyncStockFromOneCCoreAsync(bool showMessage)
    {
        if (IsStockSyncRunning)
        {
            return;
        }

        IsStockSyncRunning = true;
        StockSyncStatusText = "Остатки 1С: обновление...";
        try
        {
            var report = await _oneCStockSyncService.SyncAsync(CancellationToken.None);
            _dbContext.ChangeTracker.Clear();
            await Dashboard.LoadAsync();
            await Demand.LoadAsync();
            await Normalization.LoadAsync();
            await Stock.LoadAsync();
            await Calculation.LoadLastRunAsync();
            StockSyncStatusText = $"Остатки 1С обновлены: {report.SyncedAt.ToLocalTime():dd.MM.yyyy HH:mm}; склад: {report.WarehouseRows}; НЗП/ЦМО: {report.WipRows}";
            if (showMessage)
            {
                MessageBox.Show($"Остатки 1С обновлены.\nСклад: {report.WarehouseRows} строк\nНЗП/ЦМО: {report.WipRows} строк\nДата: {report.SyncedAt.ToLocalTime():dd.MM.yyyy HH:mm}", "Обновление данных 1С", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "1C stock sync failed");
            var lastStock = await _dbContext.StockSnapshots.AsNoTracking()
                .OrderByDescending(x => x.ImportedAt)
                .Select(x => (DateTime?)x.ImportedAt)
                .FirstOrDefaultAsync();
            StockSyncStatusText = lastStock is null
                ? "Остатки 1С не обновлены"
                : $"Остатки 1С не обновлены; последний снимок {lastStock.Value.ToLocalTime():dd.MM.yyyy HH:mm}";
            if (showMessage)
            {
                MessageBox.Show($"Не удалось обновить остатки из 1С.\n{ex.GetBaseException().Message}\nПредыдущие остатки сохранены.", "Обновление данных 1С", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        finally
        {
            IsStockSyncRunning = false;
        }
    }

    [RelayCommand]
    private async Task CalculateAsync()
    {
        var batchId = await _dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();

        if (batchId is null)
        {
            MessageBox.Show("Сначала загрузите потребность.", "Расчет", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        await _calculationService.CalculateAsync(new CalculationOptions(batchId.Value), CancellationToken.None);
        await Calculation.LoadLastRunAsync();
        SelectedNavigationItem = NavigationItems.First(x => x.Page == Calculation);
        await Dashboard.LoadAsync();
    }

    private async Task ImportAsync(string title, Func<string, Task<ImportReport>> import)
    {
        var file = _fileDialogService.OpenExcelFile();
        if (file is null)
        {
            return;
        }

        try
        {
            var report = await import(file);
            _dbContext.ChangeTracker.Clear();
            var refreshWarning = await TryRefreshAfterImportAsync(title);
            MessageBox.Show($"Импорт {title} завершен.\nПрочитано строк: {report.ReadRows}\nДобавлено: {report.AddedRows}\nОбновлено: {report.UpdatedRows}\nПропущено: {report.SkippedRows}\nОшибок: {report.ErrorRows}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
            if (!string.IsNullOrWhiteSpace(refreshWarning))
            {
                MessageBox.Show(refreshWarning, "Обновление экрана", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Import {Title} failed", title);
            MessageBox.Show($"Не удалось импортировать файл.\n{BuildUserErrorMessage(ex)}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string BuildUserErrorMessage(Exception exception)
    {
        var baseException = exception.GetBaseException();
        return ReferenceEquals(baseException, exception)
            ? exception.Message
            : $"{exception.Message}\n\nПричина: {baseException.Message}";
    }

    private async Task<string?> TryRefreshAfterImportAsync(string title)
    {
        try
        {
            await RefreshAsync();
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Refresh after import {Title} failed", title);
            return $"Импорт {title} завершен, но экран не удалось обновить.\nПричина: {ex.GetBaseException().Message}";
        }
    }

    [RelayCommand]
    private async Task UndoAsync() => UndoStatusText = await UndoCenter.UndoAsync();

    [RelayCommand]
    private void NavigateTo(object? page)
    {
        if (page is null)
        {
            return;
        }

        var item = NavigationItems.FirstOrDefault(x => ReferenceEquals(x.Page, page));
        if (item is not null)
        {
            SelectedNavigationItem = item;
            CurrentPage = item.Page;
        }
    }

    [RelayCommand]
    private async Task NavigateToNsiDuplicatesAsync()
    {
        Normalization.DuplicatesOnly = true;
        await Normalization.LoadAsync();
        NavigateTo(Normalization);
    }
}

public sealed record NavigationItem(string Title, object Page);

public sealed partial class DashboardViewModel(BlankDemandPlannerDbContext dbContext) : ObservableObject
{
    public string LastSoftwareUpdate { get; } = "27.07.2026, версия v2026.07.27.3";
    public string DeveloperInfo { get; } = "Разработал Codex с участием Аракеляна А.С.";

    [ObservableProperty] private int parts;
    [ObservableProperty] private int canonicalBlanks;
    [ObservableProperty] private int partsWithoutBlank;
    [ObservableProperty] private int partsWithoutMsk;
    [ObservableProperty] private int duplicateCandidates;
    [ObservableProperty] private int outsideI012;
    [ObservableProperty] private int purchasePositions;
    [ObservableProperty] private string? lastDemandImport;
    [ObservableProperty] private string? lastStockImport;
    [ObservableProperty] private string? lastCalculation;

    public async Task LoadAsync()
    {
        var activeParts = dbContext.Parts.AsNoTracking()
            .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"));
        var activeAliases = dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive);

        Parts = await activeParts.CountAsync();
        CanonicalBlanks = await activeAliases.CountAsync();
        PartsWithoutBlank = await activeParts.CountAsync(p => !p.BlankMaps.Any(m => m.IsActive));
        PartsWithoutMsk = await activeParts.CountAsync(p => !p.HasMsk);
        var duplicateCounts = await activeAliases
            .Where(x => x.NormalizedSourceName != "")
            .GroupBy(x => x.NormalizedSourceName)
            .Where(g => g.Count() > 1)
            .Select(g => g.Count())
            .ToListAsync();
        DuplicateCandidates = duplicateCounts.Sum();
        OutsideI012 = await activeAliases.CountAsync(x => x.CanonicalBlank != null && x.CanonicalBlank.I012Status != I012Status.Allowed);

        var latestCalculationRunId = await dbContext.CalculationRuns.AsNoTracking()
            .OrderByDescending(x => x.StartedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        PurchasePositions = latestCalculationRunId is null
            ? 0
            : await dbContext.CalculationItems.AsNoTracking().CountAsync(x => x.CalculationRunId == latestCalculationRunId && x.PurchaseQuantity > 0);

        var demand = await dbContext.DemandBatches.AsNoTracking().OrderByDescending(x => x.ImportedAt).Select(x => (DateTime?)x.ImportedAt).FirstOrDefaultAsync();
        var stock = await dbContext.StockSnapshots.AsNoTracking().OrderByDescending(x => x.ImportedAt).Select(x => (DateTime?)x.ImportedAt).FirstOrDefaultAsync();
        var calculation = await dbContext.CalculationRuns.AsNoTracking().OrderByDescending(x => x.StartedAt).Select(x => (DateTime?)x.StartedAt).FirstOrDefaultAsync();

        LastDemandImport = FormatDate(demand);
        LastStockImport = FormatDate(stock);
        LastCalculation = FormatDate(calculation);
    }

    private static string? FormatDate(DateTime? value) => value?.ToLocalTime().ToString("g");
}

public sealed partial class DemandViewModel(BlankDemandPlannerDbContext dbContext, IExcelImportService? excelImportService = null) : ObservableObject
{
    public ObservableCollection<DemandRow> Rows { get; } = [];
    public ObservableCollection<DemandDetailRow> DetailRows { get; } = [];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private string statusText = "Потребность не загружена";
    [ObservableProperty] private string batchText = "Последняя потребность: нет данных";
    [ObservableProperty] private string detailTitle = "Детализация IPS";
    [ObservableProperty] private bool isDetailPanelVisible = true;
    [ObservableProperty] private DemandRow? selectedRow;
    [ObservableProperty] private int daysLimit = 0;

    public string DetailPanelButtonText => IsDetailPanelVisible ? "Скрыть детализацию" : "Отобразить детализацию";

    partial void OnSearchChanged(string value) => _ = LoadAsync();
    partial void OnDaysLimitChanged(int value) => _ = LoadAsync();
    partial void OnIsDetailPanelVisibleChanged(bool value) => OnPropertyChanged(nameof(DetailPanelButtonText));
    partial void OnSelectedRowChanged(DemandRow? value) => ShowDetails(value);

    [RelayCommand]
    public async Task LoadAsync()
    {
        var batch = await dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .FirstOrDefaultAsync();

        Rows.Clear();
        DetailRows.Clear();
        if (batch is null)
        {
            BatchText = "Последняя потребность: нет данных";
            StatusText = "Потребность не загружена";
            DetailTitle = "Детализация IPS";
            return;
        }
        BatchText = $"Последняя потребность: {batch.Name}, импорт {batch.ImportedAt.ToLocalTime():g}";
        if (await TryRefreshLegacyBatchAsync(batch))
        {
            batch = await dbContext.DemandBatches.AsNoTracking()
                .OrderByDescending(x => x.ImportedAt)
                .FirstAsync();
        }

        var query = dbContext.DemandItems.AsNoTracking()
            .Where(x => x.DemandBatchId == batch.Id);

        if (DaysLimit > 0)
        {
            var dateTo = DateTime.Today.AddDays(DaysLimit);
            query = query.Where(x => x.DemandDate == null || x.DemandDate <= dateTo);
        }

        if (!string.IsNullOrWhiteSpace(Search))
        {
            var searchValue = Search.Trim();
            query = query.Where(x =>
                (x.Project != null && x.Project.Contains(searchValue)) ||
                (x.SerialNumber != null && x.SerialNumber.Contains(searchValue)) ||
                (x.ProductionSystem != null && x.ProductionSystem.Contains(searchValue)) ||
                (x.SourcePartName != null && x.SourcePartName.Contains(searchValue)) ||
                x.Ips.Contains(searchValue));
        }

        var items = await query
            .OrderBy(x => x.DemandDate)
            .ThenBy(x => x.Project)
            .ThenBy(x => x.SerialNumber)
            .ThenBy(x => x.Id)
            .Take(5000)
            .ToListAsync();

        var workInProgressByIps = await LoadWorkInProgressByIpsAsync(
            items.Select(x => x.Ips).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());

        foreach (var item in items)
        {
            var inProductionBefore = WorkInProgressBefore(item.Ips, workInProgressByIps);
            var inProduction = ApplyWorkInProgress(item.Ips, item.Quantity, workInProgressByIps);
            var inProductionAfter = WorkInProgressBefore(item.Ips, workInProgressByIps);
            Rows.Add(new DemandRow(
                item.Id,
                UiText.Clean(item.Project),
                UiText.Clean(item.SerialNumber),
                item.Ips,
                UiText.Clean(item.SourcePartName ?? item.Ips),
                UiText.Clean(item.Unit),
                FormatDecimal(item.Quantity),
                FormatDecimal(inProductionBefore),
                FormatDecimal(inProductionBefore),
                FormatDecimal(inProductionAfter),
                item.DemandDate?.ToLocalTime().ToString("dd.MM.yyyy") ?? string.Empty,
                false));
        }

        var uniqueIps = Rows
            .Where(x => !string.IsNullOrWhiteSpace(x.Ips))
            .Select(x => x.Ips.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        var detailQuantity = items.Sum(x => x.Quantity);
        StatusText = $"Показано строк потребности: {Rows.Count}; уникальных IPS: {uniqueIps}; количество деталей: {FormatDecimal(detailQuantity)}";
        SelectedRow = Rows.FirstOrDefault();
    }

    [RelayCommand]
    private void AddManualRow()
    {
        Rows.Insert(0, new DemandRow(0, string.Empty, string.Empty, string.Empty, string.Empty, "шт", "1", "0", "0", "0", DateTime.Today.ToString("dd.MM.yyyy"), true));
        StatusText = "Добавлена новая строка. Заполните IPS детали, количество и сохраните.";
    }

    [RelayCommand]
    private async Task DeleteDemandRowsAsync(object? parameter)
    {
        var rows = ExtractDemandRows(parameter);
        if (rows.Count == 0 && SelectedRow is not null)
        {
            rows.Add(SelectedRow);
        }

        if (rows.Count == 0)
        {
            StatusText = "Выделите одну или несколько строк потребности для удаления.";
            return;
        }

        var batch = await dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .FirstOrDefaultAsync();
        var batchId = batch?.Id ?? 0;
        var before = batchId == 0 ? [] : await SnapshotDemandBatchAsync(batchId);
        var unsavedRows = rows
            .Where(x => x.Id == 0)
            .Select(x => (Index: Math.Max(0, Rows.IndexOf(x)), Row: x))
            .ToList();

        foreach (var row in rows)
        {
            Rows.Remove(row);
        }

        var ids = rows.Where(x => x.Id > 0).Select(x => x.Id).Distinct().ToArray();
        if (ids.Length > 0)
        {
            var items = await dbContext.DemandItems.Where(x => ids.Contains(x.Id)).ToListAsync();
            dbContext.DemandItems.RemoveRange(items);
            await dbContext.SaveChangesAsync();
            await LoadAsync();
        }

        UndoCenter.Push("Отмена удаления строк потребности", async () =>
        {
            if (batchId != 0)
            {
                await RestoreDemandBatchAsync(batchId, before);
            }

            foreach (var (index, row) in unsavedRows.OrderBy(x => x.Index))
            {
                Rows.Insert(Math.Min(index, Rows.Count), row);
            }
        });

        StatusText = rows.Count == 1
            ? "Удалена строка потребности."
            : $"Удалено строк потребности: {rows.Count}.";
    }

    [RelayCommand]
    public void ShowDetails(DemandRow? row)
    {
        DetailRows.Clear();
        if (row is null || string.IsNullOrWhiteSpace(row.Ips))
        {
            DetailTitle = "Детализация IPS";
            return;
        }

        var details = Rows
            .Where(x => string.Equals(x.Ips, row.Ips, StringComparison.OrdinalIgnoreCase))
            .ToList();
        var totalQuantity = details.Sum(x => TryParseDisplayQuantity(x.Quantity, out var quantity) ? quantity : 0m);
        var totalInProduction = details.Sum(x => TryParseDisplayQuantity(x.InProductionQuantity, out var quantity) ? quantity : 0m);
        var deficit = Math.Max(0m, totalQuantity - totalInProduction);
        DetailTitle = $"Детализация IPS {row.Ips}; всего деталей: {FormatDecimal(totalQuantity)}; дефицит: {FormatDecimal(deficit)}";
        foreach (var detail in details
            .Select(x => new DemandDetailRow(
                x.DemandDate,
                x.Quantity,
                x.InProductionQuantity,
                x.WorkInProgressAfter,
                x.Project,
                x.MachineNumber)))
        {
            DetailRows.Add(detail);
        }
    }

    private static bool TryParseDisplayQuantity(string value, out decimal quantity) =>
        decimal.TryParse((value ?? string.Empty).Trim().Replace('.', ','), NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out quantity);

    private static List<DemandRow> ExtractDemandRows(object? parameter)
    {
        if (parameter is DemandRow row)
        {
            return [row];
        }

        if (parameter is IEnumerable selectedItems && parameter is not string)
        {
            var rows = selectedItems.OfType<DemandRow>().ToList();
            if (rows.Count > 0)
            {
                return rows;
            }

            return selectedItems.OfType<DataGridCellInfo>()
                .Select(x => x.Item)
                .OfType<DemandRow>()
                .DistinctBy(x => x.Id == 0 ? RuntimeHelpers.GetHashCode(x) : x.Id)
                .ToList();
        }

        return [];
    }

    [RelayCommand]
    private void ToggleDetailPanel() => IsDetailPanelVisible = !IsDetailPanelVisible;

    [RelayCommand]
    private async Task SaveManualRowsAsync()
    {
        var batch = await GetOrCreateManualBatchAsync();
        var before = await SnapshotDemandBatchAsync(batch.Id);

        foreach (var row in Rows.Where(x => x.IsManual || x.Id > 0))
        {
            if (string.IsNullOrWhiteSpace(row.Ips))
            {
                continue;
            }

            var item = row.Id == 0
                ? new DemandItem { DemandBatchId = batch.Id }
                : await dbContext.DemandItems.FirstOrDefaultAsync(x => x.Id == row.Id);
            if (item is null)
            {
                continue;
            }

            item.Project = NullIfWhiteSpace(row.Project);
            item.SerialNumber = NullIfWhiteSpace(row.MachineNumber);
            item.ProductionSystem = null;
            item.Ips = row.Ips.Trim();
            item.SourcePartName = NullIfWhiteSpace(row.Name);
            item.Unit = string.IsNullOrWhiteSpace(row.UnitName) ? "шт" : row.UnitName.Trim();
            item.Quantity = TryParseQuantity(row.Quantity, out var quantity) ? quantity : 0m;
            item.DemandDate = TryParseDate(row.DemandDate, out var date) ? date : null;

            if (row.Id == 0)
            {
                dbContext.DemandItems.Add(item);
            }
        }

        await dbContext.SaveChangesAsync();
        var batchId = batch.Id;
        UndoCenter.Push("Отмена изменения потребности", async () => await RestoreDemandBatchAsync(batchId, before));
        await LoadAsync();
        StatusText = "Ручная потребность сохранена.";
    }

    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));

    private async Task<Dictionary<string, decimal>> LoadWorkInProgressByIpsAsync(string[] ipsValues)
    {
        if (ipsValues.Length == 0)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();

        if (latestSnapshotId is null)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var ipsKeys = ipsValues.Select(StockCodeNormalizer.NormalizeForComparison).Where(x => x.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stockItems = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.StockSnapshotId == latestSnapshotId.Value && x.Unit == MeasurementUnit.Piece)
            .ToListAsync();

        return stockItems
            .Where(x => StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse) && ipsKeys.Contains(StockCodeNormalizer.NormalizeForComparison(x.OneCCode)))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.Sum(i => i.Quantity), StringComparer.OrdinalIgnoreCase);
    }

    private static decimal ApplyWorkInProgress(string ips, decimal demandQuantity, IDictionary<string, decimal> workInProgressByIps)
    {
        var key = StockCodeNormalizer.NormalizeForComparison(ips);
        if (!workInProgressByIps.TryGetValue(key, out var inProduction) || inProduction <= 0)
        {
            return 0m;
        }

        var used = Math.Min(demandQuantity, inProduction);
        workInProgressByIps[key] = inProduction - used;
        return used;
    }

    private static decimal WorkInProgressBefore(string ips, IDictionary<string, decimal> workInProgressByIps) =>
        workInProgressByIps.TryGetValue(StockCodeNormalizer.NormalizeForComparison(ips), out var value) ? Math.Max(0m, value) : 0m;

    private static bool TryParseQuantity(string? value, out decimal quantity)
    {
        var normalized = value?.Trim().Replace('.', ',') ?? string.Empty;
        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out quantity);
    }

    private static bool TryParseDate(string? value, out DateTime date) =>
        DateTime.TryParse(value, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out date);

    private static string? NullIfWhiteSpace(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private async Task<DemandBatch> GetOrCreateManualBatchAsync()
    {
        var batch = await dbContext.DemandBatches
            .OrderByDescending(x => x.ImportedAt)
            .FirstOrDefaultAsync();
        if (batch is not null)
        {
            return batch;
        }

        batch = new DemandBatch { Name = "Ручной ввод потребности", SourceFile = "Ручной ввод" };
        dbContext.DemandBatches.Add(batch);
        await dbContext.SaveChangesAsync();
        return batch;
    }

    private async Task<List<DemandSnapshot>> SnapshotDemandBatchAsync(long batchId)
    {
        return await dbContext.DemandItems.AsNoTracking()
            .Where(x => x.DemandBatchId == batchId)
            .Select(x => new DemandSnapshot(x.Id, x.Project, x.SerialNumber, x.ProductionSystem, x.Ips, x.SourcePartName, x.Unit, x.Quantity, x.DemandDate))
            .ToListAsync();
    }

    private async Task RestoreDemandBatchAsync(long batchId, List<DemandSnapshot> snapshot)
    {
        var current = await dbContext.DemandItems.Where(x => x.DemandBatchId == batchId).ToListAsync();
        dbContext.DemandItems.RemoveRange(current.Where(x => snapshot.All(s => s.Id != x.Id)));
        foreach (var saved in snapshot)
        {
            var item = current.FirstOrDefault(x => x.Id == saved.Id);
            if (item is null)
            {
                item = new DemandItem { Id = saved.Id, DemandBatchId = batchId };
                dbContext.DemandItems.Add(item);
            }

            item.Project = saved.Project;
            item.SerialNumber = saved.SerialNumber;
            item.ProductionSystem = saved.ProductionSystem;
            item.Ips = saved.Ips;
            item.SourcePartName = saved.SourcePartName;
            item.Unit = saved.Unit;
            item.Quantity = saved.Quantity;
            item.DemandDate = saved.DemandDate;
        }

        await dbContext.SaveChangesAsync();
        await LoadAsync();
    }

    private async Task<bool> TryRefreshLegacyBatchAsync(DemandBatch batch)
    {
        if (excelImportService is null ||
            string.IsNullOrWhiteSpace(batch.SourceFile) ||
            !File.Exists(batch.SourceFile))
        {
            return false;
        }

        var needsRefresh = await dbContext.DemandItems.AsNoTracking()
            .AnyAsync(x => x.DemandBatchId == batch.Id &&
                (x.Project == null || x.SerialNumber == null || x.Unit == null));
        if (!needsRefresh)
        {
            return false;
        }

        await excelImportService.ImportDemandAsync(batch.SourceFile, null, CancellationToken.None);
        return true;
    }
}

public sealed partial class LibraryViewModel(
    BlankDemandPlannerDbContext dbContext,
    IBlankNormalizationService normalizationService,
    IExcelImportService? excelImportService = null,
    IFileDialogService? fileDialogService = null,
    IReportExportService? reportExportService = null) : ObservableObject
{
    private static readonly BlankType[] MeterBasedBlankTypes =
    [
        BlankType.RoundBar,
        BlankType.SquareBar,
        BlankType.HexBar,
        BlankType.PipeRound,
        BlankType.PipeRectangular,
        BlankType.Angle,
        BlankType.Channel,
        BlankType.IBeam,
        BlankType.BronzeBar
    ];

    public ObservableCollection<LibraryRow> Rows { get; } = [];
    public ObservableCollection<LibraryBlankOption> BlankSuggestions { get; } = [];

    public IReadOnlyList<DisplayOption<BlankType>> BlankTypes { get; } = UiText.BlankTypes;
    public IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes { get; } = UiText.ConsumptionUnitTypes;
    public IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFilters { get; } = UiText.BlankTypeFilters;
    public IReadOnlyList<DisplayOption<MeasurementUnit?>> UnitFilters { get; } = UiText.ConsumptionUnitFilters;
    public IReadOnlyList<DisplayOption<bool?>> BlankStatusFilters { get; } =
    [
        new((bool?)null, "Все"),
        new(false, "Без заготовки"),
        new(true, "С заготовкой")
    ];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankTypeFilter = UiText.BlankTypeFilters[0];
    [ObservableProperty] private DisplayOption<MeasurementUnit?> selectedUnitFilter = UiText.ConsumptionUnitFilters[0];
    [ObservableProperty] private DisplayOption<bool?> selectedBlankStatusFilter = new(null, "Все");
    [ObservableProperty] private string editIps = string.Empty;
    [ObservableProperty] private string editDesignation = string.Empty;
    [ObservableProperty] private string editPartName = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType> selectedBlankType = UiText.BlankTypes[0];
    [ObservableProperty] private string blankSearch = string.Empty;
    [ObservableProperty] private LibraryBlankOption? selectedBlank;
    [ObservableProperty] private string selectedOneCCode = string.Empty;
    [ObservableProperty] private decimal consumptionQuantity = 1m;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private string blankLeadTimeDaysText = "30";
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiText.UnitTypes[0];
    [ObservableProperty] private LibraryRow? selectedRow;
    [ObservableProperty] private string editSource = "Ручной ввод";
    [ObservableProperty] private string editorStatus = string.Empty;
    [ObservableProperty] private string summaryText = "Деталей: 0; без заготовки: 0";
    private bool suppressBlankSearchReload;
    private bool suppressSelectedRowEditorLoad;

    partial void OnSearchChanged(string value) => _ = LoadAsync();
    partial void OnSelectedBlankTypeFilterChanged(DisplayOption<BlankType?> value) => _ = LoadAsync();
    partial void OnSelectedUnitFilterChanged(DisplayOption<MeasurementUnit?> value) => _ = LoadAsync();
    partial void OnSelectedBlankStatusFilterChanged(DisplayOption<bool?> value) => _ = LoadAsync();
    partial void OnSelectedBlankTypeChanged(DisplayOption<BlankType> value) => _ = LoadBlankSuggestionsAsync();
    partial void OnBlankSearchChanged(string value)
    {
        if (!suppressBlankSearchReload)
        {
            if (SelectedBlank is not null && !string.Equals(value, SelectedBlank.DisplayName, StringComparison.Ordinal))
            {
                SelectedBlank = null;
            }

            _ = LoadBlankSuggestionsAsync();
        }
    }
    partial void OnConsumptionQuantityChanged(decimal value) => ConsumptionQuantityText = FormatDecimal(value);
    partial void OnConsumptionQuantityTextChanged(string value)
    {
        if (TryParseQuantity(value, out var quantity))
        {
            ConsumptionQuantity = quantity;
        }
    }
    partial void OnSelectedRowChanged(LibraryRow? value)
    {
        if (!suppressSelectedRowEditorLoad && value is not null)
        {
            _ = EditLibraryRowAsync(value);
        }
    }

    partial void OnEditorStatusChanged(string value)
    {
        var cleaned = UiText.Clean(value);
        if (!string.Equals(cleaned, value, StringComparison.Ordinal))
        {
            EditorStatus = cleaned;
        }
    }

    partial void OnSelectedBlankChanged(LibraryBlankOption? value)
    {
        SelectedOneCCode = value?.OneCCode ?? string.Empty;
        if (value is not null)
        {
            suppressBlankSearchReload = true;
            BlankSearch = value.DisplayName;
            suppressBlankSearchReload = false;
            var unit = IsMeterBasedBlankType(value.BlankType) ? MeasurementUnit.Meter : value.BaseUnit;
            SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == unit) ?? SelectedUnit;
        }
    }

    [RelayCommand]
    private void NewLibraryEntry()
    {
        EditIps = string.Empty;
        EditDesignation = string.Empty;
        EditPartName = string.Empty;
        SelectedBlankType = BlankTypes[0];
        BlankSearch = string.Empty;
        SelectedBlank = null;
        SelectedOneCCode = string.Empty;
        ConsumptionQuantity = 1m;
        ConsumptionQuantityText = "1";
        BlankLeadTimeDaysText = "30";
        SelectedUnit = UnitTypes[0];
        EditSource = "Ручной ввод";
        SelectedRow = null;
        EditorStatus = "Заполните поля и нажмите \"Сохранить\", чтобы добавить строку библиотеки вручную.";
    }

    [RelayCommand]
    private async Task ImportLibraryAsync()
    {
        if (excelImportService is null || fileDialogService is null)
        {
            EditorStatus = "Импорт библиотеки недоступен в тестовом режиме.";
            return;
        }

        var file = fileDialogService.OpenExcelFile();
        if (file is null)
        {
            EditorStatus = "Импорт библиотеки отменен.";
            return;
        }

        try
        {
            var report = await excelImportService.ImportManufacturingBlankLibraryAsync(file, null, CancellationToken.None);
            EditorStatus = $"Импорт библиотеки завершен. Прочитано: {report.ReadRows}; добавлено: {report.AddedRows}; обновлено: {report.UpdatedRows}; ошибок: {report.ErrorRows}.";
            await LoadAsync();
            MessageBox.Show(EditorStatus, "Импорт библиотеки", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            EditorStatus = "Не удалось импортировать библиотеку.";
            MessageBox.Show($"Не удалось импортировать библиотеку.\n{ex.GetBaseException().Message}", "Импорт библиотеки", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async Task ExportLibraryAsync()
    {
        if (reportExportService is null || fileDialogService is null)
        {
            EditorStatus = "Выгрузка библиотеки недоступна в тестовом режиме.";
            return;
        }

        var folder = fileDialogService.SelectFolder();
        if (folder is null)
        {
            EditorStatus = "Выгрузка библиотеки отменена.";
            return;
        }

        try
        {
            var path = await reportExportService.ExportLibraryAsync(folder, CancellationToken.None);
            EditorStatus = $"Библиотека выгружена: {path}";
            MessageBox.Show(EditorStatus, "Выгрузка библиотеки", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            EditorStatus = "Не удалось выгрузить библиотеку.";
            MessageBox.Show($"Не удалось выгрузить библиотеку.\n{ex.GetBaseException().Message}", "Выгрузка библиотеки", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    public async Task LoadAsync()
    {
        await RestoreImportedArchivedLibraryPartsAsync();
        await EnsureDemandPartsInLibraryAsync();

        var query = dbContext.Parts.AsNoTracking()
            .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
            .AsQueryable();
        if (!string.IsNullOrWhiteSpace(Search))
        {
            var searchValue = Search.Trim();
            query = query.Where(x =>
                x.Ips.Contains(searchValue) ||
                (x.Designation != null && x.Designation.Contains(searchValue)) ||
                x.Name.Contains(searchValue) ||
                x.BlankMaps.Any(m => m.IsActive && m.CanonicalBlank != null &&
                    (m.CanonicalBlank.CanonicalName.Contains(searchValue) ||
                     (m.CanonicalBlank.Material != null && m.CanonicalBlank.Material.Contains(searchValue)) ||
                     m.CanonicalBlank.Aliases.Any(a => a.OneCCode.Contains(searchValue) || a.SourceName.Contains(searchValue)))));
        }

        if (SelectedBlankTypeFilter.Value is not null)
        {
            var blankType = SelectedBlankTypeFilter.Value.Value;
            query = query.Where(x => x.BlankMaps.Any(m => m.IsActive && m.CanonicalBlank != null && m.CanonicalBlank.BlankType == blankType));
        }

        if (SelectedUnitFilter.Value is not null)
        {
            var unit = SelectedUnitFilter.Value.Value;
            query = unit == MeasurementUnit.Meter
                ? query.Where(x => x.BlankMaps.Any(m => m.IsActive && (m.ConsumptionUnit == MeasurementUnit.Meter || (m.CanonicalBlank != null && MeterBasedBlankTypes.Contains(m.CanonicalBlank.BlankType)))))
                : query.Where(x => x.BlankMaps.Any(m => m.IsActive && m.ConsumptionUnit == unit && (m.CanonicalBlank == null || !MeterBasedBlankTypes.Contains(m.CanonicalBlank.BlankType))));
        }

        if (SelectedBlankStatusFilter.Value is not null)
        {
            query = SelectedBlankStatusFilter.Value.Value
                ? query.Where(x => x.BlankMaps.Any(m => m.IsActive))
                : query.Where(x => !x.BlankMaps.Any(m => m.IsActive));
        }

        var parts = await query
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .OrderBy(x => x.Ips)
            .Take(1000)
            .ToListAsync();

        Rows.Clear();
        foreach (var part in parts)
        {
            var map = part.BlankMaps.FirstOrDefault(x => x.IsActive && x.IsPrimary) ?? part.BlankMaps.FirstOrDefault(x => x.IsActive);
            var blank = map?.CanonicalBlank;
            var activeAlias = blank?.Aliases.FirstOrDefault(x => x.IsActive);
            var effectiveUnit = GetEffectiveUnit(map, blank);
            Rows.Add(new LibraryRow(
                part.Id,
                map?.Id,
                blank?.Id,
                part.Ips,
                UiText.Clean(part.Designation),
                UiText.Clean(part.Name),
                blank is null ? null : DisplayBlankType(blank.BlankType),
                UiText.Clean(activeAlias?.SourceName ?? blank?.CanonicalName),
                UiText.Clean(blank?.Material),
                activeAlias?.OneCCode,
                map?.ConsumptionQuantity,
                effectiveUnit,
                FormatDecimal(map?.ConsumptionQuantity),
                effectiveUnit is null ? string.Empty : DisplayUnit(effectiveUnit.Value),
                map?.BlankLeadTimeDays ?? 30,
                (map?.BlankLeadTimeDays ?? 30).ToString(CultureInfo.InvariantCulture),
                UiText.Clean(part.Source),
                part.UpdatedAt.ToLocalTime().ToString("g")));
        }

        SummaryText = $"Деталей: {Rows.Select(x => x.PartId).Distinct().Count()}; без заготовки: {Rows.Count(x => x.PartBlankMapId is null || x.CanonicalBlankId is null)}";
        await LoadBlankSuggestionsAsync();
    }

    private async Task RestoreImportedArchivedLibraryPartsAsync()
    {
        var parts = await dbContext.Parts
            .Where(x => x.Source != null &&
                x.Source.Contains("[ARCHIVED_LIBRARY]") &&
                x.BlankMaps.Any(m => m.IsActive))
            .ToListAsync();
        if (parts.Count == 0)
        {
            return;
        }

        foreach (var part in parts)
        {
            part.Source = RestoreLibrarySource(part.Source);
            part.UpdatedAt = DateTime.UtcNow;
        }

        await dbContext.SaveChangesAsync();
    }

    private async Task EnsureDemandPartsInLibraryAsync()
    {
        var demandItems = await dbContext.DemandItems.AsNoTracking()
            .Where(x => x.Ips != "")
            .ToListAsync();
        var demandParts = demandItems
            .Where(IsDemandLibraryPartCandidate)
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .Select(x => new
            {
                Ips = x.Key,
                Name = x.Select(i => i.SourcePartName).FirstOrDefault(n => !string.IsNullOrWhiteSpace(n)),
            })
            .ToList();
        if (demandParts.Count == 0)
        {
            return;
        }

        var existingIps = (await dbContext.Parts.AsNoTracking().Select(x => x.Ips).ToListAsync())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var added = 0;
        foreach (var demandPart in demandParts.Where(x => !existingIps.Contains(x.Ips)))
        {
            var (designation, name) = SplitDesignationAndName(demandPart.Name);
            dbContext.Parts.Add(new Part
            {
                Ips = demandPart.Ips,
                Designation = designation,
                Name = FirstNotEmpty(name, demandPart.Name, "Из потребности"),
                Source = "Потребность"
            });
            added++;
        }

        if (added > 0)
        {
            await dbContext.SaveChangesAsync();
            EditorStatus = $"Добавлено новых деталей из потребности: {added}.";
        }
    }

    [RelayCommand]
    private async Task LoadBlankSuggestionsAsync()
    {
        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);

        var selectedType = SelectedBlankType.Value;
        var useDetectedShapeFilters = selectedType is BlankType.Unknown or BlankType.RoundBar or BlankType.SquareBar or BlankType.HexBar or BlankType.Sheet or BlankType.Plate or BlankType.PipeRound or BlankType.PipeRectangular;
        if (selectedType != BlankType.Unknown)
        {
            query = selectedType == BlankType.Purchased
                ? query.Where(x => x.BlankType == BlankType.Purchased || x.BlankType == BlankType.CustomBlank || x.BlankType == BlankType.Unknown)
                : query.Where(x => x.BlankType == selectedType);
        }

        var searchText = BlankSearch.Trim();
        if (!string.IsNullOrWhiteSpace(searchText))
        {
            var normalized = await normalizationService.NormalizeAsync(searchText, CancellationToken.None);
            if (selectedType == BlankType.Unknown && normalized.BlankType != BlankType.Unknown)
            {
                query = query.Where(x => x.BlankType == normalized.BlankType);
            }

            if (useDetectedShapeFilters && normalized.Diameter is not null)
            {
                query = query.Where(x => x.DiameterMm == normalized.Diameter);
            }

            if (useDetectedShapeFilters && normalized.Width is not null)
            {
                query = query.Where(x => x.WidthMm == normalized.Width);
            }

            if (useDetectedShapeFilters && !string.IsNullOrWhiteSpace(normalized.Material))
            {
                var material = NormalizeText(normalized.Material);
                query = query.Where(x => x.Material != null && x.Material.ToUpper().Replace(" ", "").Contains(material));
            }

            var parts = searchText.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var namePart = parts.FirstOrDefault();
            var codePart = parts.Length > 1 ? parts[^1] : null;
            query = query.Where(x =>
                x.CanonicalName.Contains(searchText) ||
                (x.Material != null && x.Material.Contains(searchText)) ||
                x.Aliases.Any(a => a.OneCCode.Contains(searchText) || a.SourceName.Contains(searchText)) ||
                (!string.IsNullOrWhiteSpace(namePart) && (x.CanonicalName.Contains(namePart) || x.Aliases.Any(a => a.SourceName.Contains(namePart)))) ||
                (!string.IsNullOrWhiteSpace(codePart) && x.Aliases.Any(a => a.OneCCode.Contains(codePart) || a.SourceName.Contains(codePart))));
        }

        var options = await query
            .OrderBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Take(50)
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .ToListAsync();

        BlankSuggestions.Clear();
        foreach (var option in options)
        {
            BlankSuggestions.Add(option);
        }

        if (SelectedBlank is not null)
        {
            var current = options.FirstOrDefault(x => x.Id == SelectedBlank.Id);
            if (current is not null)
            {
                SelectedBlank = current;
            }
            else
            {
                BlankSuggestions.Insert(0, SelectedBlank);
            }
        }
    }

    [RelayCommand]
    private async Task SaveLibraryEntryAsync()
    {
        if (string.IsNullOrWhiteSpace(EditIps))
        {
            EditorStatus = "Укажите IPS.";
            return;
        }

        if (string.IsNullOrWhiteSpace(EditPartName))
        {
            EditorStatus = "Укажите наименование детали.";
            return;
        }

        var ips = EditIps.Trim();
        var part = await dbContext.Parts
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .FirstOrDefaultAsync(x => x.Ips == ips);

        if (part is null)
        {
            part = new Part
            {
                Ips = ips,
                Designation = NullIfWhiteSpace(EditDesignation),
                Name = EditPartName.Trim(),
                Source = NullIfWhiteSpace(EditSource) ?? "Ручной ввод"
            };
            dbContext.Parts.Add(part);
        }
        else
        {
            part.Designation = NullIfWhiteSpace(EditDesignation);
            part.Name = EditPartName.Trim();
            part.Source = NullIfWhiteSpace(EditSource) ?? part.Source;
            part.UpdatedAt = DateTime.UtcNow;
        }

        var selectedBlank = SelectedBlank ?? await ResolveBlankFromSearchAsync();
        if (selectedBlank is null)
        {
            await dbContext.SaveChangesAsync();
            EditorStatus = $"Сохранена карточка детали: {ips}.";
            await LoadAsync();
            return;
        }

        SelectedBlank = selectedBlank;

        if (!TryParseQuantity(ConsumptionQuantityText, out var quantity) || quantity <= 0)
        {
            EditorStatus = "Количество должно быть положительным числом. Можно вводить через запятую: 0,1 или 0,35.";
            return;
        }

        if (!int.TryParse(BlankLeadTimeDaysText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var leadTimeDays) || leadTimeDays < 0)
        {
            EditorStatus = "Срок потребности заготовки должен быть целым числом дней: 0, 30, 45.";
            return;
        }

        var sameActiveMap = part.BlankMaps.FirstOrDefault(x => x.CanonicalBlankId == selectedBlank.Id && x.IsActive);
        if (sameActiveMap is null)
        {
            foreach (var activeMap in part.BlankMaps.Where(x => x.IsActive))
            {
                activeMap.IsActive = false;
                activeMap.IsPrimary = false;
                activeMap.UpdatedAt = DateTime.UtcNow;
            }

            dbContext.PartBlankMaps.Add(new PartBlankMap
            {
                Part = part,
                CanonicalBlankId = selectedBlank.Id,
                ConsumptionQuantity = quantity,
                ConsumptionUnit = SelectedUnit.Value,
                BlankLeadTimeDays = leadTimeDays,
                Source = NullIfWhiteSpace(EditSource) ?? "Ручной ввод",
                IsPrimary = true,
                IsActive = true
            });
        }
        else
        {
            foreach (var activeMap in part.BlankMaps.Where(x => x.IsActive && x.Id != sameActiveMap.Id))
            {
                activeMap.IsActive = false;
                activeMap.IsPrimary = false;
                activeMap.UpdatedAt = DateTime.UtcNow;
            }

            sameActiveMap.ConsumptionQuantity = quantity;
            sameActiveMap.ConsumptionUnit = SelectedUnit.Value;
            sameActiveMap.BlankLeadTimeDays = leadTimeDays;
            sameActiveMap.Source = NullIfWhiteSpace(EditSource) ?? sameActiveMap.Source;
            sameActiveMap.IsPrimary = true;
            sameActiveMap.UpdatedAt = DateTime.UtcNow;
        }

        await dbContext.SaveChangesAsync();
        EditorStatus = $"Сохранено: {ips} -> {UiText.Clean(selectedBlank.CanonicalName)}";
        await LoadAsync();
    }

    private async Task<LibraryBlankOption?> ResolveBlankFromSearchAsync()
    {
        var searchText = BlankSearch.Trim();
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return null;
        }

        var parts = searchText.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var namePart = parts.FirstOrDefault();
        var codePart = parts.Length > 1 ? parts[^1] : null;
        var selectedType = SelectedBlankType.Value;

        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);

        if (selectedType != BlankType.Unknown)
        {
            query = selectedType == BlankType.Purchased
                ? query.Where(x => x.BlankType == BlankType.Purchased || x.BlankType == BlankType.CustomBlank || x.BlankType == BlankType.Unknown)
                : query.Where(x => x.BlankType == selectedType);
        }

        query = query.Where(x =>
            x.CanonicalName.Contains(searchText) ||
            x.Aliases.Any(a => a.OneCCode.Contains(searchText) || a.SourceName.Contains(searchText)) ||
            (!string.IsNullOrWhiteSpace(namePart) && (x.CanonicalName.Contains(namePart) || x.Aliases.Any(a => a.SourceName.Contains(namePart)))) ||
            (!string.IsNullOrWhiteSpace(codePart) && x.Aliases.Any(a => a.OneCCode.Contains(codePart) || a.SourceName.Contains(codePart))));

        return await query
            .OrderBy(x => x.Aliases.Any(a => a.OneCCode == codePart) ? 0 : 1)
            .ThenBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .FirstOrDefaultAsync();
    }

    [RelayCommand]
    private async Task EditLibraryRowAsync(object? parameter)
    {
        var row = parameter as LibraryRow ?? SelectedRow;
        if (row is null)
        {
            EditorStatus = "Выберите строку библиотеки.";
            return;
        }

        suppressSelectedRowEditorLoad = true;
        try
        {
            SelectedRow = row;
            EditIps = row.Ips;
            EditDesignation = row.Designation ?? string.Empty;
            EditPartName = row.PartName;
            ConsumptionQuantity = row.ConsumptionQuantity ?? 1m;
            ConsumptionQuantityText = FormatDecimal(ConsumptionQuantity);
            BlankLeadTimeDaysText = row.BlankLeadTimeDays.ToString(CultureInfo.InvariantCulture);
            SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == row.ConsumptionUnit) ?? UnitTypes[0];
            EditSource = row.Source ?? "Ручной ввод";
            BlankSearch = row.BlankName ?? string.Empty;

            if (row.CanonicalBlankId is not null)
            {
                var blank = await dbContext.CanonicalBlanks.AsNoTracking().FirstOrDefaultAsync(x => x.Id == row.CanonicalBlankId.Value);
                if (blank is not null)
                {
                    SelectedBlankType = BlankTypes.FirstOrDefault(x => x.Value == blank.BlankType) ?? BlankTypes[0];
                }
            }

            await LoadBlankSuggestionsAsync();
            SelectedBlank = row.CanonicalBlankId is null ? null : BlankSuggestions.FirstOrDefault(x => x.Id == row.CanonicalBlankId.Value);
            if (SelectedBlank is null && row.CanonicalBlankId is not null)
            {
                var option = await dbContext.CanonicalBlanks.AsNoTracking()
                    .Where(x => x.Id == row.CanonicalBlankId.Value)
                    .Select(x => new LibraryBlankOption(
                        x.Id,
                        x.BlankType,
                        x.CanonicalName,
                        x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                        x.Material,
                        x.BaseUnit,
                        x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
                    .FirstOrDefaultAsync();
                if (option is not null)
                {
                    BlankSuggestions.Insert(0, option);
                    SelectedBlank = option;
                }
            }

            EditorStatus = $"Строка загружена в редактор: {row.Ips}.";
        }
        finally
        {
            suppressSelectedRowEditorLoad = false;
        }
    }

    [RelayCommand]
    private async Task DeleteLibraryRowsAsync(object? parameter)
    {
        var rows = ExtractLibraryRows(parameter);
        if (rows.Count == 0 && SelectedRow is not null)
        {
            rows.Add(SelectedRow);
        }

        if (rows.Count == 0)
        {
            EditorStatus = "Выделите одну или несколько строк библиотеки для удаления.";
            return;
        }

        var snapshot = await SnapshotLibraryAsync(rows);
        foreach (var row in rows.DistinctBy(x => new { x.PartId, x.PartBlankMapId }))
        {
            await ArchiveLibraryRowAsync(row, reload: false);
        }

        EditorStatus = rows.Count == 1
            ? $"Удалена или архивирована строка библиотеки: {rows[0].Ips}."
            : $"Удалено или архивировано строк библиотеки: {rows.Count}.";
        UndoCenter.Push("Отмена удаления из библиотеки", async () => await RestoreLibraryAsync(snapshot));
        await LoadAsync();
    }

    private async Task<LibraryDeleteSnapshot> SnapshotLibraryAsync(IEnumerable<LibraryRow> rows)
    {
        var partIds = rows.Select(x => x.PartId).Distinct().ToArray();
        var maps = await dbContext.PartBlankMaps.AsNoTracking()
            .Where(x => partIds.Contains(x.PartId))
            .Select(x => new LibraryMapSnapshot(x.Id, x.IsActive, x.IsPrimary, x.UpdatedAt))
            .ToListAsync();
        var parts = await dbContext.Parts.AsNoTracking()
            .Where(x => partIds.Contains(x.Id))
            .Select(x => new LibraryPartSnapshot(x.Id, x.Source, x.UpdatedAt))
            .ToListAsync();
        return new LibraryDeleteSnapshot(maps, parts);
    }

    private async Task RestoreLibraryAsync(LibraryDeleteSnapshot snapshot)
    {
        var ids = snapshot.Maps.Select(x => x.Id).ToArray();
        var maps = await dbContext.PartBlankMaps.Where(x => ids.Contains(x.Id)).ToListAsync();
        foreach (var saved in snapshot.Maps)
        {
            var map = maps.FirstOrDefault(x => x.Id == saved.Id);
            if (map is null)
            {
                continue;
            }

            map.IsActive = saved.IsActive;
            map.IsPrimary = saved.IsPrimary;
            map.UpdatedAt = saved.UpdatedAt;
        }

        var partIds = snapshot.Parts.Select(x => x.Id).ToArray();
        var parts = await dbContext.Parts.Where(x => partIds.Contains(x.Id)).ToListAsync();
        foreach (var saved in snapshot.Parts)
        {
            var part = parts.FirstOrDefault(x => x.Id == saved.Id);
            if (part is null)
            {
                continue;
            }

            part.Source = saved.Source;
            part.UpdatedAt = saved.UpdatedAt;
        }

        await dbContext.SaveChangesAsync();
        await LoadAsync();
    }

    private static List<LibraryRow> ExtractLibraryRows(object? parameter)
    {
        if (parameter is LibraryRow row)
        {
            return [row];
        }

        if (parameter is IEnumerable selectedItems && parameter is not string)
        {
            var rows = selectedItems.OfType<LibraryRow>().ToList();
            if (rows.Count > 0)
            {
                return rows;
            }

            return selectedItems.OfType<DataGridCellInfo>()
                .Select(x => x.Item)
                .OfType<LibraryRow>()
                .DistinctBy(x => new { x.PartId, x.PartBlankMapId })
                .ToList();
        }

        return [];
    }

    [RelayCommand]
    private async Task DeleteLibraryRowAsync(object? parameter)
    {
        var row = parameter as LibraryRow;
        await ArchiveLibraryRowAsync(row ?? SelectedRow, reload: true);
    }

    private async Task ArchiveLibraryRowAsync(LibraryRow? row, bool reload)
    {
        row ??= SelectedRow;
        if (row is null)
        {
            EditorStatus = "Выберите строку библиотеки для удаления.";
            return;
        }

        var part = await dbContext.Parts
            .Include(x => x.BlankMaps)
            .FirstOrDefaultAsync(x => x.Id == row.PartId);
        if (part is null)
        {
            EditorStatus = "Выбранная строка библиотеки не найдена.";
            if (reload)
            {
                await LoadAsync();
            }
            return;
        }

        if (row.PartBlankMapId is not null)
        {
            var map = part.BlankMaps.FirstOrDefault(x => x.Id == row.PartBlankMapId.Value);
            if (map is not null)
            {
                map.IsActive = false;
                map.IsPrimary = false;
                map.UpdatedAt = DateTime.UtcNow;
            }
        }
        else if (row.CanonicalBlankId is null)
        {
            var demandItems = await dbContext.DemandItems.Where(x => x.PartId == part.Id).ToListAsync();
            foreach (var demandItem in demandItems)
            {
                demandItem.PartId = null;
            }

            part.Source = BuildArchivedSource(part.Source, "Удалено из библиотеки");
            part.UpdatedAt = DateTime.UtcNow;
            EditorStatus = $"Деталь без заготовки убрана из библиотеки: {row.Ips}.";
            await dbContext.SaveChangesAsync();
            if (reload)
            {
                await LoadAsync();
            }
            return;
        }
        else
        {
            foreach (var map in part.BlankMaps.Where(x => x.IsActive))
            {
                map.IsActive = false;
                map.IsPrimary = false;
                map.UpdatedAt = DateTime.UtcNow;
            }
        }

        part.UpdatedAt = DateTime.UtcNow;
        EditorStatus = $"Строка библиотеки архивирована: {row.Ips}.";

        await dbContext.SaveChangesAsync();
        if (reload)
        {
            await LoadAsync();
        }
    }

    private static string? NullIfWhiteSpace(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
    private static bool IsDemandLibraryPartCandidate(DemandItem item)
    {
        if (IsArchivedLibrarySource(item.Part?.Source))
        {
            return false;
        }

        var unit = UiText.Clean(item.Unit).Trim();
        if (!string.IsNullOrWhiteSpace(unit) && !unit.Contains("шт", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var name = UiText.Clean(item.SourcePartName);
        if (string.IsNullOrWhiteSpace(item.Ips) || string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        var lower = name.ToLowerInvariant();
        if (lower.Contains("операц", StringComparison.Ordinal) ||
            lower.Contains("сбор", StringComparison.Ordinal) ||
            lower.Contains("узел", StringComparison.Ordinal) ||
            lower.Contains("комплект", StringComparison.Ordinal) ||
            lower.Contains("монтаж", StringComparison.Ordinal))
        {
            return false;
        }

        var (designation, partName) = SplitDesignationAndName(name);
        return !string.IsNullOrWhiteSpace(designation) && !string.IsNullOrWhiteSpace(partName);
    }

    private static bool IsArchivedLibrarySource(string? source) =>
        !string.IsNullOrWhiteSpace(source) && source.Contains("[ARCHIVED_LIBRARY]", StringComparison.OrdinalIgnoreCase);

    private static string? RestoreLibrarySource(string? source)
    {
        if (string.IsNullOrWhiteSpace(source))
        {
            return source;
        }

        var restored = source.Replace("[ARCHIVED_LIBRARY]", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        return string.IsNullOrWhiteSpace(restored) ? null : restored;
    }

    private static string BuildArchivedSource(string? source, string reason)
    {
        var marker = $"[ARCHIVED_LIBRARY] {reason} {DateTime.UtcNow:yyyy-MM-dd HH:mm}";
        return string.IsNullOrWhiteSpace(source) ? marker : $"{source}; {marker}";
    }

    private static (string? Designation, string Name) SplitDesignationAndName(string? value)
    {
        var text = UiText.Clean(value).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return (null, string.Empty);
        }

        var index = text.IndexOf(' ');
        if (index <= 0)
        {
            return (null, text);
        }

        var first = text[..index].Trim();
        var rest = text[(index + 1)..].Trim();
        var looksLikeDesignation = first.Any(char.IsDigit) &&
            first.Length >= 5 &&
            first.Any(ch => ch is '.' or '-' or '_' or '/');
        return looksLikeDesignation && !string.IsNullOrWhiteSpace(rest)
            ? (first, rest)
            : (null, text);
    }
    private static string NormalizeText(string value) => value.ToUpperInvariant().Replace(" ", string.Empty, StringComparison.Ordinal);
    private static string FormatDecimal(decimal? quantity) => quantity is null ? string.Empty : quantity.Value.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    private static bool TryParseQuantity(string? value, out decimal quantity)
    {
        var normalized = value?.Trim().Replace('.', ',') ?? string.Empty;
        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out quantity);
    }

    private static MeasurementUnit? GetEffectiveUnit(PartBlankMap? map, CanonicalBlank? blank)
    {
        if (map is null)
        {
            return null;
        }

        return blank is not null && IsMeterBasedBlankType(blank.BlankType) && map.ConsumptionUnit == MeasurementUnit.Piece
            ? MeasurementUnit.Meter
            : map.ConsumptionUnit;
    }

    private static bool IsMeterBasedBlankType(BlankType blankType) => MeterBasedBlankTypes.Contains(blankType);

    private static string DisplayUnit(MeasurementUnit unit) => UiText.DisplayUnit(unit);

    private static string DisplayBlankType(BlankType type) => UiText.DisplayBlankType(type);
}

public sealed partial class NormalizationViewModel(
    BlankDemandPlannerDbContext dbContext,
    IExcelImportService excelImportService,
    IFileDialogService fileDialogService) : ObservableObject
{
    private static readonly BlankType[] MeterBasedBlankTypes =
    [
        BlankType.RoundBar,
        BlankType.SquareBar,
        BlankType.HexBar,
        BlankType.PipeRound,
        BlankType.PipeRectangular,
        BlankType.Angle,
        BlankType.Channel,
        BlankType.IBeam,
        BlankType.BronzeBar
    ];

    public ObservableCollection<NsiBlankRow> Rows { get; } = [];
    public ObservableCollection<NsiUsageRow> UsageRows { get; } = [];
    public ObservableCollection<string> MaterialSuggestions { get; } = [];

    public IReadOnlyList<DisplayOption<BlankType>> BlankTypes { get; } = UiText.BlankTypes;
    public IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFilters { get; } = UiText.BlankTypeFiltersWithUnknown;
    public IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes { get; } = UiText.UnitTypes;

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankTypeFilter = UiText.BlankTypeFiltersWithUnknown[0];
    [ObservableProperty] private string sizeFilter = string.Empty;
    [ObservableProperty] private string materialFilter = string.Empty;
    [ObservableProperty] private bool duplicatesOnly;
    [ObservableProperty] private string statusText = string.Empty;
    [ObservableProperty] private NsiBlankRow? selectedRow;
    [ObservableProperty] private DisplayOption<BlankType> editBlankType = UiText.BlankTypes.First(x => x.Value == BlankType.Unknown);
    [ObservableProperty] private DisplayOption<MeasurementUnit> editUnit = UiText.UnitTypes[0];
    [ObservableProperty] private string editOneCCode = string.Empty;
    [ObservableProperty] private string editSourceName = string.Empty;
    [ObservableProperty] private string editMaterial = string.Empty;
    [ObservableProperty] private string editSize = string.Empty;
    [ObservableProperty] private string editMaterialGost = string.Empty;
    [ObservableProperty] private string editProfileGost = string.Empty;
    [ObservableProperty] private bool isUsagePanelVisible = true;
    [ObservableProperty] private string usageStatusText = "Выберите заготовку";
    [ObservableProperty] private string summaryText = "Строк НСИ: 0";
    private bool suppressAutoRestoreNsi;
    public string UsagePanelButtonText => IsUsagePanelVisible ? "Скрыть применяемость" : "Отобразить применяемость";

    partial void OnSearchChanged(string value) => _ = LoadAsync();
    partial void OnSelectedBlankTypeFilterChanged(DisplayOption<BlankType?> value) => _ = LoadAsync();
    partial void OnSizeFilterChanged(string value) => _ = LoadAsync();
    partial void OnMaterialFilterChanged(string value) => _ = LoadAsync();
    partial void OnDuplicatesOnlyChanged(bool value) => _ = LoadAsync();
    partial void OnEditMaterialChanged(string value) => _ = LoadMaterialSuggestionsAsync(value);
    partial void OnSelectedRowChanged(NsiBlankRow? value) => _ = LoadUsageAsync(value);
    partial void OnStatusTextChanged(string value)
    {
        var cleaned = UiText.Clean(value);
        if (!string.Equals(cleaned, value, StringComparison.Ordinal))
        {
            StatusText = cleaned;
        }
    }
    partial void OnIsUsagePanelVisibleChanged(bool value) => OnPropertyChanged(nameof(UsagePanelButtonText));

    [RelayCommand]
    private async Task ImportAsync()
    {
        var file = fileDialogService.OpenExcelFile();
        if (file is null)
        {
            StatusText = "Импорт НСИ отменен.";
            return;
        }

        try
        {
            var report = await excelImportService.ImportOneCBlanksAsync(file, null, CancellationToken.None);
            StatusText = $"Импорт НСИ завершен. Прочитано: {report.ReadRows}; добавлено: {report.AddedRows}; обновлено: {report.UpdatedRows}; ошибок: {report.ErrorRows}.";
            await LoadAsync();
            MessageBox.Show(StatusText, "НСИ", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Не удалось импортировать НСИ.\n{ex.GetBaseException().Message}", "НСИ", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private void ToggleUsagePanel() => IsUsagePanelVisible = !IsUsagePanelVisible;

    [RelayCommand]
    private async Task LoadMaterialSuggestionsAsync(string? value)
    {
        var text = (value ?? string.Empty).Trim();
        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Where(x => x.IsActive && x.Material != null && x.Material != "");

        if (!string.IsNullOrWhiteSpace(text))
        {
            query = query.Where(x => x.Material!.StartsWith(text) || x.Material.Contains(text));
        }

        var suggestions = await query
            .Select(x => x.Material!)
            .Distinct()
            .OrderBy(x => x)
            .Take(40)
            .ToListAsync();

        MaterialSuggestions.Clear();
        foreach (var material in suggestions.Select(UiText.Clean).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            MaterialSuggestions.Add(material);
        }
    }

    [RelayCommand]
    private void EditNsiRow(NsiBlankRow? row)
    {
        row ??= SelectedRow;
        if (row is null)
        {
            StatusText = "Выберите строку НСИ для редактирования.";
            return;
        }

        SelectedRow = row;
        LoadSelectedEditor(row);
        StatusText = $"Строка загружена в редактор: {row.OneCCode}.";
    }

    [RelayCommand]
    private void NewNsiBlank()
    {
        SelectedRow = null;
        EditBlankType = BlankTypes.First(x => x.Value == BlankType.Unknown);
        EditUnit = UnitTypes[0];
        EditOneCCode = string.Empty;
        EditSourceName = string.Empty;
        EditMaterial = string.Empty;
        EditSize = string.Empty;
        EditMaterialGost = string.Empty;
        EditProfileGost = string.Empty;
        StatusText = "Заполните данные заготовки и нажмите \"Сохранить правку НСИ\".";
    }

    [RelayCommand]
    private async Task SaveSelectedEditAsync()
    {
        if (SelectedRow is null)
        {
            await CreateManualNsiBlankAsync();
            return;
        }

        var aliasId = SelectedRow.AliasId;
        var blank = await dbContext.CanonicalBlanks.FirstOrDefaultAsync(x => x.Id == SelectedRow.CanonicalBlankId);
        if (blank is null)
        {
            StatusText = "Заготовка для выбранной позиции НСИ не найдена.";
            return;
        }

        var code = NullIfWhiteSpace(EditOneCCode);
        var name = NullIfWhiteSpace(EditSourceName);
        if (code is null)
        {
            StatusText = "Укажите код УТ / IPS.";
            return;
        }

        if (name is null)
        {
            StatusText = "Укажите наименование по 1С УТ.";
            return;
        }

        var duplicateCode = await dbContext.BlankAliases
            .AsNoTracking()
            .AnyAsync(x => x.Id != aliasId && x.OneCCode == code);
        if (duplicateCode)
        {
            StatusText = $"Код УТ / IPS {code} уже есть в НСИ. Укажите другой код.";
            return;
        }

        blank.BlankType = EditBlankType.Value;
        blank.BaseUnit = EditUnit.Value;
        blank.CanonicalName = name;

        if (!TryApplySize(EditSize, blank, out var sizeError))
        {
            StatusText = sizeError;
            return;
        }

        blank.Material = NullIfWhiteSpace(EditMaterial);
        blank.MaterialGost = NullIfWhiteSpace(EditMaterialGost);
        blank.ProfileGost = NullIfWhiteSpace(EditProfileGost);
        blank.UpdatedAt = DateTime.UtcNow;
        var alias = await dbContext.BlankAliases.FirstOrDefaultAsync(x => x.Id == aliasId);
        if (alias is not null)
        {
            alias.OneCCode = code;
            alias.SourceName = name;
            alias.NormalizedSourceName = NormalizeKey(alias.SourceName);
            alias.UpdatedAt = DateTime.UtcNow;
        }

        try
        {
            await dbContext.SaveChangesAsync();
            StatusText = $"Сохранено: {code}, {DisplayBlankType(blank.BlankType)}, {blank.Material ?? "материал не указан"}, {FormatSize(blank)}";
            await LoadAsync();
            SelectedRow = Rows.FirstOrDefault(x => x.AliasId == aliasId);
        }
        catch (Exception ex)
        {
            var message = ex.GetBaseException().Message;
            try
            {
                dbContext.ChangeTracker.Clear();
                dbContext.Logs.Add(new AppLog
                {
                    Level = "Error",
                    Message = "Ошибка сохранения правки НСИ",
                    Exception = ex.ToString()
                });
                await dbContext.SaveChangesAsync();
            }
            catch
            {
                // If database logging also fails, keep the user-facing error visible.
            }

            StatusText = "Не удалось сохранить правку НСИ. Подробности записаны в лог.";
            MessageBox.Show($"Не удалось сохранить правку НСИ.\n{message}", "НСИ", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    [RelayCommand]
    private async Task ArchiveSelectedAsync() => await ArchiveSelectedRowsAsync(SelectedRow);


    [RelayCommand]
    private async Task DeleteSelectedAsync() => await DeleteSelectedRowsAsync(SelectedRow);


    [RelayCommand]
    private async Task DeleteSelectedRowsAsync(object? parameter)
    {
        var rows = GetSelectedNsiRows(parameter).DistinctBy(x => x.AliasId).ToArray();
        if (rows.Length == 0)
        {
            StatusText = "Выберите одну или несколько строк НСИ для удаления.";
            return;
        }

        suppressAutoRestoreNsi = true;
        var aliasIds = rows.Select(x => x.AliasId).Distinct().ToArray();
        var aliases = await dbContext.BlankAliases
            .Where(x => aliasIds.Contains(x.Id))
            .ToListAsync();
        var aliasesWithHistory = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.BlankAliasId.HasValue && aliasIds.Contains(x.BlankAliasId.Value))
            .Select(x => x.BlankAliasId!.Value)
            .Distinct()
            .ToListAsync();
        var historyIds = aliasesWithHistory.ToHashSet();
        var archived = 0;
        var deleted = 0;

        foreach (var alias in aliases)
        {
            if (historyIds.Contains(alias.Id))
            {
                alias.IsActive = false;
                alias.UpdatedAt = DateTime.UtcNow;
                archived++;
            }
            else
            {
                dbContext.BlankAliases.Remove(alias);
                deleted++;
            }
        }

        await dbContext.SaveChangesAsync();
        StatusText = $"Удаление НСИ: удалено {deleted}, перенесено в архив {archived}.";
        await LoadAsync();
    }

    private async Task CreateManualNsiBlankAsync()
    {
        var code = NullIfWhiteSpace(EditOneCCode);
        var name = NullIfWhiteSpace(EditSourceName);
        if (code is null)
        {
            StatusText = "Укажите код УТ / IPS для новой заготовки.";
            return;
        }

        if (name is null)
        {
            StatusText = "Укажите наименование по 1С УТ для новой заготовки.";
            return;
        }

        var existingAlias = await dbContext.BlankAliases.FirstOrDefaultAsync(x => x.OneCCode == code);
        if (existingAlias is not null)
        {
            existingAlias.IsActive = true;
            existingAlias.SourceName = name;
            existingAlias.NormalizedSourceName = NormalizeKey(name);
            existingAlias.UpdatedAt = DateTime.UtcNow;
            StatusText = $"Позиция НСИ {code} уже была в базе, данные обновлены.";
            await dbContext.SaveChangesAsync();
            await LoadAsync();
            SelectedRow = Rows.FirstOrDefault(x => x.AliasId == existingAlias.Id);
            return;
        }

        var canonical = new CanonicalBlank
        {
            CanonicalName = name,
            CanonicalKey = $"manual:{code.ToUpperInvariant()}",
            BlankType = EditBlankType.Value,
            BaseUnit = EditUnit.Value,
            Material = NullIfWhiteSpace(EditMaterial),
            MaterialGost = NullIfWhiteSpace(EditMaterialGost),
            ProfileGost = NullIfWhiteSpace(EditProfileGost)
        };

        if (!TryApplySize(EditSize, canonical, out var sizeError))
        {
            StatusText = sizeError;
            return;
        }

        var alias = new BlankAlias
        {
            CanonicalBlank = canonical,
            OneCCode = code,
            SourceName = name,
            NormalizedSourceName = NormalizeKey(name),
            Source = "Ручной ввод НСИ",
            IsActive = true
        };

        dbContext.CanonicalBlanks.Add(canonical);
        dbContext.BlankAliases.Add(alias);
        await dbContext.SaveChangesAsync();
        StatusText = $"Добавлена позиция НСИ: {code}.";
        await LoadAsync();
        SelectedRow = Rows.FirstOrDefault(x => x.AliasId == alias.Id);
    }

    [RelayCommand]
    private async Task ArchiveSelectedRowsAsync(object? parameter)
    {
        var rows = GetSelectedNsiRows(parameter).DistinctBy(x => x.AliasId).ToArray();
        if (rows.Length == 0)
        {
            StatusText = "Выберите одну или несколько строк НСИ для переноса в архив.";
            return;
        }

        suppressAutoRestoreNsi = true;
        var aliasIds = rows.Select(x => x.AliasId).Distinct().ToArray();
        var aliases = await dbContext.BlankAliases
            .Where(x => aliasIds.Contains(x.Id))
            .ToListAsync();

        foreach (var alias in aliases)
        {
            alias.IsActive = false;
            alias.UpdatedAt = DateTime.UtcNow;
        }

        await dbContext.SaveChangesAsync();
        StatusText = $"Перенесено в архив позиций НСИ: {aliases.Count}.";
        await LoadAsync();
    }

    private IEnumerable<NsiBlankRow> GetSelectedNsiRows(object? parameter)
    {
        if (parameter is IEnumerable selectedItems)
        {
            foreach (var item in selectedItems)
            {
                if (item is NsiBlankRow row)
                {
                    yield return row;
                }
                else if (item is DataGridCellInfo cell && cell.Item is NsiBlankRow cellRow)
                {
                    yield return cellRow;
                }
            }
        }
        else if (parameter is NsiBlankRow row)
        {
            yield return row;
        }
        else if (parameter is DataGridCellInfo cell && cell.Item is NsiBlankRow cellRow)
        {
            yield return cellRow;
        }

        if (SelectedRow is not null)
        {
            yield return SelectedRow;
        }
    }

    private async Task EnsureNsiRowsAvailableAsync()
    {
        if (suppressAutoRestoreNsi || await dbContext.BlankAliases.AnyAsync(x => x.IsActive))
        {
            return;
        }

        var inactiveAliases = await dbContext.BlankAliases.Where(x => !x.IsActive).ToListAsync();
        if (inactiveAliases.Count > 0)
        {
            foreach (var alias in inactiveAliases)
            {
                alias.IsActive = true;
                alias.UpdatedAt = DateTime.UtcNow;
            }

            StatusText = $"Восстановлено позиций НСИ из архива: {inactiveAliases.Count}.";
            await dbContext.SaveChangesAsync();
            return;
        }

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        if (latestSnapshotId is null)
        {
            return;
        }

        var stockItems = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.StockSnapshotId == latestSnapshotId.Value && x.OneCCode != "" && x.SourceName != "")
            .OrderBy(x => x.OneCCode)
            .ToListAsync();
        var restored = 0;
        foreach (var stock in stockItems.GroupBy(x => x.OneCCode, StringComparer.OrdinalIgnoreCase).Select(x => x.First()))
        {
            var code = stock.OneCCode.Trim();
            var name = stock.SourceName.Trim();
            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var canonicalKey = $"stock:{code.ToUpperInvariant()}";
            var canonical = await dbContext.CanonicalBlanks.FirstOrDefaultAsync(x => x.CanonicalKey == canonicalKey);
            if (canonical is null)
            {
                canonical = new CanonicalBlank
                {
                    CanonicalName = name,
                    CanonicalKey = canonicalKey,
                    BlankType = BlankType.Unknown,
                    BaseUnit = stock.Unit
                };
                dbContext.CanonicalBlanks.Add(canonical);
            }

            dbContext.BlankAliases.Add(new BlankAlias
            {
                CanonicalBlank = canonical,
                OneCCode = code,
                SourceName = name,
                NormalizedSourceName = NormalizeKey(name),
                Source = "Восстановлено из остатков"
            });
            restored++;
        }

        if (restored > 0)
        {
            StatusText = $"Восстановлено позиций НСИ из последнего снимка остатков: {restored}.";
            await dbContext.SaveChangesAsync();
        }
    }
    [RelayCommand]
    public async Task LoadAsync()
    {
        await EnsureNsiRowsAvailableAsync();

        var query = dbContext.BlankAliases.AsNoTracking()
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive);

        if (!string.IsNullOrWhiteSpace(Search))
        {
            var searchValue = Search.Trim();
            query = query.Where(x =>
                x.OneCCode.Contains(searchValue) ||
                x.SourceName.Contains(searchValue) ||
                x.NormalizedSourceName.Contains(searchValue) ||
                (x.CanonicalBlank != null && (
                    x.CanonicalBlank.CanonicalName.Contains(searchValue) ||
                    (x.CanonicalBlank.Material != null && x.CanonicalBlank.Material.Contains(searchValue)))));
        }

        if (SelectedBlankTypeFilter.Value is not null)
        {
            var blankType = SelectedBlankTypeFilter.Value.Value;
            query = query.Where(x => x.CanonicalBlank != null && x.CanonicalBlank.BlankType == blankType);
        }

        if (!string.IsNullOrWhiteSpace(MaterialFilter))
        {
            var material = MaterialFilter.Trim();
            query = query.Where(x =>
                (x.CanonicalBlank != null && x.CanonicalBlank.Material != null && x.CanonicalBlank.Material.Contains(material)) ||
                x.SourceName.Contains(material) ||
                x.NormalizedSourceName.Contains(material));
        }

        if (!string.IsNullOrWhiteSpace(SizeFilter))
        {
            var sizeText = SizeFilter.Trim();
            if (TryParseQuantity(sizeText, out var size))
            {
                query = query.Where(x => x.CanonicalBlank != null &&
                    (x.CanonicalBlank.DiameterMm == size ||
                     x.CanonicalBlank.WidthMm == size ||
                     x.CanonicalBlank.HeightMm == size ||
                     x.CanonicalBlank.ThicknessMm == size ||
                     x.CanonicalBlank.WallThicknessMm == size ||
                     x.CanonicalBlank.LengthMm == size ||
                     x.SourceName.Contains(sizeText) ||
                     x.NormalizedSourceName.Contains(sizeText)));
            }
            else
            {
                query = query.Where(x =>
                    x.SourceName.Contains(sizeText) ||
                    x.NormalizedSourceName.Contains(sizeText) ||
                    (x.CanonicalBlank != null && x.CanonicalBlank.CanonicalName.Contains(sizeText)));
            }
        }

        if (DuplicatesOnly)
        {
            var duplicateNames = await dbContext.BlankAliases.AsNoTracking()
                .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive && x.NormalizedSourceName != "")
                .GroupBy(x => x.NormalizedSourceName)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToListAsync();
            query = query.Where(x => duplicateNames.Contains(x.NormalizedSourceName));
        }

        var aliases = await query
            .OrderBy(x => x.OneCCode)
            .ToListAsync();

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        var stockItems = latestSnapshotId is null
            ? new List<StockItem>()
            : await dbContext.StockItems.AsNoTracking()
                .Where(x => x.StockSnapshotId == latestSnapshotId.Value)
                .ToListAsync();

        var cmoStockByCode = stockItems
            .Where(x => StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                x => x.Key,
                x => (Quantity: x.Sum(i => i.Quantity), Unit: x.Select(i => i.Unit).FirstOrDefault()),
                StringComparer.OrdinalIgnoreCase);
        var warehouseStockByCode = stockItems
            .Where(x => StockWarehouseRules.IsProductionWarehouse(x.Warehouse))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                x => x.Key,
                x => (Quantity: x.Sum(i => i.Quantity), Unit: x.Select(i => i.Unit).FirstOrDefault()),
                StringComparer.OrdinalIgnoreCase);

        Rows.Clear();
        foreach (var alias in aliases)
        {
            var blank = alias.CanonicalBlank;
            var unit = blank is null
                ? MeasurementUnit.Piece
                : MeterBasedBlankTypes.Contains(blank.BlankType) ? MeasurementUnit.Meter : blank.BaseUnit;
            var stockCodeKey = StockCodeNormalizer.NormalizeForComparison(alias.OneCCode);
            var hasCmoStock = cmoStockByCode.TryGetValue(stockCodeKey, out var cmoStock);
            var hasWarehouseStock = warehouseStockByCode.TryGetValue(stockCodeKey, out var warehouseStock);
            Rows.Add(new NsiBlankRow(
                alias.Id,
                alias.CanonicalBlankId,
                alias.OneCCode,
                UiText.Clean(alias.SourceName),
                DisplayUnit(unit),
                blank is null ? "Не распознано" : DisplayBlankType(blank.BlankType),
                FormatSize(blank),
                UiText.Clean(blank?.Material),
                UiText.Clean(blank?.MaterialGost),
                UiText.Clean(blank?.ProfileGost),
                hasCmoStock ? FormatDecimal(cmoStock.Quantity) : "0",
                hasWarehouseStock ? FormatDecimal(warehouseStock.Quantity) : "0",
                hasCmoStock ? DisplayUnit(cmoStock.Unit) : hasWarehouseStock ? DisplayUnit(warehouseStock.Unit) : string.Empty,
                UiText.Clean(alias.SourceName),
                UiText.Clean(alias.Source),
                alias.UpdatedAt.ToLocalTime().ToString("g")));
        }

        SummaryText = $"Строк НСИ: {Rows.Count}";
        StatusText = $"Показано позиций НСИ: {Rows.Count}";
        if (SelectedRow is not null && Rows.All(x => x.AliasId != SelectedRow.AliasId))
        {
            SelectedRow = null;
        }

        await LoadMaterialSuggestionsAsync(EditMaterial);
    }

    [RelayCommand]
    private async Task LoadUsageAsync(NsiBlankRow? row)
    {
        UsageRows.Clear();
        if (row is null)
        {
            UsageStatusText = "Выберите заготовку";
            return;
        }

        var maps = await dbContext.PartBlankMaps.AsNoTracking()
            .Include(x => x.Part)
            .Include(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .Where(x => x.IsActive && x.Part != null && (x.Part.Source == null || !x.Part.Source.Contains("[ARCHIVED_LIBRARY]")))
            .ToListAsync();

        maps = maps
            .Where(x => IsUsageMatch(x, row))
            .OrderBy(x => x.Part?.Ips)
            .ToList();

        foreach (var map in maps)
        {
            var unit = map.CanonicalBlank is not null && MeterBasedBlankTypes.Contains(map.CanonicalBlank.BlankType) && map.ConsumptionUnit == MeasurementUnit.Piece
                ? MeasurementUnit.Meter
                : map.ConsumptionUnit;
            UsageRows.Add(new NsiUsageRow(
                map.Part?.Ips ?? string.Empty,
                UiText.Clean(map.Part?.Designation),
                UiText.Clean(map.Part?.Name),
                FormatDecimal(map.ConsumptionQuantity),
                DisplayUnit(unit),
                map.Source));
        }

        UsageStatusText = UsageRows.Count == 0
            ? "Применяемость не найдена"
            : $"Деталей: {UsageRows.Count}";
    }

    private static string FormatSize(CanonicalBlank? blank)
    {
        if (blank is null)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        if (blank.DiameterMm is not null)
        {
            parts.Add($"D{FormatDecimal(blank.DiameterMm)}");
        }

        if (blank.WidthMm is not null)
        {
            parts.Add($"W{FormatDecimal(blank.WidthMm)}");
        }

        if (blank.HeightMm is not null)
        {
            parts.Add($"H{FormatDecimal(blank.HeightMm)}");
        }

        if (blank.ThicknessMm is not null)
        {
            parts.Add($"T{FormatDecimal(blank.ThicknessMm)}");
        }

        if (blank.WallThicknessMm is not null)
        {
            parts.Add($"S{FormatDecimal(blank.WallThicknessMm)}");
        }

        if (blank.LengthMm is not null)
        {
            parts.Add($"L{FormatDecimal(blank.LengthMm)}");
        }

        return string.Join(" ", parts);
    }

    private void LoadSelectedEditor(NsiBlankRow? row)
    {
        if (row is null)
        {
            EditBlankType = BlankTypes.First(x => x.Value == BlankType.Unknown);
            EditMaterial = string.Empty;
            EditSize = string.Empty;
            EditMaterialGost = string.Empty;
            EditProfileGost = string.Empty;
            return;
        }

        EditBlankType = BlankTypes.FirstOrDefault(x => x.DisplayName == row.BlankType) ?? BlankTypes.First(x => x.Value == BlankType.Unknown);
        EditUnit = UnitTypes.FirstOrDefault(x => x.DisplayName == row.UnitName) ?? UnitTypes[0];
        EditOneCCode = row.OneCCode;
        EditSourceName = row.SourceName;
        EditMaterial = row.Material;
        EditSize = row.Size;
        EditMaterialGost = row.MaterialGost;
        EditProfileGost = row.ProfileGost;
    }

    private static bool TryApplySize(string? value, CanonicalBlank blank, out string error)
    {
        error = string.Empty;
        var text = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            ClearSize(blank);
            return true;
        }

        var normalized = text
            .Replace("С…", "x", StringComparison.Ordinal)
            .Replace("Р Тђ", "x", StringComparison.Ordinal)
            .Replace("Р“вЂ”", "x", StringComparison.Ordinal)
            .Replace("Р В Р’В Р РЋР’ВР В Р’В Р РЋР’В", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();

        var diameter = Regex.Match(normalized, @"(?:^|\s)D\s*(\d+(?:[\.,]\d+)?)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (diameter.Success)
        {
            ClearSize(blank);
            blank.DiameterMm = ParseRuDecimal(diameter.Groups[1].Value);
            ApplyNamedSizeTokens(normalized, blank);
            return true;
        }

        if (TryApplyNamedSizeTokens(normalized, blank))
        {
            return true;
        }

        var dimensionMatches = Regex.Matches(normalized, @"\d+(?:[\.,]\d+)?", RegexOptions.CultureInvariant)
            .Select(x => ParseRuDecimal(x.Value))
            .ToArray();

        if (normalized.Contains('x', StringComparison.Ordinal) && dimensionMatches.Length is 2 or 3)
        {
            ClearSize(blank);
            blank.WidthMm = dimensionMatches[0];
            blank.HeightMm = dimensionMatches[1];
            if (dimensionMatches.Length == 3)
            {
                blank.LengthMm = dimensionMatches[2];
                if (blank.BlankType is not BlankType.PipeRectangular and not BlankType.PipeRound)
                {
                    blank.BlankType = BlankType.Sheet;
                }
            }

            return true;
        }

        if (dimensionMatches.Length == 1)
        {
            ClearSize(blank);
            if (blank.BlankType is BlankType.RoundBar or BlankType.PipeRound or BlankType.BronzeBar)
            {
                blank.DiameterMm = dimensionMatches[0];
            }
            else
            {
                blank.WidthMm = dimensionMatches[0];
            }

            return true;
        }

        error = "Размер не распознан. Используйте формат D40, 40, 40x100 или 40x100x1030.";
        return false;
    }

    private static bool TryApplyNamedSizeTokens(string normalized, CanonicalBlank blank)
    {
        if (!Regex.IsMatch(normalized, @"(?:^|\s)[DWHTSL]\s*\d", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            return false;
        }

        ClearSize(blank);
        ApplyNamedSizeTokens(normalized, blank);
        return blank.DiameterMm is not null ||
               blank.WidthMm is not null ||
               blank.HeightMm is not null ||
               blank.ThicknessMm is not null ||
               blank.WallThicknessMm is not null ||
               blank.LengthMm is not null;
    }

    private static void ApplyNamedSizeTokens(string normalized, CanonicalBlank blank)
    {
        foreach (Match match in Regex.Matches(normalized, @"(?:^|\s)([DWHTSL])\s*(\d+(?:[\.,]\d+)?)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
        {
            var value = ParseRuDecimal(match.Groups[2].Value);
            switch (match.Groups[1].Value.ToUpperInvariant())
            {
                case "D":
                    blank.DiameterMm = value;
                    break;
                case "W":
                    blank.WidthMm = value;
                    break;
                case "H":
                    blank.HeightMm = value;
                    break;
                case "T":
                    blank.ThicknessMm = value;
                    break;
                case "S":
                    blank.WallThicknessMm = value;
                    break;
                case "L":
                    blank.LengthMm = value;
                    break;
            }
        }
    }

    private static void ClearSize(CanonicalBlank blank)
    {
        blank.DiameterMm = null;
        blank.WidthMm = null;
        blank.HeightMm = null;
        blank.ThicknessMm = null;
        blank.WallThicknessMm = null;
        blank.LengthMm = null;
    }

    private static decimal ParseRuDecimal(string value) =>
        decimal.Parse(value.Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture);

    private static string? NullIfWhiteSpace(string? value) => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static bool IsUsageMatch(PartBlankMap map, NsiBlankRow row)
    {
        return map.CanonicalBlankId == row.CanonicalBlankId;
    }

    private static string NormalizeKey(string? value) => UiText.Clean(value)
        .ToUpperInvariant()
        .Replace(" ", string.Empty, StringComparison.Ordinal)
        .Replace("РњРњ", string.Empty, StringComparison.Ordinal)
        .Replace("Р“вЂ”", "Р Тђ", StringComparison.Ordinal)
        .Replace("ММ", string.Empty, StringComparison.Ordinal)
        .Replace("Х", "X", StringComparison.Ordinal)
        .Replace("С…", "X", StringComparison.Ordinal)
        .Replace("×", "X", StringComparison.Ordinal);

    private static string FormatDecimal(decimal? quantity) => quantity is null ? string.Empty : quantity.Value.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    private static bool TryParseQuantity(string? value, out decimal quantity)
    {
        var normalized = value?.Trim().Replace('.', ',') ?? string.Empty;
        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out quantity);
    }

    private static string DisplayUnit(MeasurementUnit unit) => unit switch
    {
        MeasurementUnit.Meter => "пог. м",
        MeasurementUnit.Piece => "шт",
        MeasurementUnit.Kilogram => "кг",
        _ => unit.ToString()
    };

    private static string DisplayBlankType(BlankType type) => type switch
    {
        BlankType.RoundBar => "Круг",
        BlankType.SquareBar => "Квадрат",
        BlankType.HexBar => "Шестигранник",
        BlankType.Sheet => "Лист",
        BlankType.Plate => "Плита",
        BlankType.PipeRound => "Труба профильная круглая",
        BlankType.PipeRectangular => "Труба профильная прямоугольная",
        BlankType.Angle => "Уголок",
        BlankType.Channel => "Швеллер",
        BlankType.IBeam => "Двутавр",
        BlankType.BronzeBar => "Пруток бронзовый",
        BlankType.BronzeSheet => "Лист бронзовый",
        BlankType.WeldingElement => "Сварочный элемент",
        BlankType.Purchased => "Покупная",
        BlankType.Casting => "Литье",
        BlankType.Forging => "Поковка",
        BlankType.CustomBlank => "Не распознано",
        BlankType.Unknown => "Не распознано",
        _ => type.ToString()
    };
}

public sealed partial class StockViewModel(BlankDemandPlannerDbContext dbContext) : ObservableObject
{
    public ObservableCollection<StockRow> Rows { get; } = [];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private string statusText = "Остатки не загружены";
    [ObservableProperty] private string snapshotText = "Остаток УТ на дату: нет данных";

    partial void OnSearchChanged(string value) => _ = LoadAsync();

    [RelayCommand]
    public async Task LoadAsync()
    {
        var snapshot = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .FirstOrDefaultAsync();

        Rows.Clear();
        if (snapshot is null)
        {
            SnapshotText = "Остаток УТ на дату: нет данных";
            StatusText = "Остатки не загружены";
            return;
        }

        SnapshotText = $"Остаток УТ на дату: {snapshot.SnapshotDate.ToLocalTime():dd.MM.yyyy}";

        var query = dbContext.StockItems.AsNoTracking()
            .Where(x => x.StockSnapshotId == snapshot.Id);

        if (!string.IsNullOrWhiteSpace(Search))
        {
            var searchValue = Search.Trim();
            query = query.Where(x =>
                x.OneCCode.Contains(searchValue) ||
                x.SourceName.Contains(searchValue) ||
                (x.Warehouse != null && x.Warehouse.Contains(searchValue)));
        }

        var items = await query
            .OrderBy(x => x.OneCCode)
            .Take(2000)
            .ToListAsync();

        foreach (var item in items)
        {
            Rows.Add(new StockRow(
                item.OneCCode,
                UiText.Clean(item.SourceName),
                FormatDecimal(item.Quantity),
                UiText.DisplayUnit(item.Unit),
                UiText.Clean(item.Warehouse)));
        }

        StatusText = $"Показано остатков: {Rows.Count}; источник: {Path.GetFileName(snapshot.SourceFile ?? string.Empty)}";
    }

    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
}

public sealed partial class MskViewModel(
    BlankDemandPlannerDbContext dbContext,
    ILogger logger,
    IIpsDrawingService? ipsDrawingService = null,
    bool autoOpenDrawings = false) : ObservableObject
{
    private const string DefaultMskFolder = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК";

    private readonly List<MskLibraryRow> allRows = [];
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();
    private readonly HashSet<string> attemptedDrawingIps = new(StringComparer.OrdinalIgnoreCase);

    public ObservableCollection<MskLibraryRow> Rows { get; } = [];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private string statusText = "МСК не загружены.";
    [ObservableProperty] private MskLibraryRow? selectedRow;
    [ObservableProperty] private string detailText = "Выберите деталь для просмотра МСК.";
    [ObservableProperty] private bool isBusy;

    partial void OnSearchChanged(string value) => ApplyFilter();

    partial void OnSelectedRowChanged(MskLibraryRow? value)
    {
        DetailText = value is null
            ? "Выберите деталь для просмотра МСК."
            : BuildDetailText(value);
        if (value is not null && autoOpenDrawings)
        {
            _ = TryOpenDrawingAsync(value, forceRetry: false);
        }
    }

    [RelayCommand]
    public async Task LoadAsync()
    {
        await LoadFromDatabaseAsync();
    }

    [RelayCommand]
    private async Task RefreshFromFolderAsync()
    {
        if (IsBusy)
        {
            return;
        }

        IsBusy = true;
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (!Directory.Exists(DefaultMskFolder))
            {
                StatusText = $"Папка МСК недоступна: {DefaultMskFolder}";
                return;
            }

            var files = Directory.EnumerateFiles(DefaultMskFolder, "*.xls*", SearchOption.AllDirectories)
                .Where(IsSupportedExcelFile)
                .OrderBy(x => x)
                .ToList();
            var records = new List<MskRecord>();
            var errors = 0;
            foreach (var file in files)
            {
                try
                {
                    var record = ReadMskRecord(file);
                    if (!string.IsNullOrWhiteSpace(record.Ips))
                    {
                        records.Add(record);
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    logger.LogWarning(ex, "Could not read MSK file {File}", file);
                }
            }

            var latestByIps = records
                .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
                .Select(x => x.OrderByDescending(r => File.Exists(r.FileName) ? File.GetLastWriteTime(r.FileName) : DateTime.MinValue).First())
                .ToList();

            var existing = await dbContext.MskRecords.ToListAsync();
            dbContext.MskRecords.RemoveRange(existing);
            await dbContext.SaveChangesAsync();

            dbContext.MskRecords.AddRange(latestByIps);

            var parts = await dbContext.Parts.ToListAsync();
            var mskIps = latestByIps.Select(x => x.Ips).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var part in parts)
            {
                part.HasMsk = mskIps.Contains(part.Ips);
                part.UpdatedAt = DateTime.UtcNow;
            }

            await dbContext.SaveChangesAsync();
            await LoadFromDatabaseAsync();
            StatusText = $"МСК обновлены из X: файлов {files.Count}, записей {latestByIps.Count}, ошибок чтения {errors}.";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void OpenSelectedMsk()
    {
        var row = SelectedRow ?? ResolveRowFromSearch();
        if (row is null || string.IsNullOrWhiteSpace(row.FilePath))
        {
            StatusText = "Выберите строку с файлом МСК.";
            return;
        }

        if (!File.Exists(row.FilePath))
        {
            StatusText = $"Файл МСК не найден: {row.FilePath}";
            return;
        }

        SelectedRow = row;
        Process.Start(new ProcessStartInfo(row.FilePath) { UseShellExecute = true });
        StatusText = $"Открыт файл МСК: {row.FilePath}";
    }

    [RelayCommand]
    private async Task OpenDrawingAsync(object? parameter)
    {
        var row = parameter as MskLibraryRow ?? SelectedRow ?? ResolveRowFromSearch();
        if (row is null)
        {
            StatusText = "Выберите строку МСК для открытия чертежа.";
            return;
        }

        SelectedRow = row;
        await TryOpenDrawingAsync(row, forceRetry: true);
    }

    private async Task LoadFromDatabaseAsync()
    {
        var parts = await dbContext.Parts.AsNoTracking()
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
            .AsSplitQuery()
            .ToListAsync();
        var records = await dbContext.MskRecords.AsNoTracking().ToListAsync();
        var mskDetails = LoadMskDetailsFromReport();
        var recordsByIps = records
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(r => r.ImportedAt).First(), StringComparer.OrdinalIgnoreCase);
        var partsByIps = parts
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(p => p.UpdatedAt).First(), StringComparer.OrdinalIgnoreCase);

        allRows.Clear();
        foreach (var ips in partsByIps.Keys.Concat(recordsByIps.Keys).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x))
        {
            partsByIps.TryGetValue(ips, out var part);
            recordsByIps.TryGetValue(ips, out var record);
            var detail = ResolveMskDetail(record, mskDetails);
            var map = part?.BlankMaps.FirstOrDefault(x => x.IsActive && x.IsPrimary) ?? part?.BlankMaps.FirstOrDefault(x => x.IsActive);
            var blank = map?.CanonicalBlank;
            var alias = blank?.Aliases.FirstOrDefault(x => x.IsActive) ?? blank?.Aliases.FirstOrDefault();
            var unit = map?.ConsumptionUnit is null ? string.Empty : UiText.DisplayUnit(map.ConsumptionUnit);
            allRows.Add(new MskLibraryRow(
                ips,
                UiText.Clean(part?.Designation ?? record?.Designation),
                UiText.Clean(FirstNotEmpty(part?.Name, record?.Name)),
                record is null ? "Нет" : "Да",
                part is null ? "Нет" : "Да",
                record?.FileName ?? string.Empty,
                File.Exists(record?.FileName) ? Path.GetFileName(record!.FileName) : string.Empty,
                record?.ImportedAt.ToLocalTime().ToString("g") ?? string.Empty,
                FirstNotEmpty(blank is null ? null : UiText.DisplayBlankType(blank.BlankType), detail?.BlankType),
                FirstNotEmpty(UiText.Clean(alias?.SourceName ?? blank?.CanonicalName), detail?.BlankName),
                FirstNotEmpty(UiText.Clean(blank?.Material), detail?.Material),
                FirstNotEmpty(alias?.OneCCode, detail?.OneCCode),
                map is null ? detail?.ConsumptionQuantity ?? string.Empty : FormatDecimal(map.ConsumptionQuantity),
                FirstNotEmpty(unit, detail?.UnitName),
                (map?.BlankLeadTimeDays ?? 30).ToString(CultureInfo.InvariantCulture)));
        }

        ApplyFilter();
        StatusText = allRows.Count == 0
            ? "МСК еще не загружены. Нажмите \"Обновить данные МСК\"."
            : $"Показано {Rows.Count} из {allRows.Count}; записей МСК: {records.Count}.";
    }

    private void ApplyFilter()
    {
        var searchValue = UiText.Clean(Search).Trim();
        var filtered = string.IsNullOrWhiteSpace(searchValue)
            ? allRows
            : allRows.Where(x =>
                x.Ips.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.Designation.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.Name.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.BlankType.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.BlankName.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.Material.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.OneCCode.Contains(searchValue, StringComparison.OrdinalIgnoreCase) ||
                x.FileName.Contains(searchValue, StringComparison.OrdinalIgnoreCase)).ToList();

        Rows.Clear();
        foreach (var row in filtered.Take(1000))
        {
            Rows.Add(row);
        }

        SelectedRow = ResolvePreferredRow(searchValue);
        StatusText = $"Показано {Rows.Count} из {allRows.Count}.";
    }

    private MskLibraryRow? ResolvePreferredRow(string searchValue)
    {
        if (Rows.Count == 0)
        {
            return null;
        }

        if (Rows.Count == 1)
        {
            return Rows[0];
        }

        if (SelectedRow is not null && Rows.Contains(SelectedRow))
        {
            return SelectedRow;
        }

        return string.IsNullOrWhiteSpace(searchValue)
            ? Rows[0]
            : Rows.FirstOrDefault(x => string.Equals(x.Ips, searchValue, StringComparison.OrdinalIgnoreCase)) ?? Rows[0];
    }

    private static MskRecord ReadMskRecord(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var sheet = package.Workbook.Worksheets.FirstOrDefault()
            ?? throw new InvalidOperationException("В книге МСК нет листов.");
        var ips = CleanMskCell(sheet.Cells["C5"].Text);
        var designation = CleanMskCell(sheet.Cells["E5"].Text);
        var name = CleanMskCell(sheet.Cells["L5"].Text);

        if (string.IsNullOrWhiteSpace(ips))
        {
            ips = Path.GetFileNameWithoutExtension(filePath);
        }

        return new MskRecord
        {
            Ips = ips,
            Designation = designation,
            Name = string.IsNullOrWhiteSpace(name) ? Path.GetFileNameWithoutExtension(filePath) : name,
            FileName = filePath,
            ImportedAt = DateTime.UtcNow
        };
    }

    private static bool IsSupportedExcelFile(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static string CleanMskCell(string? value) => UiText.Clean(value).Trim();

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

    private static string BuildDetailText(MskLibraryRow row) =>
        $"IPS: {row.Ips}\n" +
        $"Обозначение: {row.Designation}\n" +
        $"Наименование: {row.Name}\n" +
        $"Вид заготовки: {row.BlankType}\n" +
        $"Заготовка: {row.BlankName}\n" +
        $"Материал: {row.Material}\n" +
        $"Код УТ: {row.OneCCode}\n" +
        $"Норма расхода: {row.ConsumptionQuantity} {row.UnitName}\n" +
        $"Срок заготовки, дней: {row.BlankLeadTimeDays}\n" +
        $"Есть в библиотеке: {row.HasLibraryPart}\n" +
        $"Есть МСК: {row.HasMsk}\n" +
        $"Файл: {row.FilePath}";

    private async Task TryOpenDrawingAsync(MskLibraryRow row, bool forceRetry)
    {
        if (string.IsNullOrWhiteSpace(row.Ips))
        {
            return;
        }

        if (!forceRetry && !attemptedDrawingIps.Add(row.Ips))
        {
            return;
        }

        if (forceRetry)
        {
            attemptedDrawingIps.Add(row.Ips);
        }

        try
        {
            var query = FirstNotEmpty(row.Ips, row.Designation, row.Name);
            var outputDirectory = GetDrawingCacheDirectory();
            var drawing = await drawingService.FindDrawingPdfAsync(query, outputDirectory, CancellationToken.None);
            if (drawing is null)
            {
                StatusText = $"PDF-чертеж IPS {row.Ips} в IPS Bridge не найден.";
                return;
            }

            Process.Start(new ProcessStartInfo(drawing.FullName) { UseShellExecute = true });
            StatusText = $"Открыт PDF-чертеж IPS {row.Ips}: {drawing.Name}";
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not open IPS drawing for {Ips}", row.Ips);
            StatusText = $"PDF-чертеж IPS {row.Ips} не открыт: {ex.GetBaseException().Message}";
        }
    }

    private static DirectoryInfo GetDrawingCacheDirectory()
    {
        var path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BlankDemandPlanner",
            "IpsDrawings");
        return Directory.CreateDirectory(path);
    }

    private MskLibraryRow? ResolveRowFromSearch()
    {
        if (Rows.Count == 1)
        {
            return Rows[0];
        }

        var ips = UiText.Clean(Search).Trim();
        return string.IsNullOrWhiteSpace(ips)
            ? null
            : Rows.FirstOrDefault(x => string.Equals(x.Ips, ips, StringComparison.OrdinalIgnoreCase));
    }

    private static MskCsvDetail? ResolveMskDetail(MskRecord? record, IReadOnlyDictionary<string, MskCsvDetail> details)
    {
        if (record is null)
        {
            return null;
        }

        return details.TryGetValue(record.FileName, out var byFile)
            ? byFile
            : details.TryGetValue($"IPS:{record.Ips}", out var byIps) ? byIps : null;
    }

    private static IReadOnlyDictionary<string, MskCsvDetail> LoadMskDetailsFromReport()
    {
        var path = FindWorkspaceFile(Path.Combine("reports", "msk_excel_analysis_20260723", "msk_library_extract.csv"));
        if (path is null || !File.Exists(path))
        {
            return new Dictionary<string, MskCsvDetail>(StringComparer.OrdinalIgnoreCase);
        }

        var result = new Dictionary<string, MskCsvDetail>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in File.ReadLines(path, Encoding.UTF8).Skip(1))
        {
            var fields = SplitSemicolonCsv(line).ToArray();
            if (fields.Length < 13)
            {
                continue;
            }

            var detail = new MskCsvDetail(
                Source: fields[0],
                Ips: fields[1],
                BlankType: fields[4],
                BlankName: fields[5],
                Material: fields[7],
                OneCCode: fields[12],
                ConsumptionQuantity: fields[10],
                UnitName: fields[11]);
            if (!string.IsNullOrWhiteSpace(detail.Source))
            {
                result[detail.Source] = detail;
            }

            if (!string.IsNullOrWhiteSpace(detail.Ips))
            {
                result[$"IPS:{detail.Ips}"] = detail;
            }
        }

        return result;
    }

    private static IEnumerable<string> SplitSemicolonCsv(string line)
    {
        var field = new StringBuilder();
        var inQuotes = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (ch == ';' && !inQuotes)
            {
                yield return field.ToString();
                field.Clear();
            }
            else
            {
                field.Append(ch);
            }
        }

        yield return field.ToString();
    }

    private static string? FindWorkspaceFile(string relativePath)
    {
        foreach (var basePath in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            var directory = new DirectoryInfo(basePath);
            for (var i = 0; directory is not null && i < 10; i++, directory = directory.Parent)
            {
                var path = Path.Combine(directory.FullName, relativePath);
                if (File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }

    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
}

public interface IIpsDrawingService
{
    Task<FileInfo?> FindDrawingPdfAsync(string query, DirectoryInfo outputDirectory, CancellationToken cancellationToken);
}

public sealed class IpsBridgeDrawingService : IIpsDrawingService
{
    private static readonly string[] RequiredTools =
    [
        "login",
        "search_by_string",
        "get_object",
        "get_structure",
        "export_file_attribute",
        "read_staged_file_chunk"
    ];

    public async Task<FileInfo?> FindDrawingPdfAsync(string query, DirectoryInfo outputDirectory, CancellationToken cancellationToken)
    {
        var settings = IpsBridgeSettings.Load();
        if (!settings.IsConfigured)
        {
            return null;
        }

        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds) };
        var mcpSessionId = await InitializeAsync(http, settings, cancellationToken);
        var tools = await ListToolsAsync(http, settings, mcpSessionId, cancellationToken);
        if (RequiredTools.Any(tool => !tools.ContainsKey(tool)))
        {
            return null;
        }

        var ipsSessionId = await LoginAsync(http, settings, mcpSessionId, tools, cancellationToken);
        try
        {
            var objects = await SearchObjectsAsync(http, settings, mcpSessionId, tools, ipsSessionId, query, cancellationToken);
            foreach (var objectId in await GetCandidateObjectIdsAsync(http, settings, mcpSessionId, tools, ipsSessionId, objects, cancellationToken))
            {
                using var details = await CallToolAsync(http, settings, mcpSessionId, tools["get_object"], new Dictionary<string, object?>
                {
                    ["session_id"] = ipsSessionId,
                    ["object_id"] = objectId
                }, cancellationToken);
                var payload = ExtractPayload(details);
                if (payload.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                var pdfAttribute = FindPdfAttribute(payload);
                if (pdfAttribute is null)
                {
                    continue;
                }

                var exportObjectId = ReadInt(payload, "objectId") ?? ReadInt(payload, "versionId") ?? objectId;
                var (content, fileName) = await ExportFileAttributeAsync(
                    http,
                    settings,
                    mcpSessionId,
                    tools,
                    ipsSessionId,
                    exportObjectId,
                    pdfAttribute.AttributeId,
                    pdfAttribute.Index,
                    cancellationToken);
                if (content.Length > 4 && content[0] == '%' && content[1] == 'P' && content[2] == 'D' && content[3] == 'F')
                {
                    outputDirectory.Create();
                    var safeName = MakeSafePdfFileName(string.IsNullOrWhiteSpace(fileName) ? $"{query}.pdf" : fileName);
                    var path = Path.Combine(outputDirectory.FullName, safeName);
                    await File.WriteAllBytesAsync(path, content, cancellationToken);
                    return new FileInfo(path);
                }
            }

            return null;
        }
        finally
        {
            if (tools.TryGetValue("logout", out var logout))
            {
                try
                {
                    await CallToolAsync(http, settings, mcpSessionId, logout, new Dictionary<string, object?>
                    {
                        ["session_id"] = ipsSessionId
                    }, cancellationToken);
                }
                catch
                {
                    // Logout is best-effort: the bridge will expire the session.
                }
            }
        }
    }

    private static async Task<string?> InitializeAsync(HttpClient http, IpsBridgeSettings settings, CancellationToken cancellationToken)
    {
        var payload = new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = 1,
            ["method"] = "initialize",
            ["params"] = new Dictionary<string, object?>
            {
                ["protocolVersion"] = "2024-11-05",
                ["capabilities"] = new Dictionary<string, object?>(),
                ["clientInfo"] = new Dictionary<string, object?> { ["name"] = "blank-demand-planner", ["version"] = "1.0" }
            }
        };
        using var response = await PostAsync(http, settings, payload, null, cancellationToken);
        var sessionId = response.Headers.TryGetValues("mcp-session-id", out var values)
            ? values.FirstOrDefault()
            : response.Headers.TryGetValues("Mcp-Session-Id", out var altValues) ? altValues.FirstOrDefault() : null;
        using var _ = await PostAsync(http, settings, new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["method"] = "notifications/initialized",
            ["params"] = new Dictionary<string, object?>()
        }, sessionId, cancellationToken);
        return sessionId;
    }

    private static async Task<Dictionary<string, JsonElement>> ListToolsAsync(HttpClient http, IpsBridgeSettings settings, string? sessionId, CancellationToken cancellationToken)
    {
        using var document = await RpcJsonAsync(http, settings, new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = 2,
            ["method"] = "tools/list",
            ["params"] = new Dictionary<string, object?>()
        }, sessionId, cancellationToken);
        var result = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        if (document.RootElement.TryGetProperty("result", out var resultElement) &&
            resultElement.TryGetProperty("tools", out var toolsElement) &&
            toolsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var tool in toolsElement.EnumerateArray())
            {
                var name = ReadString(tool, "name");
                if (!string.IsNullOrWhiteSpace(name))
                {
                    result[name] = tool.Clone();
                }
            }
        }

        return result;
    }

    private static async Task<string> LoginAsync(HttpClient http, IpsBridgeSettings settings, string? sessionId, IReadOnlyDictionary<string, JsonElement> tools, CancellationToken cancellationToken)
    {
        var args = new Dictionary<string, object?>
        {
            ["username"] = settings.Username,
            ["password"] = settings.Password
        };
        if (settings.RoleId is not null)
        {
            args["role_id"] = settings.RoleId.Value;
        }
        else if (!string.IsNullOrWhiteSpace(settings.RoleName))
        {
            args["role_name"] = settings.RoleName;
        }

        using var document = await CallToolAsync(http, settings, sessionId, tools["login"], args, cancellationToken);
        var payload = ExtractPayload(document);
        var session = FindString(payload, "session_id") ?? FindString(payload, "sessionId") ?? FindString(payload, "id");
        return string.IsNullOrWhiteSpace(session)
            ? throw new InvalidOperationException("IPS Bridge login не вернул session_id.")
            : session;
    }

    private static async Task<List<JsonElement>> SearchObjectsAsync(HttpClient http, IpsBridgeSettings settings, string? mcpSessionId, IReadOnlyDictionary<string, JsonElement> tools, string ipsSessionId, string query, CancellationToken cancellationToken)
    {
        using var document = await CallToolAsync(http, settings, mcpSessionId, tools["search_by_string"], new Dictionary<string, object?>
        {
            ["session_id"] = ipsSessionId,
            ["query"] = query,
            ["result_limit"] = 5
        }, cancellationToken);
        var payload = ExtractPayload(document);
        return payload.ValueKind == JsonValueKind.Object &&
            payload.TryGetProperty("objects", out var objects) &&
            objects.ValueKind == JsonValueKind.Array
            ? objects.EnumerateArray().Select(x => x.Clone()).ToList()
            : [];
    }

    private static async Task<List<int>> GetCandidateObjectIdsAsync(HttpClient http, IpsBridgeSettings settings, string? mcpSessionId, IReadOnlyDictionary<string, JsonElement> tools, string ipsSessionId, IReadOnlyList<JsonElement> objects, CancellationToken cancellationToken)
    {
        var ids = new List<int>();
        foreach (var obj in objects)
        {
            AddUnique(ids, ReadInt(obj, "versionId"));
            AddUnique(ids, ReadInt(obj, "masterId"));
            var objectId = ReadInt(obj, "versionId") ?? ReadInt(obj, "masterId");
            if (objectId is null)
            {
                continue;
            }

            try
            {
                using var structure = await CallToolAsync(http, settings, mcpSessionId, tools["get_structure"], new Dictionary<string, object?>
                {
                    ["session_id"] = ipsSessionId,
                    ["object_id"] = objectId.Value,
                    ["rel_type_id"] = -1,
                    ["direction"] = 0
                }, cancellationToken);
                var payload = ExtractPayload(structure);
                if (payload.ValueKind == JsonValueKind.Object &&
                    payload.TryGetProperty("items", out var items) &&
                    items.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in items.EnumerateArray())
                    {
                        AddUnique(ids, ReadInt(item, "childObjectId"));
                        AddUnique(ids, ReadInt(item, "childId"));
                        AddUnique(ids, ReadInt(item, "-2"));
                        AddUnique(ids, ReadInt(item, "-3"));
                    }
                }
            }
            catch
            {
                // Not every object has a readable structure; continue with direct candidates.
            }
        }

        return ids;
    }

    private static PdfAttribute? FindPdfAttribute(JsonElement details)
    {
        if (!details.TryGetProperty("attributeDetails", out var attributes) || attributes.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        foreach (var attribute in attributes.EnumerateObject())
        {
            var detail = attribute.Value;
            if (detail.ValueKind != JsonValueKind.Object ||
                !string.Equals(ReadString(detail, "dataType"), "ftfile", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var attributeId = ReadInt(detail, "attributeId");
            if (attributeId is null)
            {
                continue;
            }

            var names = new List<string>();
            AddStrings(names, detail, "descriptions");
            AddStrings(names, detail, "values");
            var single = ReadString(detail, "description") ?? ReadString(detail, "value");
            if (!string.IsNullOrWhiteSpace(single))
            {
                names.Add(single);
            }

            if (names.Count == 0)
            {
                names.Add(string.Empty);
            }

            for (var index = 0; index < names.Count; index++)
            {
                var name = names[index];
                if (name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("pdf", StringComparison.OrdinalIgnoreCase) ||
                    name.Contains("черт", StringComparison.OrdinalIgnoreCase))
                {
                    return new PdfAttribute(attributeId.Value, index);
                }
            }

            if (attribute.Name.Contains("черт", StringComparison.OrdinalIgnoreCase))
            {
                return new PdfAttribute(attributeId.Value, 0);
            }
        }

        return null;
    }

    private static async Task<(byte[] Content, string? FileName)> ExportFileAttributeAsync(HttpClient http, IpsBridgeSettings settings, string? mcpSessionId, IReadOnlyDictionary<string, JsonElement> tools, string ipsSessionId, int objectId, int attributeId, int index, CancellationToken cancellationToken)
    {
        using var exported = await CallToolAsync(http, settings, mcpSessionId, tools["export_file_attribute"], new Dictionary<string, object?>
        {
            ["session_id"] = ipsSessionId,
            ["object_id"] = objectId,
            ["attribute_id"] = attributeId,
            ["index"] = index,
            ["content_type"] = "application/pdf"
        }, cancellationToken);
        var payload = ExtractPayload(exported);
        var fileInfo = payload;
        if (payload.ValueKind == JsonValueKind.Object &&
            payload.TryGetProperty("files", out var files) &&
            files.ValueKind == JsonValueKind.Array &&
            files.GetArrayLength() > 0)
        {
            fileInfo = files[0];
        }

        var fileId = ReadString(fileInfo, "fileId") ?? ReadString(fileInfo, "file_id");
        var token = ReadString(fileInfo, "token");
        var fileName = ReadString(fileInfo, "fileName") ?? ReadString(fileInfo, "filename");
        if (string.IsNullOrWhiteSpace(fileId) || string.IsNullOrWhiteSpace(token))
        {
            throw new InvalidOperationException("IPS Bridge export_file_attribute не вернул fileId/token.");
        }

        var chunks = new List<byte>();
        var offset = 0;
        while (true)
        {
            using var chunkDocument = await CallToolAsync(http, settings, mcpSessionId, tools["read_staged_file_chunk"], new Dictionary<string, object?>
            {
                ["file_id"] = fileId,
                ["token"] = token,
                ["offset"] = offset,
                ["size"] = 1_048_576
            }, cancellationToken);
            var chunk = ExtractPayload(chunkDocument);
            var dataBase64 = ReadString(chunk, "dataBase64") ?? ReadString(chunk, "data") ?? ReadString(chunk, "content") ?? ReadString(chunk, "base64");
            if (string.IsNullOrWhiteSpace(dataBase64))
            {
                throw new InvalidOperationException("IPS Bridge read_staged_file_chunk не вернул dataBase64.");
            }

            chunks.AddRange(Convert.FromBase64String(dataBase64));
            if (ReadBool(chunk, "eof") == true)
            {
                break;
            }

            var nextOffset = ReadInt(chunk, "nextOffset");
            if (nextOffset is null || nextOffset.Value <= offset)
            {
                break;
            }

            offset = nextOffset.Value;
        }

        return (chunks.ToArray(), fileName);
    }

    private static async Task<JsonDocument> CallToolAsync(HttpClient http, IpsBridgeSettings settings, string? sessionId, JsonElement tool, Dictionary<string, object?> arguments, CancellationToken cancellationToken)
    {
        var toolName = ReadString(tool, "name") ?? throw new InvalidOperationException("IPS Bridge вернул tool без имени.");
        return await RpcJsonAsync(http, settings, new Dictionary<string, object?>
        {
            ["jsonrpc"] = "2.0",
            ["id"] = 3,
            ["method"] = "tools/call",
            ["params"] = new Dictionary<string, object?>
            {
                ["name"] = toolName,
                ["arguments"] = arguments
            }
        }, sessionId, cancellationToken);
    }

    private static async Task<JsonDocument> RpcJsonAsync(HttpClient http, IpsBridgeSettings settings, Dictionary<string, object?> payload, string? sessionId, CancellationToken cancellationToken)
    {
        using var response = await PostAsync(http, settings, payload, sessionId, cancellationToken);
        var text = (await response.Content.ReadAsStringAsync(cancellationToken)).Trim();
        if (text.StartsWith("event:", StringComparison.OrdinalIgnoreCase) || text.Contains("\ndata:", StringComparison.OrdinalIgnoreCase))
        {
            text = ExtractSseData(text);
        }

        var document = JsonDocument.Parse(text);
        if (document.RootElement.TryGetProperty("error", out var error))
        {
            throw new InvalidOperationException($"IPS MCP error: {error}");
        }

        return document;
    }

    private static async Task<HttpResponseMessage> PostAsync(HttpClient http, IpsBridgeSettings settings, Dictionary<string, object?> payload, string? sessionId, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, settings.Url)
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
        if (!string.IsNullOrWhiteSpace(settings.Username) || !string.IsNullOrWhiteSpace(settings.Password))
        {
            var token = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{settings.Username}:{settings.Password}"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", token);
        }

        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            request.Headers.TryAddWithoutValidation("Mcp-Session-Id", sessionId);
        }

        var response = await http.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();
        return response;
    }

    private static JsonElement ExtractPayload(JsonDocument document) => ExtractPayload(document.RootElement);

    private static JsonElement ExtractPayload(JsonElement element)
    {
        var payload = element;
        if (payload.ValueKind == JsonValueKind.Object && payload.TryGetProperty("result", out var result))
        {
            payload = result;
        }

        if (payload.ValueKind == JsonValueKind.Object && payload.TryGetProperty("structuredContent", out var structured))
        {
            return structured.Clone();
        }

        var text = ExtractFirstJsonText(payload);
        if (!string.IsNullOrWhiteSpace(text))
        {
            using var parsed = JsonDocument.Parse(text);
            return parsed.RootElement.Clone();
        }

        return payload.Clone();
    }

    private static string? ExtractFirstJsonText(JsonElement element)
    {
        foreach (var nested in Walk(element))
        {
            if (nested.ValueKind == JsonValueKind.Object &&
                string.Equals(ReadString(nested, "type"), "text", StringComparison.OrdinalIgnoreCase))
            {
                var value = ReadString(nested, "text");
                if (!string.IsNullOrWhiteSpace(value) && value.TrimStart().StartsWith('{'))
                {
                    return value;
                }
            }
            else if (nested.ValueKind == JsonValueKind.String)
            {
                var value = nested.GetString();
                if (!string.IsNullOrWhiteSpace(value) && value.TrimStart().StartsWith('{'))
                {
                    return value;
                }
            }
        }

        return null;
    }

    private static IEnumerable<JsonElement> Walk(JsonElement element)
    {
        yield return element;
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                foreach (var nested in Walk(property.Value))
                {
                    yield return nested;
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                foreach (var nested in Walk(item))
                {
                    yield return nested;
                }
            }
        }
    }

    private static string? FindString(JsonElement element, string name)
    {
        foreach (var nested in Walk(element))
        {
            var value = ReadString(nested, name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static string? ReadString(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.ToString(),
            _ => null
        };
    }

    private static int? ReadInt(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
        {
            return null;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number))
        {
            return number;
        }

        return value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out number) ? number : null;
    }

    private static bool? ReadBool(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(value.GetString(), out var parsed) => parsed,
            _ => null
        };
    }

    private static void AddStrings(List<string> values, JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object ||
            !element.TryGetProperty(name, out var nested) ||
            nested.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        values.AddRange(nested.EnumerateArray().Select(x => x.ToString()));
    }

    private static void AddUnique(List<int> values, int? value)
    {
        if (value is > 0 && !values.Contains(value.Value))
        {
            values.Add(value.Value);
        }
    }

    private static string ExtractSseData(string text) => string.Join(
        "\n",
        text.Split('\n')
            .Select(x => x.TrimEnd('\r'))
            .Where(x => x.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            .Select(x => x["data:".Length..].Trim()));

    private static string MakeSafePdfFileName(string value)
    {
        var name = Regex.Replace(value, @"[^A-Za-zА-Яа-я0-9_. -]+", "_").Trim(' ', '_', '.');
        if (string.IsNullOrWhiteSpace(name))
        {
            name = "drawing";
        }

        if (name.Length > 120)
        {
            name = name[..120].Trim(' ', '_', '.');
        }

        return name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? name : $"{name}.pdf";
    }

    private sealed record PdfAttribute(int AttributeId, int Index);
}

public sealed record IpsBridgeSettings(
    string Url,
    string Username,
    string Password,
    int? RoleId,
    string? RoleName,
    double TimeoutSeconds)
{
    public bool IsConfigured => !string.IsNullOrWhiteSpace(Url);

    public static IpsBridgeSettings Load()
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in CandidateConfigFiles().Where(File.Exists))
        {
            foreach (var pair in ReadKeyValues(path))
            {
                values[pair.Key] = pair.Value;
            }
        }

        foreach (var name in new[] { "IPS_MCP_URL", "IPS_MCP_USERNAME", "IPS_MCP_PASSWORD", "IPS_MCP_ROLE_ID", "IPS_MCP_ROLE_NAME", "IPS_MCP_TIMEOUT_SECONDS" })
        {
            var value = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrWhiteSpace(value))
            {
                values[name] = value;
            }
        }

        int? roleId = null;
        if (int.TryParse(Get(values, "IPS_MCP_ROLE_ID", "ips_mcp_role_id", "role_id"), out var parsedRoleId))
        {
            roleId = parsedRoleId;
        }

        var timeout = double.TryParse(Get(values, "IPS_MCP_TIMEOUT_SECONDS", "ips_mcp_timeout_seconds", "timeout_seconds"), NumberStyles.Number, CultureInfo.InvariantCulture, out var parsedTimeout)
            ? parsedTimeout
            : 30d;
        return new IpsBridgeSettings(
            Get(values, "IPS_MCP_URL", "ips_mcp_url", "url"),
            Get(values, "IPS_MCP_USERNAME", "ips_mcp_username", "username"),
            Get(values, "IPS_MCP_PASSWORD", "ips_mcp_password", "password"),
            roleId,
            Get(values, "IPS_MCP_ROLE_NAME", "ips_mcp_role_name", "role_name"),
            Math.Clamp(timeout, 5d, 120d));
    }

    private static IEnumerable<string> CandidateConfigFiles()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "config.toml");
        yield return Path.Combine(AppContext.BaseDirectory, "ips_bridge_config.toml");
        yield return Path.Combine(AppContext.BaseDirectory, "Данные для работы", "config.toml");
        yield return Path.Combine(AppContext.BaseDirectory, "Данные для работы", "ips_bridge_config.toml");
        yield return Path.Combine(Environment.CurrentDirectory, "config.toml");
        yield return Path.Combine(Environment.CurrentDirectory, "Данные для работы", "config.toml");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Интеграция с сервисами", "factory_ai_assistant", "config.toml");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Интеграция с сервисами", "factory_ai_assistant", ".env");
    }

    private static IEnumerable<KeyValuePair<string, string>> ReadKeyValues(string path)
    {
        foreach (var line in File.ReadLines(path, Encoding.UTF8))
        {
            var trimmed = line.Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#') || trimmed.StartsWith('['))
            {
                continue;
            }

            var index = trimmed.IndexOf('=');
            if (index <= 0)
            {
                continue;
            }

            var key = trimmed[..index].Trim();
            var value = trimmed[(index + 1)..].Trim().Trim('"', '\'');
            if (!string.IsNullOrWhiteSpace(key))
            {
                yield return new KeyValuePair<string, string>(key, value);
            }
        }
    }

    private static string Get(IReadOnlyDictionary<string, string> values, params string[] names)
    {
        foreach (var name in names)
        {
            if (values.TryGetValue(name, out var value))
            {
                return value;
            }
        }

        return string.Empty;
    }
}

public sealed partial class BlankSelectionViewModel(BlankDemandPlannerDbContext dbContext) : ObservableObject
{
    private bool suppressPartSearchChanged;

    public ObservableCollection<BlankSelectionRow> Rows { get; } = [];
    public ObservableCollection<PartWithoutBlankOption> PartSuggestions { get; } = [];
    public IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes { get; } = UiText.ConsumptionUnitTypes;
    public IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFilters { get; } =
    [
        new((BlankType?)null, "Все подходящие"),
        ..UiText.BlankTypes
            .Where(x => x.Value is not BlankType.Unknown and not BlankType.CustomBlank and not BlankType.Purchased)
            .Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName))
    ];

    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankType = new(null, "Все подходящие");
    [ObservableProperty] private string partSearch = string.Empty;
    [ObservableProperty] private bool isPartSuggestionsOpen;
    [ObservableProperty] private PartWithoutBlankOption? selectedPart;
    [ObservableProperty] private BlankSelectionRow? selectedBlankRow;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiText.ConsumptionUnitTypes[0];
    [ObservableProperty] private string material = string.Empty;
    [ObservableProperty] private string diameterText = string.Empty;
    [ObservableProperty] private string widthText = string.Empty;
    [ObservableProperty] private string heightText = string.Empty;
    [ObservableProperty] private string thicknessText = string.Empty;
    [ObservableProperty] private string wallThicknessText = string.Empty;
    [ObservableProperty] private string lengthText = string.Empty;
    [ObservableProperty] private string statusText = "Введите размеры детали и нажмите \"Подобрать\".";

    partial void OnPartSearchChanged(string value)
    {
        if (suppressPartSearchChanged)
        {
            return;
        }

        SelectedPart = null;
        _ = LoadPartSuggestionsAsync(openDropDown: !string.IsNullOrWhiteSpace(value));
    }

    partial void OnSelectedPartChanged(PartWithoutBlankOption? value)
    {
        if (value is null)
        {
            return;
        }

        suppressPartSearchChanged = true;
        PartSearch = value.DisplayName;
        suppressPartSearchChanged = false;
        IsPartSuggestionsOpen = false;
        StatusText = $"Выбрана деталь без заготовки: {value.Ips}. Подберите заготовку и нажмите \"Добавить в библиотеку\".";
    }

    partial void OnSelectedBlankRowChanged(BlankSelectionRow? value)
    {
        if (value is not null)
        {
            SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == value.BaseUnit) ?? SelectedUnit;
        }
    }

    public async Task LoadAsync() => await LoadPartSuggestionsAsync(openDropDown: false);

    [RelayCommand]
    public async Task LoadPartSuggestionsAsync() => await LoadPartSuggestionsAsync(openDropDown: false);

    private async Task LoadPartSuggestionsAsync(bool openDropDown)
    {
        var search = (PartSearch ?? string.Empty).Trim();
        var query = dbContext.Parts.AsNoTracking()
            .Where(x => !x.BlankMaps.Any(m => m.IsActive));

        if (!string.IsNullOrWhiteSpace(search))
        {
            var normalized = ExtractIps(search);
            query = query.Where(x =>
                x.Ips.Contains(search) ||
                x.Ips.Contains(normalized) ||
                (x.Designation != null && x.Designation.Contains(search)) ||
                x.Name.Contains(search));
        }

        var parts = await query
            .OrderBy(x => x.Ips)
            .Take(80)
            .Select(x => new PartWithoutBlankOption(x.Id, x.Ips, x.Designation, x.Name))
            .ToListAsync();

        PartSuggestions.Clear();
        foreach (var part in parts)
        {
            PartSuggestions.Add(part);
        }

        IsPartSuggestionsOpen = openDropDown && PartSuggestions.Count > 0;
    }

    [RelayCommand]
    public async Task FindAsync()
    {
        var request = new BlankSelectionRequest(
            ParseNullable(DiameterText),
            ParseNullable(WidthText),
            ParseNullable(HeightText),
            ParseNullable(ThicknessText),
            ParseNullable(WallThicknessText),
            ParseNullable(LengthText));

        var hasSize = request.HasAnySize;
        if (!hasSize && string.IsNullOrWhiteSpace(Material) && SelectedBlankType.Value is null)
        {
            Rows.Clear();
            StatusText = "Укажите хотя бы вид, материал или один размер детали.";
            return;
        }

        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);
        if (SelectedBlankType.Value is not null)
        {
            var type = SelectedBlankType.Value.Value;
            query = query.Where(x => x.BlankType == type);
        }

        var blanks = await query
            .OrderBy(x => x.BlankType)
            .ThenBy(x => x.CanonicalName)
            .Take(20000)
            .ToListAsync();
        if (!string.IsNullOrWhiteSpace(Material))
        {
            var materialKey = NormalizeMaterial(Material);
            blanks = blanks
                .Where(x => MaterialMatches(x, materialKey))
                .ToList();
        }

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        var stockByCode = latestSnapshotId is null
            ? new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            : await dbContext.StockItems.AsNoTracking()
                .Where(x => x.StockSnapshotId == latestSnapshotId.Value)
                .GroupBy(x => x.OneCCode)
                .ToDictionaryAsync(x => x.Key, x => x.Sum(i => i.Quantity), StringComparer.OrdinalIgnoreCase);

        var matches = blanks
            .Select(blank => TryBuildMatch(blank, request, stockByCode))
            .Where(x => x is not null)
            .Select(x => x!)
            .OrderBy(x => x.AllowanceScore)
            .ThenBy(x => x.BlankType)
            .ThenBy(x => x.SourceName)
            .Take(200)
            .ToList();

        Rows.Clear();
        foreach (var row in matches)
        {
            Rows.Add(row);
        }

        SelectedBlankRow = Rows.FirstOrDefault();
        StatusText = Rows.Count == 0
            ? "Подходящие заготовки не найдены. Проверьте вид, материал и размеры."
            : $"Подобрано заготовок: {Rows.Count}. Сортировка от меньшего припуска к большему.";
    }

    [RelayCommand]
    public async Task AddToLibraryAsync()
    {
        var part = await ResolvePartAsync();
        if (part is null)
        {
            StatusText = "Выберите или введите IPS детали без заготовки.";
            return;
        }

        if (SelectedBlankRow is null)
        {
            StatusText = "Выберите заготовку из результата подбора.";
            return;
        }

        if (!decimal.TryParse((ConsumptionQuantityText ?? string.Empty).Trim().Replace('.', ','), NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out var quantity) || quantity <= 0)
        {
            StatusText = "Укажите норму расхода больше 0.";
            return;
        }

        var activeMaps = await dbContext.PartBlankMaps
            .Where(x => x.PartId == part.Id && x.IsActive)
            .ToListAsync();
        foreach (var map in activeMaps)
        {
            map.IsActive = false;
            map.IsPrimary = false;
        }

        dbContext.PartBlankMaps.Add(new PartBlankMap
        {
            PartId = part.Id,
            CanonicalBlankId = SelectedBlankRow.BlankId,
            ConsumptionQuantity = quantity,
            ConsumptionUnit = SelectedUnit.Value,
            IsPrimary = true,
            IsActive = true,
            Source = "Подбор заготовок"
        });
        await dbContext.SaveChangesAsync();
        dbContext.ChangeTracker.Clear();

        StatusText = $"Добавлено в библиотеку: {part.Ips} -> {SelectedBlankRow.OneCCode} ({FormatDecimal(quantity)} {SelectedUnit.DisplayName}).";
        SelectedPart = null;
        IsPartSuggestionsOpen = false;
        PartSearch = string.Empty;
        await LoadPartSuggestionsAsync();
    }

    [RelayCommand]
    private void Clear()
    {
        SelectedBlankType = BlankTypeFilters[0];
        PartSearch = string.Empty;
        SelectedPart = null;
        SelectedBlankRow = null;
        ConsumptionQuantityText = "1";
        SelectedUnit = UnitTypes[0];
        Material = string.Empty;
        DiameterText = string.Empty;
        WidthText = string.Empty;
        HeightText = string.Empty;
        ThicknessText = string.Empty;
        WallThicknessText = string.Empty;
        LengthText = string.Empty;
        Rows.Clear();
        StatusText = "Поля очищены.";
    }

    private async Task<Part?> ResolvePartAsync()
    {
        if (SelectedPart is not null)
        {
            return await dbContext.Parts.FirstOrDefaultAsync(x => x.Id == SelectedPart.PartId);
        }

        var ips = ExtractIps(PartSearch);
        if (string.IsNullOrWhiteSpace(ips))
        {
            return null;
        }

        return await dbContext.Parts
            .Where(x => !x.BlankMaps.Any(m => m.IsActive))
            .FirstOrDefaultAsync(x => x.Ips == ips);
    }

    private static BlankSelectionRow? TryBuildMatch(CanonicalBlank blank, BlankSelectionRequest request, IReadOnlyDictionary<string, decimal> stockByCode)
    {
        var comparisons = BuildComparisons(blank, request).ToArray();
        if (request.HasAnySize && comparisons.Length == 0)
        {
            return null;
        }

        if (comparisons.Any(x => x.BlankValue < x.RequiredValue))
        {
            return null;
        }

        var alias = blank.Aliases.FirstOrDefault(x => x.IsActive) ?? blank.Aliases.FirstOrDefault();
        var stock = blank.Aliases
            .Where(x => !string.IsNullOrWhiteSpace(x.OneCCode) && stockByCode.TryGetValue(x.OneCCode, out _))
            .Sum(x => stockByCode.TryGetValue(x.OneCCode, out var quantity) ? quantity : 0m);
        var allowance = comparisons.Sum(x => x.BlankValue - x.RequiredValue);
        var details = comparisons.Length == 0
            ? "Размеры не заданы, показано по виду/материалу"
            : string.Join("; ", comparisons.Select(x => $"{x.Label}: {FormatDecimal(x.BlankValue)} - {FormatDecimal(x.RequiredValue)} = {FormatDecimal(x.BlankValue - x.RequiredValue)}"));

        return new BlankSelectionRow(
            blank.Id,
            alias?.OneCCode ?? string.Empty,
            UiText.Clean(alias?.SourceName ?? blank.CanonicalName),
            UiText.DisplayBlankType(blank.BlankType),
            UiText.Clean(blank.Material),
            FormatSize(blank),
            UiText.DisplayUnit(blank.BaseUnit),
            blank.BaseUnit,
            FormatDecimal(stock),
            FormatDecimal(allowance),
            allowance,
            details);
    }

    private static IEnumerable<DimensionComparison> BuildComparisons(CanonicalBlank blank, BlankSelectionRequest request)
    {
        if (blank.BlankType is BlankType.RoundBar or BlankType.BronzeBar or BlankType.PipeRound)
        {
            var requiredDiameter = request.Diameter ?? Max(request.Width, request.Height, request.Thickness);
            if (requiredDiameter is not null && blank.DiameterMm is not null)
            {
                yield return new DimensionComparison("D", requiredDiameter.Value, blank.DiameterMm.Value);
            }

            if (request.WallThickness is not null && blank.WallThicknessMm is not null)
            {
                yield return new DimensionComparison("S", request.WallThickness.Value, blank.WallThicknessMm.Value);
            }
        }
        else if (blank.BlankType is BlankType.SquareBar or BlankType.HexBar)
        {
            var requiredWidth = request.Width ?? request.Diameter ?? Max(request.Height, request.Thickness);
            if (requiredWidth is not null && blank.WidthMm is not null)
            {
                yield return new DimensionComparison("W", requiredWidth.Value, blank.WidthMm.Value);
            }
        }
        else
        {
            if (request.Width is not null && blank.WidthMm is not null)
            {
                yield return new DimensionComparison("W", request.Width.Value, blank.WidthMm.Value);
            }

            if (request.Height is not null && blank.HeightMm is not null)
            {
                yield return new DimensionComparison("H", request.Height.Value, blank.HeightMm.Value);
            }

            if (request.Thickness is not null && blank.ThicknessMm is not null)
            {
                yield return new DimensionComparison("T", request.Thickness.Value, blank.ThicknessMm.Value);
            }
        }

        if (request.Length is not null && blank.LengthMm is not null)
        {
            yield return new DimensionComparison("L", request.Length.Value, blank.LengthMm.Value);
        }
    }

    private static decimal? Max(params decimal?[] values)
    {
        var present = values.Where(x => x is not null).Select(x => x!.Value).ToArray();
        return present.Length == 0 ? null : present.Max();
    }

    private static decimal? ParseNullable(string? value) =>
        decimal.TryParse((value ?? string.Empty).Trim().Replace('.', ','), NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out var result)
            ? result
            : null;

    private static string ExtractIps(string? value)
    {
        var text = (value ?? string.Empty).Trim();
        var separator = text.IndexOf('|', StringComparison.Ordinal);
        return separator > 0 ? text[..separator].Trim() : text;
    }

    private static string NormalizeMaterial(string value) => value.ToUpperInvariant().Replace(" ", string.Empty, StringComparison.Ordinal);
    private static bool MaterialMatches(CanonicalBlank blank, string materialKey)
    {
        if (string.IsNullOrWhiteSpace(materialKey))
        {
            return true;
        }

        return NormalizeMaterial(blank.Material ?? string.Empty).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
            NormalizeMaterial(blank.CanonicalName).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
            blank.Aliases.Any(a =>
                NormalizeMaterial(a.SourceName).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
                NormalizeMaterial(a.NormalizedSourceName).Contains(materialKey, StringComparison.OrdinalIgnoreCase));
    }
    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    private static string FormatSize(CanonicalBlank blank)
    {
        var parts = new List<string>();
        if (blank.DiameterMm is not null) parts.Add($"D{FormatDecimal(blank.DiameterMm.Value)}");
        if (blank.WidthMm is not null) parts.Add($"W{FormatDecimal(blank.WidthMm.Value)}");
        if (blank.HeightMm is not null) parts.Add($"H{FormatDecimal(blank.HeightMm.Value)}");
        if (blank.ThicknessMm is not null) parts.Add($"T{FormatDecimal(blank.ThicknessMm.Value)}");
        if (blank.WallThicknessMm is not null) parts.Add($"S{FormatDecimal(blank.WallThicknessMm.Value)}");
        if (blank.LengthMm is not null) parts.Add($"L{FormatDecimal(blank.LengthMm.Value)}");
        return string.Join(" ", parts);
    }
}
public sealed partial class CalculationViewModel(
    BlankDemandPlannerDbContext dbContext,
    IBlankDemandCalculationService calculationService,
    IReportExportService reportExportService,
    IFileDialogService fileDialogService) : ObservableObject
{
    public ObservableCollection<CalculationMaterialRow> Rows { get; } = [];
    public ObservableCollection<LibraryBlankOption> BlankSuggestions { get; } = [];
    public IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes { get; } = UiText.ConsumptionUnitTypes;
    private const string OneTimeBlankComment = "Разовая заготовка из расчета";
    [ObservableProperty] private string statusText = string.Empty;
    [ObservableProperty] private CalculationMaterialRow? selectedRow;
    [ObservableProperty] private string blankSearch = string.Empty;
    [ObservableProperty] private LibraryBlankOption? selectedBlank;
    [ObservableProperty] private string selectedOneCCode = string.Empty;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiText.ConsumptionUnitTypes[0];
    [ObservableProperty] private string assignmentStatus = string.Empty;
    private long? _lastRunId;
    private long? _lastBatchId;

    partial void OnBlankSearchChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnSelectedBlankChanged(LibraryBlankOption? value)
    {
        SelectedOneCCode = value?.OneCCode ?? string.Empty;
        if (value is not null)
        {
            SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == value.BaseUnit) ?? SelectedUnit;
        }
    }

    partial void OnSelectedRowChanged(CalculationMaterialRow? value)
    {
        if (value is null)
        {
            return;
        }

        AssignmentStatus = value.IsMissingBlank
            ? $"Нет подобранной заготовки для IPS {value.Ips}. Выберите заготовку из НСИ и укажите норму расхода."
            : $"Выбрана строка IPS {value.Ips}.";
    }

    [RelayCommand]
    public async Task CalculateAsync()
    {
        var batchId = await dbContext.DemandBatches.AsNoTracking().OrderByDescending(x => x.ImportedAt).Select(x => (long?)x.Id).FirstOrDefaultAsync();
        if (batchId is null)
        {
            StatusText = "Нет импортированной или введенной потребности.";
            return;
        }

        var run = await calculationService.CalculateAsync(new CalculationOptions(batchId.Value), CancellationToken.None);
        _lastRunId = run.Id;
        _lastBatchId = batchId.Value;
        await LoadLastRunAsync();
    }

    [RelayCommand]
    public async Task ExportAsync()
    {
        if (_lastRunId is null)
        {
            await LoadLastRunAsync();
        }

        if (_lastRunId is null)
        {
            StatusText = "Нет расчета для экспорта.";
            return;
        }

        var folder = fileDialogService.SelectFolder();
        if (folder is null)
        {
            return;
        }

        try
        {
            var path = await reportExportService.ExportCalculationRunAsync(_lastRunId.Value, folder, CancellationToken.None);
            StatusText = $"Экспортировано: {path}";
            MessageBox.Show(StatusText, "Экспорт расчета материалов", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusText = "Не удалось выгрузить расчет материалов в Excel.";
            MessageBox.Show($"Не удалось выгрузить расчет материалов в Excel.\n{ex.GetBaseException().Message}", "Экспорт расчета материалов", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public async Task LoadLastRunAsync()
    {
        var run = await dbContext.CalculationRuns.AsNoTracking()
            .Include(x => x.Items)
            .ThenInclude(x => x.Sources)
            .OrderByDescending(x => x.StartedAt)
            .FirstOrDefaultAsync();
        Rows.Clear();
        if (run is null)
        {
            StatusText = "Расчетов пока нет.";
            return;
        }

        _lastRunId = run.Id;
        _lastBatchId = run.DemandBatchId;
        await LoadMaterialRowsAsync(run);
        StatusText = $"Расчет материалов #{run.Id}, строк заявки: {Rows.Count}";
    }

    [RelayCommand]
    private async Task AssignBlankToSelectedAsync()
    {
        if (SelectedRow is null)
        {
            AssignmentStatus = "Выберите строку расчета.";
            return;
        }

        if (SelectedBlank is null)
        {
            AssignmentStatus = "Выберите заготовку из НСИ.";
            return;
        }

        if (!TryParseQuantity(ConsumptionQuantityText, out var quantity) || quantity <= 0)
        {
            AssignmentStatus = "Количество материала на деталь должно быть положительным числом, например 0,1 или 0,35.";
            return;
        }

        if (_lastRunId is null)
        {
            AssignmentStatus = "Нет текущего расчета для разового назначения.";
            return;
        }

        if (!TryParseQuantity(SelectedRow.MaterialQuantity, out var demandQuantity) || demandQuantity <= 0)
        {
            AssignmentStatus = "В выбранной строке нет количества для разовой заготовки.";
            return;
        }

        var blank = await dbContext.CanonicalBlanks
            .Include(x => x.Aliases)
            .FirstAsync(x => x.Id == SelectedBlank.Id);
        var alias = blank.Aliases.FirstOrDefault(x => x.IsActive);
        var required = demandQuantity * quantity;
        var calculationItem = new CalculationItem
        {
            CalculationRunId = _lastRunId.Value,
            CanonicalBlankId = blank.Id,
            CanonicalName = UiText.Clean(alias?.SourceName ?? blank.CanonicalName),
            PrimaryOneCCode = alias?.OneCCode,
            OneCCodes = alias?.OneCCode ?? string.Empty,
            Unit = SelectedUnit.Value,
            TotalRequired = required,
            TotalStock = 0,
            PurchaseQuantity = required,
            Status = CalculationStatus.Ok,
            Comment = $"{OneTimeBlankComment}; DemandItemId={SelectedRow.DemandItemId}"
        };
        calculationItem.Sources.Add(new CalculationItemSource
        {
            Ips = SelectedRow.Ips,
            PartName = SelectedRow.Name,
            DemandQuantity = demandQuantity,
            RequiredQuantity = required,
            Unit = SelectedUnit.Value
        });
        dbContext.CalculationItems.Add(calculationItem);

        await dbContext.SaveChangesAsync();
        UndoCenter.Push("Отмена разовой заготовки", async () =>
        {
            var saved = await dbContext.CalculationItems.FirstOrDefaultAsync(x => x.Id == calculationItem.Id);
            if (saved is not null)
            {
                dbContext.CalculationItems.Remove(saved);
                await dbContext.SaveChangesAsync();
            }

            await LoadLastRunAsync();
        });
        AssignmentStatus = $"Разовая заготовка добавлена в текущий расчет: IPS {SelectedRow.Ips} -> {SelectedBlank.DisplayName}.";
        await LoadLastRunAsync();
    }

    [RelayCommand]
    private async Task LoadBlankSuggestionsAsync()
    {
        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);

        var searchText = BlankSearch.Trim();
        var options = await query
            .OrderBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Take(5000)
            .ToListAsync();
        if (!string.IsNullOrWhiteSpace(searchText))
        {
            options = options
                .Where(x =>
                    ContainsSearch(x.CanonicalName, searchText) ||
                    ContainsSearch(x.Material, searchText) ||
                    x.Aliases.Any(a => ContainsSearch(a.OneCCode, searchText) || ContainsSearch(a.SourceName, searchText)))
                .Take(50)
                .ToList();
        }
        else
        {
            options = options.Take(50).ToList();
        }

        var blankOptions = options
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .ToList();

        BlankSuggestions.Clear();
        foreach (var option in blankOptions)
        {
            BlankSuggestions.Add(option);
        }

        if (SelectedBlank is not null && BlankSuggestions.All(x => x.Id != SelectedBlank.Id))
        {
            SelectedBlank = null;
        }
    }

    private async Task LoadMaterialRowsAsync(CalculationRun run)
    {
        var purchaseRemaining = run.Items
            .Where(x => x.CanonicalBlankId is not null && x.PurchaseQuantity > 0 && x.Status == CalculationStatus.Ok)
            .GroupBy(x => (BlankId: x.CanonicalBlankId!.Value, x.Unit))
            .ToDictionary(x => x.Key, x => x.Sum(i => i.PurchaseQuantity));

        var demands = await dbContext.DemandItems.AsNoTracking()
            .Where(x => x.DemandBatchId == run.DemandBatchId)
            .OrderBy(x => x.DemandDate)
            .ThenBy(x => x.Project)
            .ThenBy(x => x.SerialNumber)
            .ThenBy(x => x.Ips)
            .ToListAsync();

        var ipsValues = demands.Select(x => x.Ips).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var parts = await dbContext.Parts.AsNoTracking()
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .Where(x => ipsValues.Contains(x.Ips))
            .ToDictionaryAsync(x => x.Ips, StringComparer.OrdinalIgnoreCase);
        var workInProgressByIps = await LoadWorkInProgressByIpsAsync(ipsValues);
        var oneTimeAssignments = run.Items
            .Where(x => x.Comment?.StartsWith(OneTimeBlankComment, StringComparison.OrdinalIgnoreCase) == true)
            .Select(x => (Item: x, DemandItemId: TryGetOneTimeDemandItemId(x.Comment)))
            .Where(x => x.DemandItemId is not null)
            .ToDictionary(x => x.DemandItemId!.Value, x => x.Item);

        Rows.Clear();
        var number = 1;
        foreach (var demand in demands)
        {
            var inProduction = ApplyWorkInProgress(demand.Ips, demand.Quantity, workInProgressByIps, out var effectiveDemandQuantity);
            if (effectiveDemandQuantity <= 0)
            {
                continue;
            }

            if (oneTimeAssignments.TryGetValue(demand.Id, out var oneTimeItem))
            {
                var source = oneTimeItem.Sources.FirstOrDefault();
                Rows.Add(new CalculationMaterialRow(
                    number++,
                    UiText.Clean(demand.Project),
                    UiText.Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    UiText.Clean(FirstNotEmpty(demand.SourcePartName, source?.PartName, demand.Ips)),
                    FormatDecimal(demand.Quantity),
                    FormatDecimal(inProduction),
                    oneTimeItem.PrimaryOneCCode ?? string.Empty,
                    string.Empty,
                    UiText.Clean(oneTimeItem.CanonicalName),
                    UiText.DisplayUnit(oneTimeItem.Unit),
                    FormatDecimal(oneTimeItem.PurchaseQuantity),
                    FormatBlankDemandDate(demand.DemandDate, 30),
                    false,
                    demand.Id));
                continue;
            }

            parts.TryGetValue(demand.Ips, out var part);
            var maps = part?.BlankMaps
                .Where(x => x.IsActive)
                .GroupBy(x => new { x.CanonicalBlankId, x.ConsumptionUnit })
                .Select(x => x.OrderByDescending(m => m.IsPrimary).ThenByDescending(m => m.UpdatedAt).First())
                .OrderByDescending(x => x.IsPrimary)
                .ToList() ?? [];
            if (maps.Count == 0)
            {
                Rows.Add(new CalculationMaterialRow(
                    number++,
                    UiText.Clean(demand.Project),
                    UiText.Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    UiText.Clean(FirstNotEmpty(demand.SourcePartName, part?.Name, demand.Ips)),
                    FormatDecimal(demand.Quantity),
                    FormatDecimal(inProduction),
                    string.Empty,
                    string.Empty,
                    "Новая номенклатура",
                    UiText.Clean(demand.Unit),
                    FormatDecimal(effectiveDemandQuantity),
                    FormatBlankDemandDate(demand.DemandDate, 30),
                    true,
                    demand.Id));
                continue;
            }

            foreach (var map in maps)
            {
                var blank = map.CanonicalBlank;
                var alias = blank?.Aliases.FirstOrDefault(x => x.IsActive);
                var required = effectiveDemandQuantity * map.ConsumptionQuantity * (1 + map.LossPercent / 100m);
                var materialQuantity = AllocatePurchaseQuantity(purchaseRemaining, map.CanonicalBlankId, map.ConsumptionUnit, required);
                if (materialQuantity <= 0)
                {
                    continue;
                }

                Rows.Add(new CalculationMaterialRow(
                    number++,
                    UiText.Clean(demand.Project),
                    UiText.Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    UiText.Clean(FirstNotEmpty(demand.SourcePartName, part?.Name, demand.Ips)),
                    FormatDecimal(demand.Quantity),
                    FormatDecimal(inProduction),
                    alias?.OneCCode ?? string.Empty,
                    blank is null ? string.Empty : UiText.DisplayBlankType(blank.BlankType),
                    UiText.Clean(alias?.SourceName ?? blank?.CanonicalName),
                    UiText.DisplayUnit(map.ConsumptionUnit),
                    FormatDecimal(materialQuantity),
                    FormatBlankDemandDate(demand.DemandDate, map.BlankLeadTimeDays),
                    false,
                    demand.Id));
            }
        }
    }

    private static decimal AllocatePurchaseQuantity(IDictionary<(long BlankId, MeasurementUnit Unit), decimal> purchaseRemaining, long blankId, MeasurementUnit unit, decimal required)
    {
        if (!purchaseRemaining.TryGetValue((blankId, unit), out var remaining) || remaining <= 0 || required <= 0)
        {
            return 0m;
        }

        var allocated = Math.Min(required, remaining);
        purchaseRemaining[(blankId, unit)] = remaining - allocated;
        return allocated;
    }

    private async Task<Dictionary<string, decimal>> LoadWorkInProgressByIpsAsync(string[] ipsValues)
    {
        if (ipsValues.Length == 0)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();

        if (latestSnapshotId is null)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var ipsKeys = ipsValues.Select(StockCodeNormalizer.NormalizeForComparison).Where(x => x.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stockItems = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.StockSnapshotId == latestSnapshotId.Value && x.Unit == MeasurementUnit.Piece)
            .ToListAsync();

        return stockItems
            .Where(x => StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse) && ipsKeys.Contains(StockCodeNormalizer.NormalizeForComparison(x.OneCCode)))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.Sum(i => i.Quantity), StringComparer.OrdinalIgnoreCase);
    }

    private static decimal ApplyWorkInProgress(string ips, decimal demandQuantity, IDictionary<string, decimal> workInProgressByIps, out decimal effectiveDemandQuantity)
    {
        var key = StockCodeNormalizer.NormalizeForComparison(ips);
        if (!workInProgressByIps.TryGetValue(key, out var inProduction) || inProduction <= 0)
        {
            effectiveDemandQuantity = demandQuantity;
            return 0m;
        }

        var used = Math.Min(demandQuantity, inProduction);
        workInProgressByIps[key] = inProduction - used;
        effectiveDemandQuantity = demandQuantity - used;
        return used;
    }

    private static string FormatBlankDemandDate(DateTime? demandDate, int leadTimeDays) =>
        demandDate?.AddDays(-Math.Max(0, leadTimeDays)).ToLocalTime().ToString("dd.MM.yyyy") ?? string.Empty;

    private async Task RestoreMapsAsync(List<LibraryMapSnapshot> snapshot)
    {
        var ids = snapshot.Select(x => x.Id).ToArray();
        var maps = await dbContext.PartBlankMaps.Where(x => ids.Contains(x.Id)).ToListAsync();
        foreach (var saved in snapshot)
        {
            var map = maps.FirstOrDefault(x => x.Id == saved.Id);
            if (map is null)
            {
                continue;
            }

            map.IsActive = saved.IsActive;
            map.IsPrimary = saved.IsPrimary;
            map.UpdatedAt = saved.UpdatedAt;
        }

        await dbContext.SaveChangesAsync();
        await LoadLastRunAsync();
    }

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
    private static bool ContainsSearch(string? value, string searchText) =>
        !string.IsNullOrWhiteSpace(value) && value.Contains(searchText, StringComparison.OrdinalIgnoreCase);
    private static long? TryGetOneTimeDemandItemId(string? comment)
    {
        const string marker = "DemandItemId=";
        var index = comment?.IndexOf(marker, StringComparison.OrdinalIgnoreCase) ?? -1;
        if (index < 0 || comment is null)
        {
            return null;
        }

        var value = comment[(index + marker.Length)..].Trim();
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : null;
    }
    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    private static bool TryParseQuantity(string? value, out decimal quantity)
    {
        var normalized = value?.Trim().Replace('.', ',') ?? string.Empty;
        return decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out quantity);
    }
}
public sealed partial class HistoryViewModel : ObservableObject
{
    public ObservableCollection<VersionHistoryRow> Rows { get; } = [];

    [ObservableProperty] private string statusText = "История обновлений не загружена";

    [RelayCommand]
    public Task LoadAsync()
    {
        Rows.Clear();
        var path = FindProjectFile("VERSION.md");
        if (path is null)
        {
            Rows.Add(new VersionHistoryRow("Нет данных", string.Empty, "Файл VERSION.md не найден рядом с приложением."));
            StatusText = "История версий не найдена.";
            return Task.CompletedTask;
        }

        foreach (var row in ParseVersionFile(path))
        {
            Rows.Add(row);
        }

        StatusText = $"Показано версий: {Rows.Count}; источник: {path}";
        return Task.CompletedTask;
    }

    private static IEnumerable<VersionHistoryRow> ParseVersionFile(string path)
    {
        string? version = null;
        string updatedAt = string.Empty;
        var bullets = new List<string>();

        foreach (var rawLine in File.ReadLines(path, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (line.StartsWith("## ", StringComparison.Ordinal))
            {
                if (version is not null)
                {
                    yield return new VersionHistoryRow(version, updatedAt, string.Join(" ", bullets.Take(4)));
                }

                (version, updatedAt) = ParseVersionHeading(line[3..].Trim());
                bullets.Clear();
                continue;
            }

            if (version is not null && line.StartsWith("- ", StringComparison.Ordinal))
            {
                bullets.Add(line[2..].Trim());
            }
        }

        if (version is not null)
        {
            yield return new VersionHistoryRow(version, updatedAt, string.Join(" ", bullets.Take(4)));
        }
    }

    private static (string Version, string UpdatedAt) ParseVersionHeading(string heading)
    {
        var separators = new[] { " — ", " - ", " | " };
        foreach (var separator in separators)
        {
            var index = heading.IndexOf(separator, StringComparison.Ordinal);
            if (index > 0)
            {
                return (heading[..index].Trim(), heading[(index + separator.Length)..].Trim());
            }
        }

        return (heading.Trim(), string.Empty);
    }

    private static string? FindProjectFile(string fileName)
    {
        foreach (var basePath in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory })
        {
            var directory = new DirectoryInfo(basePath);
            for (var i = 0; directory is not null && i < 10; i++, directory = directory.Parent)
            {
                var path = Path.Combine(directory.FullName, fileName);
                if (File.Exists(path))
                {
                    return path;
                }
            }
        }

        return null;
    }
}

public sealed partial class SettingsViewModel(BlankDemandPlannerDbContext dbContext) : ObservableObject
{
    public IReadOnlyList<string> FontFamilies { get; } = ["Segoe UI", "Arial", "Calibri", "Tahoma", "Times New Roman"];
    public IReadOnlyList<double> FontSizes { get; } = [12, 13, 14, 15, 16, 18, 20];
    public IReadOnlyList<DisplayOption<HorizontalAlignment>> ToolbarPlacements { get; } =
    [
        new(HorizontalAlignment.Left, "Слева"),
        new(HorizontalAlignment.Center, "По центру"),
        new(HorizontalAlignment.Right, "Справа")
    ];

    [ObservableProperty] private string selectedFontFamily = "Segoe UI";
    [ObservableProperty] private double selectedFontSize = 13;
    [ObservableProperty] private DisplayOption<HorizontalAlignment> selectedToolbarPlacement = new(HorizontalAlignment.Left, "Слева");
    [ObservableProperty] private HorizontalAlignment toolbarHorizontalAlignment = HorizontalAlignment.Left;
    [ObservableProperty] private string iconPath = string.Empty;
    [ObservableProperty] private string photoPath = string.Empty;
    [ObservableProperty] private string statusText = "Настройки готовы";

    [RelayCommand]
    public async Task LoadAsync()
    {
        SelectedFontFamily = await ReadSettingAsync("UI.FontFamily", "Segoe UI");
        if (double.TryParse(await ReadSettingAsync("UI.FontSize", "13"), NumberStyles.Number, CultureInfo.InvariantCulture, out var fontSize))
        {
            SelectedFontSize = Math.Clamp(fontSize, 10, 24);
        }

        var placement = await ReadSettingAsync("UI.ToolbarPlacement", "Left");
        SelectedToolbarPlacement = ToolbarPlacements.FirstOrDefault(x => string.Equals(x.Value.ToString(), placement, StringComparison.OrdinalIgnoreCase))
            ?? ToolbarPlacements[0];
        ToolbarHorizontalAlignment = SelectedToolbarPlacement.Value;
        IconPath = await ReadSettingAsync("UI.IconPath", string.Empty);
        PhotoPath = await ReadSettingAsync("UI.PhotoPath", string.Empty);
        ApplyToWindow();
    }

    [RelayCommand]
    private async Task ApplyAsync()
    {
        ToolbarHorizontalAlignment = SelectedToolbarPlacement.Value;
        ApplyToWindow();
        await WriteSettingAsync("UI.FontFamily", SelectedFontFamily);
        await WriteSettingAsync("UI.FontSize", SelectedFontSize.ToString(CultureInfo.InvariantCulture));
        await WriteSettingAsync("UI.ToolbarPlacement", SelectedToolbarPlacement.Value.ToString());
        await WriteSettingAsync("UI.IconPath", IconPath);
        await WriteSettingAsync("UI.PhotoPath", PhotoPath);
        await dbContext.SaveChangesAsync();
        StatusText = "Настройки применены и сохранены.";
    }

    [RelayCommand]
    private void ChooseIcon()
    {
        var path = SelectImageFile();
        if (path is not null)
        {
            IconPath = path;
        }
    }

    [RelayCommand]
    private void ChoosePhoto()
    {
        var path = SelectImageFile();
        if (path is not null)
        {
            PhotoPath = path;
        }
    }

    [RelayCommand]
    private async Task ResetAsync()
    {
        SelectedFontFamily = "Segoe UI";
        SelectedFontSize = 13;
        SelectedToolbarPlacement = ToolbarPlacements[0];
        IconPath = string.Empty;
        PhotoPath = string.Empty;
        await ApplyAsync();
    }

    partial void OnSelectedToolbarPlacementChanged(DisplayOption<HorizontalAlignment> value) => ToolbarHorizontalAlignment = value.Value;

    private void ApplyToWindow()
    {
        var window = Application.Current?.MainWindow;
        if (window is null)
        {
            return;
        }

        window.FontFamily = new FontFamily(SelectedFontFamily);
        window.FontSize = SelectedFontSize;
        if (!string.IsNullOrWhiteSpace(IconPath) && File.Exists(IconPath))
        {
            window.Icon = new BitmapImage(new Uri(IconPath));
        }
    }

    private async Task<string> ReadSettingAsync(string key, string defaultValue)
    {
        return await dbContext.Settings.AsNoTracking()
            .Where(x => x.Key == key)
            .Select(x => x.Value)
            .FirstOrDefaultAsync() ?? defaultValue;
    }

    private async Task WriteSettingAsync(string key, string value)
    {
        var setting = await dbContext.Settings.FirstOrDefaultAsync(x => x.Key == key);
        if (setting is null)
        {
            dbContext.Settings.Add(new AppSetting { Key = key, Value = value });
            return;
        }

        setting.Value = value;
    }

    private static string? SelectImageFile()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Изображения|*.png;*.jpg;*.jpeg;*.bmp;*.ico|Все файлы|*.*",
            CheckFileExists = true
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}

internal static class UiText
{
    public static IReadOnlyList<DisplayOption<BlankType>> BlankTypes { get; } =
    [
        new(BlankType.RoundBar, "Круг"),
        new(BlankType.SquareBar, "Квадрат"),
        new(BlankType.HexBar, "Шестигранник"),
        new(BlankType.Sheet, "Лист"),
        new(BlankType.Plate, "Плита"),
        new(BlankType.PipeRound, "Труба профильная круглая"),
        new(BlankType.PipeRectangular, "Труба профильная прямоугольная"),
        new(BlankType.Angle, "Уголок"),
        new(BlankType.Channel, "Швеллер"),
        new(BlankType.IBeam, "Двутавр"),
        new(BlankType.BronzeBar, "Пруток бронзовый"),
        new(BlankType.BronzeSheet, "Лист бронзовый"),
        new(BlankType.WeldingElement, "Сварочный элемент"),
        new(BlankType.Purchased, "Покупная"),
        new(BlankType.Casting, "Литье"),
        new(BlankType.Forging, "Поковка"),
        new(BlankType.Unknown, "Не распознано")
    ];

    public static IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFilters { get; } =
    [
        new((BlankType?)null, "Все виды"),
        ..BlankTypes.Where(x => x.Value is not BlankType.Unknown and not BlankType.CustomBlank).Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName))
    ];

    public static IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFiltersWithUnknown { get; } =
    [
        new((BlankType?)null, "Все виды"),
        ..BlankTypes.Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName))
    ];

    public static IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes { get; } =
    [
        new(MeasurementUnit.Piece, "шт"),
        new(MeasurementUnit.Meter, "пог. м"),
        new(MeasurementUnit.Kilogram, "кг")
    ];

    public static IReadOnlyList<DisplayOption<MeasurementUnit>> ConsumptionUnitTypes { get; } =
    [
        new(MeasurementUnit.Piece, "шт"),
        new(MeasurementUnit.Meter, "пог. м")
    ];

    public static IReadOnlyList<DisplayOption<MeasurementUnit?>> UnitFilters { get; } =
    [
        new((MeasurementUnit?)null, "Все ед."),
        new(MeasurementUnit.Piece, "шт"),
        new(MeasurementUnit.Meter, "пог. м"),
        new(MeasurementUnit.Kilogram, "кг")
    ];

    public static IReadOnlyList<DisplayOption<MeasurementUnit?>> ConsumptionUnitFilters { get; } =
    [
        new((MeasurementUnit?)null, "Все ед."),
        new(MeasurementUnit.Piece, "шт"),
        new(MeasurementUnit.Meter, "пог. м")
    ];

    public static string DisplayUnit(MeasurementUnit unit) => UnitTypes.FirstOrDefault(x => x.Value == unit)?.DisplayName ?? unit.ToString();
    public static string DisplayBlankType(BlankType type) =>
        type == BlankType.CustomBlank ? "Прочее" : BlankTypes.FirstOrDefault(x => x.Value == type)?.DisplayName ?? type.ToString();

    public static string Clean(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var text = value.Trim();
        for (var attempt = 0; attempt < 3 && LooksLikeMojibake(text); attempt++)
        {
            try
            {
                var bytes = Encoding.GetEncoding(1251).GetBytes(text);
                var repaired = Encoding.UTF8.GetString(bytes);
                if (CyrillicScore(repaired) < CyrillicScore(text))
                {
                    break;
                }

                text = repaired;
            }
            catch
            {
                break;
            }
        }

        text = Regex.Replace(text, @"�(?=\d)", "Р", RegexOptions.CultureInvariant);
        return text;
    }

    private static bool LooksLikeMojibake(string text) =>
        text.Contains("Р ", StringComparison.Ordinal) ||
        text.Contains("Р’", StringComparison.Ordinal) ||
        text.Contains("РЎ", StringComparison.Ordinal) ||
        text.Contains("РЋ", StringComparison.Ordinal) ||
        text.Contains("РІ", StringComparison.Ordinal) ||
        text.Contains("СЃ", StringComparison.Ordinal) ||
        text.Contains("С‚", StringComparison.Ordinal) ||
        text.Contains("СЂ", StringComparison.Ordinal) ||
        text.Contains('Ђ');

    private static int CyrillicScore(string text) => text.Count(ch => ch is >= 'А' and <= 'я' or 'ё' or 'Ё') - text.Count(ch => ch is 'Ђ' or 'С' or 'Р');
}

public sealed record VersionHistoryRow(string Version, string UpdatedAt, string Description);

public static class UndoCenter
{
    private const int MaxActions = 10;
    private static readonly Stack<UndoAction> Actions = new();

    public static event EventHandler? Changed;
    public static string StatusText => Actions.Count == 0 ? "Отменить" : $"Отменить: {Actions.Peek().Title}";

    public static void Push(string title, Func<Task> undo)
    {
        Actions.Push(new UndoAction(title, undo));
        while (Actions.Count > MaxActions)
        {
            var kept = Actions.Reverse().Take(MaxActions).Reverse().ToArray();
            Actions.Clear();
            foreach (var action in kept)
            {
                Actions.Push(action);
            }
        }

        Changed?.Invoke(null, EventArgs.Empty);
    }

    public static async Task<string> UndoAsync()
    {
        if (Actions.Count == 0)
        {
            return "Нет действий для отмены.";
        }

        var action = Actions.Pop();
        await action.Undo();
        Changed?.Invoke(null, EventArgs.Empty);
        return $"Отменено: {action.Title}";
    }

    private sealed record UndoAction(string Title, Func<Task> Undo);
}

public sealed record DemandSnapshot(long Id, string? Project, string? SerialNumber, string? ProductionSystem, string Ips, string? SourcePartName, string? Unit, decimal Quantity, DateTime? DemandDate);

public sealed record LibraryDeleteSnapshot(List<LibraryMapSnapshot> Maps, List<LibraryPartSnapshot> Parts);
public sealed record LibraryMapSnapshot(long Id, bool IsActive, bool IsPrimary, DateTime UpdatedAt);
public sealed record LibraryPartSnapshot(long Id, string? Source, DateTime UpdatedAt);

public sealed record LibraryRow(long PartId, long? PartBlankMapId, long? CanonicalBlankId, string Ips, string? Designation, string PartName, string? BlankType, string? BlankName, string? Material, string? OneCCode, decimal? ConsumptionQuantity, MeasurementUnit? ConsumptionUnit, string Quantity, string UnitName, int BlankLeadTimeDays, string BlankLeadTimeDaysText, string? Source, string UpdatedAt);

public sealed record LibraryBlankOption(long Id, BlankType BlankType, string CanonicalName, string? SourceName, string? Material, MeasurementUnit BaseUnit, string? OneCCode)
{
    public string DisplayName => $"{UiText.Clean(string.IsNullOrWhiteSpace(SourceName) ? CanonicalName : SourceName)} | {OneCCode ?? "без УТ"}";
}

public sealed partial class DemandRow(long id, string project, string machineNumber, string ips, string name, string unitName, string quantity, string inProductionQuantity, string workInProgressBefore, string workInProgressAfter, string demandDate, bool isManual) : ObservableObject
{
    public long Id { get; } = id;
    public bool IsManual { get; } = isManual;
    [ObservableProperty] private string project = project;
    [ObservableProperty] private string machineNumber = machineNumber;
    [ObservableProperty] private string ips = ips;
    [ObservableProperty] private string name = name;
    [ObservableProperty] private string unitName = unitName;
    [ObservableProperty] private string quantity = quantity;
    [ObservableProperty] private string inProductionQuantity = inProductionQuantity;
    [ObservableProperty] private string workInProgressBefore = workInProgressBefore;
    [ObservableProperty] private string workInProgressAfter = workInProgressAfter;
    [ObservableProperty] private string demandDate = demandDate;
}

public sealed record DemandDetailRow(string DemandDate, string Quantity, string InProductionQuantity, string WorkInProgressAfter, string Project, string MachineNumber);

public sealed record CalculationMaterialRow(int Number, string ProductionSystem, string MachineNumber, string Ips, string Name, string PartQuantity, string InProductionQuantity, string OneCCode, string BlankType, string Nomenclature, string UnitName, string MaterialQuantity, string DemandDate, bool IsMissingBlank, long DemandItemId);

public sealed record PartWithoutBlankOption(long PartId, string Ips, string? Designation, string PartName)
{
    public string DisplayName => string.IsNullOrWhiteSpace(Designation)
        ? $"{Ips} | {UiText.Clean(PartName)}"
        : $"{Ips} | {UiText.Clean(Designation)} {UiText.Clean(PartName)}";
}

public sealed record BlankSelectionRow(long BlankId, string OneCCode, string SourceName, string BlankType, string Material, string Size, string UnitName, MeasurementUnit BaseUnit, string StockQuantity, string Allowance, decimal AllowanceScore, string Details);

public sealed record BlankSelectionRequest(decimal? Diameter, decimal? Width, decimal? Height, decimal? Thickness, decimal? WallThickness, decimal? Length)
{
    public bool HasAnySize => Diameter is not null || Width is not null || Height is not null || Thickness is not null || WallThickness is not null || Length is not null;
}

public sealed record DimensionComparison(string Label, decimal RequiredValue, decimal BlankValue);

public sealed record NsiBlankRow(long AliasId, long CanonicalBlankId, string OneCCode, string SourceName, string UnitName, string BlankType, string Size, string Material, string MaterialGost, string ProfileGost, string CmoStockQuantity, string WarehouseStockQuantity, string StockUnitName, string BlankName, string Source, string UpdatedAt)
{
    public string StockQuantity => CmoStockQuantity;
}

public sealed record NsiUsageRow(string Ips, string Designation, string PartName, string Quantity, string UnitName, string Source);

public sealed record StockRow(string OneCCode, string SourceName, string Quantity, string UnitName, string Warehouse);

public sealed record MskLibraryRow(
    string Ips,
    string Designation,
    string Name,
    string HasMsk,
    string HasLibraryPart,
    string FilePath,
    string FileName,
    string ImportedAt,
    string BlankType,
    string BlankName,
    string Material,
    string OneCCode,
    string ConsumptionQuantity,
    string UnitName,
    string BlankLeadTimeDays);

public sealed record MskCsvDetail(string Source, string Ips, string BlankType, string BlankName, string Material, string OneCCode, string ConsumptionQuantity, string UnitName);

public sealed record DisplayOption<T>(T Value, string DisplayName);

public sealed class SimplePageViewModel(string title, string description)
{
    public string Title { get; } = title;
    public string Description { get; } = description;
}





