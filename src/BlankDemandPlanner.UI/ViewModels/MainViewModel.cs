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
using System.Windows.Data;
using System.Windows.Forms.Integration;
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
using PdfiumViewer;
using Application = System.Windows.Application;
using Binding = System.Windows.Data.Binding;
using Brushes = System.Windows.Media.Brushes;
using Button = System.Windows.Controls.Button;
using ComboBox = System.Windows.Controls.ComboBox;
using FontFamily = System.Windows.Media.FontFamily;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using TextBox = System.Windows.Controls.TextBox;

namespace BlankDemandPlanner.UI.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly BlankDemandPlannerDbContext _dbContext;
    private readonly IExcelImportService _excelImportService;
    private readonly IOneCStockSyncService _oneCStockSyncService;
    private readonly IOneCNomenclatureService _oneCNomenclatureService;
    private readonly IOneCProductionLaunchService _oneCProductionLaunchService;
    private readonly IBlankDemandCalculationService _calculationService;
    private readonly IFileDialogService _fileDialogService;
    private readonly IAppAuthService _authService;
    private readonly ILogger<MainViewModel> _logger;

    public ObservableCollection<NavigationItem> NavigationItems { get; } = [];

    [ObservableProperty] private NavigationItem? selectedNavigationItem;
    [ObservableProperty] private object? currentPage;
    [ObservableProperty] private bool canEditCurrentPage = true;
    [ObservableProperty] private string? globalSearch;
    [ObservableProperty] private string undoStatusText = "Отменить";
    [ObservableProperty] private string stockSyncStatusText = "Остатки 1С: не обновлялись";
    [ObservableProperty] private bool isStockSyncRunning;

    public string WindowTitle { get; } = "Планирование ЦМО";
    public string ApplicationTitle { get; } = AppVersionInfo.Current.ApplicationTitle;
    public string CurrentUserText => _authService.CurrentUser is null
        ? "Пользователь не определен"
        : $"{_authService.CurrentUser.DisplayName} ({(_authService.CurrentUser.IsAdmin ? "админ" : "пользователь")})";

    public DashboardViewModel Dashboard { get; }
    public DemandViewModel Demand { get; }
    public LibraryViewModel Library { get; }
    public NormalizationViewModel Normalization { get; }
    public StockViewModel Stock { get; }
    public CalculationViewModel Calculation { get; }
    public BlankSelectionViewModel BlankSelection { get; }
    public HistoryViewModel History { get; }
    public SettingsViewModel Settings { get; }
    public DeveloperModeViewModel DeveloperMode { get; }
    public MskViewModel Msk { get; }
    public ProductionPlanViewModel Plan { get; }
    public WorkshopReportViewModel WorkshopReport { get; }

    public MainViewModel(
        BlankDemandPlannerDbContext dbContext,
        IExcelImportService excelImportService,
        IOneCStockSyncService oneCStockSyncService,
        IOneCNomenclatureService oneCNomenclatureService,
        IOneCProductionLaunchService oneCProductionLaunchService,
        IOneCGoodsTransferService oneCGoodsTransferService,
        IBlankDemandCalculationService calculationService,
        IBlankNormalizationService normalizationService,
        IReportExportService reportExportService,
        IFileDialogService fileDialogService,
        IProductionPlanningService productionPlanningService,
        IPzmcNeedService pzmcNeedService,
        IPzmcNeedExportService pzmcNeedExportService,
        IPzmcPersonnelAvailabilityService pzmcPersonnelAvailabilityService,
        IPzmcProductionApiClient pzmcProductionApiClient,
        IAppAuthService authService,
        ILogger<MainViewModel> logger)
    {
        _dbContext = dbContext;
        _excelImportService = excelImportService;
        _oneCStockSyncService = oneCStockSyncService;
        _oneCNomenclatureService = oneCNomenclatureService;
        _oneCProductionLaunchService = oneCProductionLaunchService;
        _calculationService = calculationService;
        _fileDialogService = fileDialogService;
        _authService = authService;
        _logger = logger;

        var drawingService = new IpsBridgeDrawingService();
        Dashboard = new DashboardViewModel(dbContext);
        Demand = new DemandViewModel(dbContext, excelImportService, drawingService, logger);
        Library = new LibraryViewModel(dbContext, normalizationService, excelImportService, fileDialogService, reportExportService, drawingService, logger);
        Normalization = new NormalizationViewModel(dbContext, excelImportService, fileDialogService, oneCNomenclatureService);
        Stock = new StockViewModel(dbContext);
        Calculation = new CalculationViewModel(dbContext, calculationService, reportExportService, fileDialogService, drawingService, logger);
        BlankSelection = new BlankSelectionViewModel(dbContext, drawingService, logger);
        History = new HistoryViewModel();
        Settings = new SettingsViewModel(dbContext, authService);
        DeveloperMode = new DeveloperModeViewModel(dbContext, authService, RefreshReferenceListsAsync);
        Msk = new MskViewModel(dbContext, logger, drawingService, oneCProductionLaunchService, oneCGoodsTransferService);
        Plan = new ProductionPlanViewModel(
            dbContext,
            productionPlanningService,
            pzmcNeedService,
            pzmcNeedExportService,
            pzmcPersonnelAvailabilityService,
            pzmcProductionApiClient,
            fileDialogService);
        WorkshopReport = new WorkshopReportViewModel(dbContext, pzmcPersonnelAvailabilityService);
        UndoCenter.Changed += (_, _) => UndoStatusText = UndoCenter.StatusText;

        NavigationItems.Add(new NavigationItem("Главная", Dashboard));
        NavigationItems.Add(new NavigationItem("Потребность", Demand));
        NavigationItems.Add(new NavigationItem("Расчет материалов", Calculation));
        NavigationItems.Add(new NavigationItem("Библиотека", Library));
        NavigationItems.Add(new NavigationItem("НСИ", Normalization));
        NavigationItems.Add(new NavigationItem("Подбор заготовок", BlankSelection));
        NavigationItems.Add(new NavigationItem("Планирование", Msk));
        NavigationItems.Add(new NavigationItem("Отчет", WorkshopReport));
        NavigationItems.Add(new NavigationItem("План", Plan));
        NavigationItems.Add(new NavigationItem("История", History));
        NavigationItems.Add(new NavigationItem("Настройки", Settings));
        NavigationItems.Add(new NavigationItem("Режим разработчика", DeveloperMode));
        ApplyNavigationPermissions();
        SelectedNavigationItem = NavigationItems[0];
        _ = Dashboard.LoadAsync();
        _ = Settings.LoadAsync();
        _ = SyncStockFromOneCCoreAsync(showMessage: false);
    }

    private void ApplyNavigationPermissions()
    {
        for (var i = NavigationItems.Count - 1; i >= 0; i--)
        {
            var item = NavigationItems[i];
            var pageKey = GetPageKey(item.Page);
            var canRead = pageKey == "Settings" || _authService.CanRead(pageKey);
            if (!canRead)
            {
                NavigationItems.RemoveAt(i);
                continue;
            }

            NavigationItems[i] = item with { PageKey = pageKey, CanEdit = pageKey == "Settings" || _authService.CanEdit(pageKey) };
        }
    }

    private string GetPageKey(object page)
    {
        if (ReferenceEquals(page, Dashboard)) return "Dashboard";
        if (ReferenceEquals(page, Demand)) return "Demand";
        if (ReferenceEquals(page, Calculation)) return "Calculation";
        if (ReferenceEquals(page, Library)) return "Library";
        if (ReferenceEquals(page, Normalization)) return "Normalization";
        if (ReferenceEquals(page, BlankSelection)) return "BlankSelection";
        if (ReferenceEquals(page, Msk)) return "Planning";
        if (ReferenceEquals(page, WorkshopReport)) return "WorkshopReport";
        if (ReferenceEquals(page, Plan)) return "Plan";
        if (ReferenceEquals(page, History)) return "History";
        if (ReferenceEquals(page, Settings)) return "Settings";
        if (ReferenceEquals(page, DeveloperMode)) return "DeveloperMode";
        return string.Empty;
    }

    partial void OnSelectedNavigationItemChanged(NavigationItem? value)
    {
        CurrentPage = value?.Page;
        CanEditCurrentPage = value?.CanEdit != false;
        ApplyGlobalSearchToPage(value?.Page);
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
        else if (value?.Page == Plan)
        {
            _ = Plan.LoadAsync();
        }
        else if (value?.Page == WorkshopReport)
        {
            _ = WorkshopReport.LoadAsync();
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
        else if (value?.Page == DeveloperMode)
        {
            _ = DeveloperMode.LoadAsync();
        }
    }

    private async Task RefreshReferenceListsAsync()
    {
        await UiReferenceData.LoadAsync(_dbContext);
        Library.RefreshReferenceLists();
        Normalization.RefreshReferenceLists();
        BlankSelection.RefreshReferenceLists();
        Calculation.RefreshReferenceLists();
    }

    partial void OnGlobalSearchChanged(string? value)
    {
        ApplyGlobalSearchToPage(CurrentPage);
    }

    private void ApplyGlobalSearchToPage(object? page)
    {
        var search = GlobalSearch ?? string.Empty;
        if (page == Demand)
        {
            Demand.Search = search;
        }
        else if (page == Library)
        {
            Library.Search = search;
        }
        else if (page == Normalization)
        {
            Normalization.Search = search;
        }
        else if (page == Stock)
        {
            Stock.Search = search;
        }
        else if (page == Calculation)
        {
            Calculation.Search = search;
        }
        else if (page == Msk)
        {
            Msk.Search = search;
        }
    }

    [RelayCommand]
    private async Task RefreshAsync()
    {
        await RefreshReferenceListsAsync();
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

    public async Task StartupLoadAsync() => await RefreshAsync();

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
            var beforeIps = (await _dbContext.Parts.AsNoTracking()
                    .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
                    .Select(x => x.Ips)
                    .ToListAsync())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var report = await _excelImportService.ImportDemandAsync(file, null, CancellationToken.None);
            _dbContext.ChangeTracker.Clear();

            SelectedNavigationItem = NavigationItems.First(x => x.Page == Demand);
            CurrentPage = Demand;
            await Demand.LoadAsync();
            await Dashboard.LoadAsync();
            await Library.LoadAsync();
            await BlankSelection.LoadAsync();
            _dbContext.ChangeTracker.Clear();

            var refreshedIps = await _dbContext.Parts.AsNoTracking()
                .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
                .Select(x => x.Ips)
                .ToListAsync();
            var addedPartCount = refreshedIps.Count(x => !beforeIps.Contains(x));
            if (addedPartCount > 0)
            {
                Demand.StatusText = $"{Demand.StatusText}; добавлено новых деталей: {addedPartCount}";
            }

            MessageBox.Show($"Импорт потребности из ПП завершен.\nПрочитано строк: {report.ReadRows}\nДобавлено: {report.AddedRows}\nОбновлено: {report.UpdatedRows}\nПропущено: {report.SkippedRows}\nОшибок: {report.ErrorRows}\nДобавлено новых деталей в библиотеку: {addedPartCount}", "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
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
            var refreshedNsiNames = await Normalization.RefreshNamesFromOneCAsync(CancellationToken.None);
            var refreshedPrices = await Normalization.RefreshPricesFromOneCAsync(CancellationToken.None);
            await Dashboard.LoadAsync();
            await Demand.LoadAsync();
            await Normalization.LoadAsync();
            await Stock.LoadAsync();
            await Calculation.LoadLastRunAsync();
            StockSyncStatusText = refreshedNsiNames > 0
                ? $"Остатки 1С обновлены: {report.SyncedAt.ToLocalTime():dd.MM.yyyy HH:mm}; склад: {report.WarehouseRows}; НЗП/ЦМО: {report.WipRows}; НСИ: {refreshedNsiNames}; цены: {refreshedPrices}"
                : $"Остатки 1С обновлены: {report.SyncedAt.ToLocalTime():dd.MM.yyyy HH:mm}; склад: {report.WarehouseRows}; НЗП/ЦМО: {report.WipRows}; цены: {refreshedPrices}";
            if (showMessage)
            {
                var nsiLine = refreshedNsiNames > 0
                    ? $"\nНаименований НСИ обновлено из 1С: {refreshedNsiNames}"
                    : "\nНаименования НСИ сверены с 1С, изменений нет";
                MessageBox.Show($"Остатки 1С обновлены.\nСклад: {report.WarehouseRows} строк\nНЗП/ЦМО: {report.WipRows} строк{nsiLine}\nЦен обновлено: {refreshedPrices}\nДата: {report.SyncedAt.ToLocalTime():dd.MM.yyyy HH:mm}", "Обновление данных 1С", MessageBoxButton.OK, MessageBoxImage.Information);
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

public sealed record NavigationItem(string Title, object Page, string PageKey = "", bool CanEdit = false);

public sealed record AppVersionInfo(string Version, string UpdatedAt)
{
    public static AppVersionInfo Current { get; } = Load();
    public string ShortVersion { get; } = ToShortVersion(Version);
    public string ApplicationTitle { get; } = $"Планирование ЦМО {ToShortVersion(Version)}";

    public string SoftwareUpdateText =>
        string.IsNullOrWhiteSpace(UpdatedAt)
            ? $"Версия {ShortVersion}"
            : $"{UpdatedAt}, версия {ShortVersion}";

    public static string ToShortVersion(string version)
    {
        var match = Regex.Match(version.Trim(), @"^v\d{4}\.\d{2}\.\d{2}\.\d+");
        return match.Success ? match.Value : version.Trim();
    }

    private static AppVersionInfo Load()
    {
        var path = FindProjectFile("VERSION.md");
        if (path is not null)
        {
            foreach (var rawLine in File.ReadLines(path, Encoding.UTF8))
            {
                var line = rawLine.Trim();
                if (line.StartsWith("## ", StringComparison.Ordinal))
                {
                    var (version, updatedAt) = ParseVersionHeading(line[3..].Trim());
                    if (!string.IsNullOrWhiteSpace(version))
                    {
                        return new AppVersionInfo(version, updatedAt);
                    }
                }
            }
        }

        return new AppVersionInfo("версия не определена", string.Empty);
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

public sealed partial class DashboardViewModel(BlankDemandPlannerDbContext dbContext) : ObservableObject
{
    public string LastSoftwareUpdate { get; } = AppVersionInfo.Current.SoftwareUpdateText;
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
        var duplicateAliases = await activeAliases
            .Include(x => x.CanonicalBlank)
            .ToListAsync();
        var duplicateCounts = duplicateAliases
            .Select(x => NsiDuplicateKey.Build(x.CanonicalBlank, x))
            .Where(x => x.Length > 0)
            .GroupBy(x => x, StringComparer.Ordinal)
            .Where(g => g.Count() > 1)
            .Select(g => g.Count())
            .ToList();
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

public sealed partial class DemandViewModel(
    BlankDemandPlannerDbContext dbContext,
    IExcelImportService? excelImportService = null,
    IIpsDrawingService? ipsDrawingService = null,
    ILogger? logger = null) : ObservableObject
{
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();

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
            var searchedItems = await query
                .OrderBy(x => x.DemandDate)
                .ThenBy(x => x.Project)
                .ThenBy(x => x.SerialNumber)
                .ThenBy(x => x.Id)
                .Take(20000)
                .ToListAsync();
            searchedItems = searchedItems
                .Where(x => UiSearchText.ContainsAnyField(searchValue,
                    x.Project,
                    x.SerialNumber,
                    x.ProductionSystem,
                    x.SourcePartName,
                    x.Ips))
                .Take(5000)
                .ToList();

            await LoadRowsAsync(searchedItems);
            return;
        }

        var items = await query
            .OrderBy(x => x.DemandDate)
            .ThenBy(x => x.Project)
            .ThenBy(x => x.SerialNumber)
            .ThenBy(x => x.Id)
            .Take(5000)
            .ToListAsync();

        await LoadRowsAsync(items);
    }

    private async Task LoadRowsAsync(IReadOnlyCollection<DemandItem> items)
    {
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
    private async Task OpenDrawingAsync(object? parameter)
    {
        var row = parameter as DemandRow ?? SelectedRow;
        if (row is null)
        {
            StatusText = "Выберите строку потребности для открытия чертежа.";
            return;
        }

        SelectedRow = row;
        StatusText = await DrawingPdfOpener.OpenExternalAsync(
            new DrawingLookupRequest(row.Ips, null, row.Name, null),
            drawingService,
            "Потребность",
            logger,
            CancellationToken.None);
    }

    [RelayCommand]
    private async Task OpenObjectCardAsync(object? parameter)
    {
        var row = parameter as DemandRow ?? SelectedRow;
        if (row is null)
        {
            StatusText = "Выберите строку для открытия карточки объекта.";
            return;
        }

        SelectedRow = row;
        PdfViewer? viewer = null;
        PdfDocument? document = null;
        WindowsFormsHost? host = null;
        var status = new TextBlock
        {
            Text = "Поиск PDF-чертежа...",
            Foreground = Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var details = new TextBlock
        {
            Text = BuildObjectCardText(row),
            TextWrapping = TextWrapping.Wrap,
            FontSize = 14
        };

        var left = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = new StackPanel
            {
                Margin = new Thickness(16),
                Children =
                {
                    new TextBlock { Text = "Карточка объекта", FontSize = 22, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 12) },
                    details,
                    new TextBlock { Text = "2 маршрут обработки и трудоемкость", FontSize = 17, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 18, 0, 6) },
                    new TextBlock { Text = "Данные маршрута пока не внесены.", TextWrapping = TextWrapping.Wrap, Foreground = Brushes.DimGray },
                    new TextBlock { Text = "3 Оснащение и расход", FontSize = 17, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 18, 0, 6) },
                    new TextBlock { Text = "Данные по оснащению и расходу пока не внесены.", TextWrapping = TextWrapping.Wrap, Foreground = Brushes.DimGray },
                    status
                }
            }
        };

        try
        {
            viewer = new PdfViewer
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                ZoomMode = PdfViewerZoomMode.FitWidth
            };
            host = new WindowsFormsHost { Child = viewer };
        }
        catch (Exception ex)
        {
            status.Text = $"PDF-viewer не запущен: {ex.GetBaseException().Message}";
        }

        var layout = new Grid();
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.42, GridUnitType.Star), MinWidth = 360 });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.58, GridUnitType.Star), MinWidth = 420 });
        layout.Children.Add(left);
        var right = new Border
        {
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1, 0, 0, 0),
            Child = host is not null ? host : new TextBlock { Text = status.Text, Margin = new Thickness(16), TextWrapping = TextWrapping.Wrap }
        };
        Grid.SetColumn(right, 1);
        layout.Children.Add(right);

        var window = new Window
        {
            Title = $"Карточка объекта IPS {row.Ips}",
            Owner = Application.Current?.MainWindow,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 1180,
            Height = 760,
            MinWidth = 920,
            MinHeight = 620,
            Content = layout
        };
        window.Closed += (_, _) =>
        {
            if (viewer is not null)
            {
                viewer.Document = null;
            }

            if (host is not null)
            {
                host.Child = null;
            }

            viewer?.Dispose();
            document?.Dispose();
        };
        window.Show();

        if (viewer is null)
        {
            return;
        }

        var drawing = await DrawingPdfOpener.ResolveDrawingPdfAsync(
            new DrawingLookupRequest(row.Ips, null, row.Name, null),
            drawingService,
            logger,
            CancellationToken.None);
        if (drawing is null)
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = $"PDF-чертеж IPS {row.Ips} не найден.";
            return;
        }

        try
        {
            document = PdfDocument.Load(drawing.FullName);
            viewer.Document = document;
            viewer.ZoomMode = PdfViewerZoomMode.FitWidth;
            status.Foreground = Brushes.SeaGreen;
            status.Text = $"PDF-чертеж открыт: {drawing.Name}";
        }
        catch (Exception ex)
        {
            document?.Dispose();
            document = null;
            status.Foreground = Brushes.Firebrick;
            status.Text = $"PDF найден, но viewer не смог открыть файл: {ex.GetBaseException().Message}";
        }
    }

    private static string BuildObjectCardText(DemandRow row) =>
        $"1 Информация по детали\n" +
        $"IPS: {row.Ips}\n" +
        $"Наименование: {row.Name}\n" +
        $"Проект: {row.Project}\n" +
        $"№ станка: {row.MachineNumber}\n" +
        $"Количество: {row.Quantity} {row.UnitName}\n" +
        $"В производстве: {row.InProductionQuantity}\n" +
        $"Дата потребности: {row.DemandDate}";

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
        var totalInProduction = details
            .Select(x => TryParseDisplayQuantity(x.InProductionQuantity, out var quantity) ? quantity : 0m)
            .DefaultIfEmpty(0m)
            .Max();
        var deficit = Math.Max(0m, totalQuantity - totalInProduction);
        DetailTitle = $"Детализация IPS {row.Ips}; всего деталей: {FormatDecimal(totalQuantity)}; в производстве: {FormatDecimal(totalInProduction)}; дефицит: {FormatDecimal(deficit)}";
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
    IReportExportService? reportExportService = null,
    IIpsDrawingService? ipsDrawingService = null,
    ILogger? logger = null) : ObservableObject
{
    private static readonly BlankType[] MeterBasedBlankTypes =
    [
        BlankType.RoundBar,
        BlankType.SquareBar,
        BlankType.HexBar,
        BlankType.PipeRound,
        BlankType.PipeRectangular,
        BlankType.Angle
    ];

    public ObservableCollection<LibraryRow> Rows { get; } = [];
    public ObservableCollection<LibraryBlankOption> BlankSuggestions { get; } = [];

    public ObservableCollection<DisplayOption<BlankType>> BlankTypes { get; } = [..UiReferenceData.BlankTypes()];
    public ObservableCollection<DisplayOption<MeasurementUnit>> UnitTypes { get; } = [..UiReferenceData.ConsumptionUnitTypes()];
    public ObservableCollection<DisplayOption<BlankType?>> BlankTypeFilters { get; } = [..UiReferenceData.BlankTypeFilters(includeUnknown: false)];
    public ObservableCollection<DisplayOption<MeasurementUnit?>> UnitFilters { get; } = [..UiReferenceData.ConsumptionUnitFilters()];
    public IReadOnlyList<DisplayOption<bool?>> BlankStatusFilters { get; } =
    [
        new((bool?)null, "Все"),
        new(false, "Без заготовки"),
        new(true, "С заготовкой")
    ];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankTypeFilter = UiReferenceData.BlankTypeFilters(includeUnknown: false)[0];
    [ObservableProperty] private DisplayOption<MeasurementUnit?> selectedUnitFilter = UiReferenceData.ConsumptionUnitFilters()[0];
    [ObservableProperty] private DisplayOption<bool?> selectedBlankStatusFilter = new(null, "Все");
    [ObservableProperty] private bool demandOnly;
    [ObservableProperty] private string editIps = string.Empty;
    [ObservableProperty] private string editDesignation = string.Empty;
    [ObservableProperty] private string editPartName = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType> selectedBlankType = UiReferenceData.BlankTypes()[0];
    [ObservableProperty] private string blankSearch = string.Empty;
    [ObservableProperty] private LibraryBlankOption? selectedBlank;
    [ObservableProperty] private string selectedOneCCode = string.Empty;
    [ObservableProperty] private decimal consumptionQuantity = 1m;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private string blankLeadTimeDaysText = "30";
    [ObservableProperty] private bool editRequiresNitriding;
    [ObservableProperty] private bool editRequiresHeatTreatment;
    [ObservableProperty] private bool editRequiresChemicalOxidation;
    [ObservableProperty] private bool editRequiresKeyway;
    [ObservableProperty] private bool editBlankSupplyRequiresHeatTreatment;
    [ObservableProperty] private bool editBlankSupplyRequiresLaserCutting;
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiReferenceData.UnitTypes()[0];
    [ObservableProperty] private LibraryRow? selectedRow;
    [ObservableProperty] private string editSource = "Ручной ввод";
    [ObservableProperty] private string editorStatus = string.Empty;
    [ObservableProperty] private string summaryText = "Деталей: 0; без заготовки: 0";
    private bool suppressBlankSearchReload;
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();

    public string NitridingServiceLabel => UiReferenceData.ServiceName("Nitriding", "Азотирование");
    public string HeatTreatmentServiceLabel => UiReferenceData.ServiceName("HeatTreatment", "ТО");
    public string ChemicalOxidationServiceLabel => UiReferenceData.ServiceName("ChemicalOxidation", "Хим. окс");
    public string KeywayServiceLabel => UiReferenceData.ServiceName("Keyway", "Шпон паз");
    public bool IsNitridingServiceVisible => UiReferenceData.IsServiceActive("Nitriding");
    public bool IsHeatTreatmentServiceVisible => UiReferenceData.IsServiceActive("HeatTreatment");
    public bool IsChemicalOxidationServiceVisible => UiReferenceData.IsServiceActive("ChemicalOxidation");
    public bool IsKeywayServiceVisible => UiReferenceData.IsServiceActive("Keyway");
    public string BlankSupplyHeatTreatmentLabel => UiReferenceData.SupplyConditionName("HeatTreatment", "ТО");
    public string BlankSupplyLaserCuttingLabel => UiReferenceData.SupplyConditionName("LaserCutting", "Лазерная резка");
    public bool IsBlankSupplyHeatTreatmentVisible => UiReferenceData.IsSupplyConditionActive("HeatTreatment");
    public bool IsBlankSupplyLaserCuttingVisible => UiReferenceData.IsSupplyConditionActive("LaserCutting");

    partial void OnSearchChanged(string value) => _ = LoadAsync();
    partial void OnSelectedBlankTypeFilterChanged(DisplayOption<BlankType?> value) => _ = LoadAsync();
    partial void OnSelectedUnitFilterChanged(DisplayOption<MeasurementUnit?> value) => _ = LoadAsync();
    partial void OnSelectedBlankStatusFilterChanged(DisplayOption<bool?> value) => _ = LoadAsync();
    partial void OnDemandOnlyChanged(bool value) => _ = LoadAsync();
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
    partial void OnEditorStatusChanged(string value)
    {
        var cleaned = UiText.Clean(value);
        if (!string.Equals(cleaned, value, StringComparison.Ordinal))
        {
            EditorStatus = cleaned;
        }
    }

    public void RefreshReferenceLists()
    {
        var selectedBlankTypeValue = SelectedBlankType.Value;
        var selectedBlankTypeFilterValue = SelectedBlankTypeFilter.Value;
        var selectedUnitValue = SelectedUnit.Value;
        var selectedUnitFilterValue = SelectedUnitFilter.Value;

        UiReferenceData.ReplaceOptions(BlankTypes, UiReferenceData.BlankTypes());
        UiReferenceData.ReplaceOptions(UnitTypes, UiReferenceData.ConsumptionUnitTypes());
        UiReferenceData.ReplaceOptions(BlankTypeFilters, UiReferenceData.BlankTypeFilters(includeUnknown: false));
        UiReferenceData.ReplaceOptions(UnitFilters, UiReferenceData.ConsumptionUnitFilters());

        SelectedBlankType = BlankTypes.FirstOrDefault(x => x.Value == selectedBlankTypeValue) ?? BlankTypes.First();
        SelectedBlankTypeFilter = BlankTypeFilters.FirstOrDefault(x => x.Value == selectedBlankTypeFilterValue) ?? BlankTypeFilters.First();
        SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == selectedUnitValue) ?? UnitTypes.First();
        SelectedUnitFilter = UnitFilters.FirstOrDefault(x => x.Value == selectedUnitFilterValue) ?? UnitFilters.First();

        OnPropertyChanged(nameof(NitridingServiceLabel));
        OnPropertyChanged(nameof(HeatTreatmentServiceLabel));
        OnPropertyChanged(nameof(ChemicalOxidationServiceLabel));
        OnPropertyChanged(nameof(KeywayServiceLabel));
        OnPropertyChanged(nameof(IsNitridingServiceVisible));
        OnPropertyChanged(nameof(IsHeatTreatmentServiceVisible));
        OnPropertyChanged(nameof(IsChemicalOxidationServiceVisible));
        OnPropertyChanged(nameof(IsKeywayServiceVisible));
        OnPropertyChanged(nameof(BlankSupplyHeatTreatmentLabel));
        OnPropertyChanged(nameof(BlankSupplyLaserCuttingLabel));
        OnPropertyChanged(nameof(IsBlankSupplyHeatTreatmentVisible));
        OnPropertyChanged(nameof(IsBlankSupplyLaserCuttingVisible));
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
            SelectedUnit = GetConsumptionUnitOption(unit);
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
        EditRequiresNitriding = false;
        EditRequiresHeatTreatment = false;
        EditRequiresChemicalOxidation = false;
        EditRequiresKeyway = false;
        EditBlankSupplyRequiresHeatTreatment = false;
        EditBlankSupplyRequiresLaserCutting = false;
        SelectedUnit = UnitTypes[0];
        EditSource = "Ручной ввод";
        SelectedRow = null;
        EditorStatus = "Заполните поля и нажмите \"Сохранить\", чтобы добавить строку библиотеки вручную.";
    }

    [RelayCommand]
    private async Task OpenDrawingAsync(object? parameter)
    {
        var row = parameter as LibraryRow ?? SelectedRow;
        if (row is null)
        {
            EditorStatus = "Выберите строку библиотеки для открытия чертежа.";
            return;
        }

        SelectedRow = row;
        EditorStatus = await DrawingPdfOpener.OpenExternalAsync(
            new DrawingLookupRequest(row.Ips, row.Designation, row.PartName, null),
            drawingService,
            "Библиотека",
            logger,
            CancellationToken.None);
    }

    [RelayCommand]
    private async Task OpenObjectCardAsync(object? parameter)
    {
        var row = parameter as LibraryRow ?? SelectedRow;
        if (row is null)
        {
            EditorStatus = "Выберите строку библиотеки для открытия карточки объекта.";
            return;
        }

        SelectedRow = row;
        var routeRows = new ObservableCollection<ObjectCardRouteRow>(await LoadObjectCardRouteRowsAsync(row.Ips));
        var toolRows = new ObservableCollection<ObjectCardToolRow>(await LoadObjectCardToolRowsAsync(row.Ips, routeRows));
        var blankOptions = new ObservableCollection<LibraryBlankOption>(await LoadLibraryBlankOptionsForCardAsync(row.CanonicalBlankId, row.BlankType));
        PdfViewer? viewer = null;
        PdfDocument? document = null;
        WindowsFormsHost? host = null;
        FileInfo? currentDrawing = null;
        var status = new TextBlock
        {
            Text = "Поиск PDF-чертежа...",
            Foreground = Brushes.DimGray,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var ipsBox = new TextBox { Width = 140, Text = row.Ips, Margin = new Thickness(0, 0, 10, 8) };
        var designationBox = new TextBox { Width = 180, Text = row.Designation ?? string.Empty, Margin = new Thickness(0, 0, 10, 8) };
        var nameBox = new TextBox { Width = 260, Text = row.PartName, Margin = new Thickness(0, 0, 10, 8) };
        var blankTypeBox = new ComboBox { Width = 180, ItemsSource = BlankTypes, DisplayMemberPath = nameof(DisplayOption<BlankType>.DisplayName), Margin = new Thickness(0, 0, 10, 8) };
        blankTypeBox.SelectedItem = row.CanonicalBlankId is null
            ? BlankTypes.FirstOrDefault()
            : BlankTypes.FirstOrDefault(x => string.Equals(x.DisplayName, row.BlankType, StringComparison.OrdinalIgnoreCase)) ?? BlankTypes.FirstOrDefault();
        var blankBox = new ComboBox
        {
            Width = 430,
            IsEditable = true,
            IsTextSearchEnabled = false,
            StaysOpenOnEdit = true,
            ItemsSource = blankOptions,
            DisplayMemberPath = nameof(LibraryBlankOption.DisplayName),
            Margin = new Thickness(0, 0, 10, 8)
        };
        blankBox.SelectedItem = row.CanonicalBlankId is null ? null : blankOptions.FirstOrDefault(x => x.Id == row.CanonicalBlankId.Value);
        blankBox.Text = row.BlankName ?? string.Empty;
        var codeBox = new TextBox { Width = 120, Text = row.OneCCode ?? string.Empty, IsReadOnly = true, Margin = new Thickness(0, 0, 10, 8) };
        var quantityBox = new TextBox { Width = 110, Text = row.Quantity, Margin = new Thickness(0, 0, 10, 8) };
        var unitBox = new ComboBox { Width = 130, ItemsSource = UnitTypes, DisplayMemberPath = nameof(DisplayOption<MeasurementUnit>.DisplayName), Margin = new Thickness(0, 0, 10, 8) };
        unitBox.SelectedItem = UnitTypes.FirstOrDefault(x => x.Value == row.ConsumptionUnit) ?? UnitTypes.FirstOrDefault();
        var leadTimeBox = new TextBox { Width = 110, Text = row.BlankLeadTimeDaysText, Margin = new Thickness(0, 0, 10, 8) };
        var nitridingBox = new System.Windows.Controls.CheckBox { Content = NitridingServiceLabel, IsChecked = row.RequiresNitriding, Visibility = IsNitridingServiceVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var heatTreatmentBox = new System.Windows.Controls.CheckBox { Content = HeatTreatmentServiceLabel, IsChecked = row.RequiresHeatTreatment, Visibility = IsHeatTreatmentServiceVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var chemicalOxBox = new System.Windows.Controls.CheckBox { Content = ChemicalOxidationServiceLabel, IsChecked = row.RequiresChemicalOxidation, Visibility = IsChemicalOxidationServiceVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var keywayBox = new System.Windows.Controls.CheckBox { Content = KeywayServiceLabel, IsChecked = row.RequiresKeyway, Visibility = IsKeywayServiceVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var blankSupplyHeatBox = new System.Windows.Controls.CheckBox { Content = BlankSupplyHeatTreatmentLabel, IsChecked = row.BlankSupplyRequiresHeatTreatment, Visibility = IsBlankSupplyHeatTreatmentVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var blankSupplyLaserBox = new System.Windows.Controls.CheckBox { Content = BlankSupplyLaserCuttingLabel, IsChecked = row.BlankSupplyRequiresLaserCutting, Visibility = IsBlankSupplyLaserCuttingVisible ? Visibility.Visible : Visibility.Collapsed, Margin = new Thickness(0, 0, 14, 8) };
        var laborText = new TextBlock { Text = BuildObjectCardLaborText(routeRows), FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 8) };
        var cardStatus = new TextBlock { Foreground = Brushes.DimGray, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0) };

        blankBox.SelectionChanged += (_, _) =>
        {
            if (blankBox.SelectedItem is LibraryBlankOption option)
            {
                blankBox.Text = option.DisplayName;
                codeBox.Text = option.OneCCode ?? string.Empty;
                unitBox.SelectedItem = UnitTypes.FirstOrDefault(x => x.Value == (IsMeterBasedBlankType(option.BlankType) ? MeasurementUnit.Meter : option.BaseUnit)) ?? unitBox.SelectedItem;
            }
        };
        blankTypeBox.SelectionChanged += async (_, _) =>
        {
            if (blankTypeBox.SelectedItem is not DisplayOption<BlankType> type)
            {
                return;
            }

            blankOptions.Clear();
            foreach (var option in await LoadLibraryBlankOptionsForCardAsync(null, type.DisplayName))
            {
                blankOptions.Add(option);
            }
        };

        var infoPanel = new DockPanel { Margin = new Thickness(10) };
        var infoButtons = new WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
        DockPanel.SetDock(infoButtons, Dock.Bottom);
        infoPanel.Children.Add(infoButtons);
        var infoContent = new StackPanel();
        infoPanel.Children.Add(infoContent);
        infoContent.Children.Add(new TextBlock { Text = "Информация по детали", FontSize = 18, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 8) });
        var infoLine1 = new WrapPanel();
        infoLine1.Children.Add(LabeledControl("IPS", ipsBox));
        infoLine1.Children.Add(LabeledControl("Обозначение", designationBox));
        infoLine1.Children.Add(LabeledControl("Наименование", nameBox));
        infoContent.Children.Add(infoLine1);
        var infoLine2 = new WrapPanel();
        infoLine2.Children.Add(LabeledControl("Вид заготовки", blankTypeBox));
        infoLine2.Children.Add(LabeledControl("Заготовка", blankBox));
        infoLine2.Children.Add(LabeledControl("Код УТ", codeBox));
        infoContent.Children.Add(infoLine2);
        var infoLine3 = new WrapPanel();
        infoLine3.Children.Add(LabeledControl("Норма", quantityBox));
        infoLine3.Children.Add(LabeledControl("Ед. измерения", unitBox));
        infoLine3.Children.Add(LabeledControl("Срок, дней", leadTimeBox));
        infoContent.Children.Add(infoLine3);
        var servicesLine = new WrapPanel();
        servicesLine.Children.Add(new TextBlock { Text = "Услуги:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 8) });
        servicesLine.Children.Add(nitridingBox);
        servicesLine.Children.Add(heatTreatmentBox);
        servicesLine.Children.Add(chemicalOxBox);
        servicesLine.Children.Add(keywayBox);
        infoContent.Children.Add(servicesLine);
        var supplyLine = new WrapPanel();
        supplyLine.Children.Add(new TextBlock { Text = "Условие поставки заготовки:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 8) });
        supplyLine.Children.Add(blankSupplyHeatBox);
        supplyLine.Children.Add(blankSupplyLaserBox);
        infoContent.Children.Add(supplyLine);
        infoContent.Children.Add(laborText);
        infoContent.Children.Add(cardStatus);

        var saveInfoButton = new Button { Content = "Сохранить информацию" };
        infoButtons.Children.Add(saveInfoButton);

        var routeGrid = new DataGrid
        {
            ItemsSource = routeRows,
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = true,
            IsReadOnly = false,
            SelectionUnit = DataGridSelectionUnit.CellOrRowHeader,
            SelectionMode = DataGridSelectionMode.Extended,
            ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader
        };
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "№ операции", Binding = new Binding(nameof(ObjectCardRouteRow.OperationNumber)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 95 });
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(ObjectCardRouteRow.OperationName)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "Рабочий центр", Binding = new Binding(nameof(ObjectCardRouteRow.WorkCenter)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 150 });
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "Т маш", Binding = new Binding(nameof(ObjectCardRouteRow.MachineTimeText)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 80 });
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "Тнал", Binding = new Binding(nameof(ObjectCardRouteRow.SetupTimeText)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 80 });
        routeGrid.Columns.Add(new DataGridTextColumn { Header = "Твсп", Binding = new Binding(nameof(ObjectCardRouteRow.AuxiliaryTimeText)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 80 });
        var routePanel = new DockPanel { Margin = new Thickness(10) };
        var routeButtons = new WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
        DockPanel.SetDock(routeButtons, Dock.Bottom);
        routePanel.Children.Add(routeButtons);
        routePanel.Children.Add(routeGrid);
        var addRouteButton = new Button { Content = "Добавить операцию" };
        var deleteRouteButton = new Button { Content = "Удалить" };
        var upRouteButton = new Button { Content = "Выше" };
        var downRouteButton = new Button { Content = "Ниже" };
        var saveRouteButton = new Button { Content = "Сохранить маршрут обработки" };
        routeButtons.Children.Add(addRouteButton);
        routeButtons.Children.Add(deleteRouteButton);
        routeButtons.Children.Add(upRouteButton);
        routeButtons.Children.Add(downRouteButton);
        routeButtons.Children.Add(saveRouteButton);

        var toolGrid = new DataGrid
        {
            ItemsSource = toolRows,
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = true,
            IsReadOnly = false,
            SelectionUnit = DataGridSelectionUnit.CellOrRowHeader,
            SelectionMode = DataGridSelectionMode.Extended,
            ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader
        };
        toolGrid.Columns.Add(new DataGridTextColumn { Header = "№ операции", Binding = new Binding(nameof(ObjectCardToolRow.OperationNumber)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 95 });
        toolGrid.Columns.Add(new DataGridTextColumn { Header = "Операция", Binding = new Binding(nameof(ObjectCardToolRow.OperationName)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 170 });
        toolGrid.Columns.Add(new DataGridTextColumn { Header = "Инструмент/оснастка", Binding = new Binding(nameof(ObjectCardToolRow.ToolingName)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        toolGrid.Columns.Add(new DataGridTextColumn { Header = "Норма расхода", Binding = new Binding(nameof(ObjectCardToolRow.ConsumptionRate)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 130 });
        var toolPanel = new DockPanel { Margin = new Thickness(10) };
        var toolButtons = new WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
        DockPanel.SetDock(toolButtons, Dock.Bottom);
        toolPanel.Children.Add(toolButtons);
        toolPanel.Children.Add(toolGrid);
        var addToolButton = new Button { Content = "Добавить строку" };
        var deleteToolButton = new Button { Content = "Удалить" };
        var syncToolsButton = new Button { Content = "Обновить операции" };
        var saveToolsButton = new Button { Content = "Сохранить СТО и расход" };
        toolButtons.Children.Add(addToolButton);
        toolButtons.Children.Add(deleteToolButton);
        toolButtons.Children.Add(syncToolsButton);
        toolButtons.Children.Add(saveToolsButton);

        var tabs = new System.Windows.Controls.TabControl();
        tabs.Items.Add(new TabItem { Header = "Информация по детали", Content = infoPanel });
        tabs.Items.Add(new TabItem { Header = "Маршрут обработки и трудоемкость", Content = routePanel });
        tabs.Items.Add(new TabItem { Header = "СТО и расход", Content = toolPanel });

        var left = new DockPanel
        {
            LastChildFill = true
        };
        left.Children.Add(new TextBlock { Text = "Карточка объекта", FontSize = 22, FontWeight = FontWeights.SemiBold, Margin = new Thickness(12, 12, 12, 6) });
        DockPanel.SetDock(left.Children[0], Dock.Top);
        left.Children.Add(tabs);

        try
        {
            viewer = new PdfViewer
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                ZoomMode = PdfViewerZoomMode.FitWidth
            };
            host = new WindowsFormsHost { Child = viewer };
        }
        catch (Exception ex)
        {
            status.Text = $"PDF-viewer не запущен: {ex.GetBaseException().Message}";
        }

        var layout = new Grid();
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.52, GridUnitType.Star), MinWidth = 520 });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.48, GridUnitType.Star), MinWidth = 420 });
        layout.Children.Add(left);
        var fullScreenButton = new Button { Content = "Открыть на все окно", HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 8, 8) };
        var rightPanel = new DockPanel { Margin = new Thickness(10) };
        var rightTop = new DockPanel();
        DockPanel.SetDock(rightTop, Dock.Top);
        rightTop.Children.Add(fullScreenButton);
        rightTop.Children.Add(status);
        rightPanel.Children.Add(rightTop);
        rightPanel.Children.Add(host is not null ? host : new TextBlock { Text = status.Text, Margin = new Thickness(16), TextWrapping = TextWrapping.Wrap });
        var right = new Border
        {
            BorderBrush = Brushes.LightGray,
            BorderThickness = new Thickness(1, 0, 0, 0),
            Child = rightPanel
        };
        Grid.SetColumn(right, 1);
        layout.Children.Add(right);

        var window = new Window
        {
            Title = $"Карточка объекта IPS {row.Ips}",
            Owner = Application.Current?.MainWindow,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 1180,
            Height = 760,
            MinWidth = 920,
            MinHeight = 620,
            Content = layout
        };
        saveInfoButton.Click += async (_, _) =>
        {
            await SaveObjectCardInfoAsync(
                row,
                ipsBox.Text,
                designationBox.Text,
                nameBox.Text,
                blankTypeBox.SelectedItem as DisplayOption<BlankType>,
                blankBox.SelectedItem as LibraryBlankOption,
                blankBox.Text,
                quantityBox.Text,
                unitBox.SelectedItem as DisplayOption<MeasurementUnit>,
                leadTimeBox.Text,
                nitridingBox.IsChecked.GetValueOrDefault(),
                heatTreatmentBox.IsChecked.GetValueOrDefault(),
                chemicalOxBox.IsChecked.GetValueOrDefault(),
                keywayBox.IsChecked.GetValueOrDefault(),
                blankSupplyHeatBox.IsChecked.GetValueOrDefault(),
                blankSupplyLaserBox.IsChecked.GetValueOrDefault(),
                cardStatus);
        };
        addRouteButton.Click += (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var next = routeRows.Count == 0 ? 10 : routeRows.Select(x => x.SequenceValue).DefaultIfEmpty(0).Max() + 10;
            routeRows.Add(new ObjectCardRouteRow(next.ToString(CultureInfo.InvariantCulture), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty));
            laborText.Text = BuildObjectCardLaborText(routeRows);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        deleteRouteButton.Click += (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            RemoveSelectedRows(routeGrid, routeRows);
            laborText.Text = BuildObjectCardLaborText(routeRows);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        upRouteButton.Click += (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            MoveSelectedRow(routeGrid, routeRows, -1);
            RenumberRouteRows(routeRows);
            laborText.Text = BuildObjectCardLaborText(routeRows);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        downRouteButton.Click += (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            MoveSelectedRow(routeGrid, routeRows, 1);
            RenumberRouteRows(routeRows);
            laborText.Text = BuildObjectCardLaborText(routeRows);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        saveRouteButton.Click += async (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            await SaveObjectCardRouteRowsAsync(ipsBox.Text, routeRows, cardStatus);
            laborText.Text = BuildObjectCardLaborText(routeRows);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        addToolButton.Click += (_, _) =>
        {
            toolGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var route = routeRows.FirstOrDefault();
            toolRows.Add(new ObjectCardToolRow(route?.OperationNumber ?? string.Empty, route?.OperationName ?? string.Empty, string.Empty, string.Empty));
        };
        deleteToolButton.Click += (_, _) =>
        {
            toolGrid.CommitEdit(DataGridEditingUnit.Row, true);
            RemoveSelectedRows(toolGrid, toolRows);
        };
        syncToolsButton.Click += (_, _) =>
        {
            routeGrid.CommitEdit(DataGridEditingUnit.Row, true);
            SyncToolRowsWithRoute(toolRows, routeRows);
        };
        saveToolsButton.Click += async (_, _) =>
        {
            toolGrid.CommitEdit(DataGridEditingUnit.Row, true);
            await SaveObjectCardToolRowsAsync(ipsBox.Text, toolRows, cardStatus);
        };
        fullScreenButton.Click += (_, _) =>
        {
            if (currentDrawing is null)
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = "PDF-чертеж еще не открыт.";
                return;
            }

            ShowPdfFullScreen(currentDrawing, row);
        };
        window.Closed += (_, _) =>
        {
            if (viewer is not null)
            {
                viewer.Document = null;
            }

            if (host is not null)
            {
                host.Child = null;
            }

            viewer?.Dispose();
            document?.Dispose();
        };
        window.Show();

        if (viewer is null)
        {
            return;
        }

        currentDrawing = await DrawingPdfOpener.ResolveDrawingPdfAsync(
            new DrawingLookupRequest(row.Ips, row.Designation, row.PartName, null),
            drawingService,
            logger,
            CancellationToken.None);
        if (currentDrawing is null)
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = $"PDF-чертеж IPS {row.Ips} не найден.";
            return;
        }

        try
        {
            document = PdfDocument.Load(currentDrawing.FullName);
            viewer.Document = document;
            viewer.ZoomMode = PdfViewerZoomMode.FitWidth;
            status.Foreground = Brushes.SeaGreen;
            status.Text = $"PDF-чертеж открыт: {currentDrawing.Name}";
        }
        catch (Exception ex)
        {
            document?.Dispose();
            document = null;
            status.Foreground = Brushes.Firebrick;
            status.Text = $"PDF найден, но viewer не смог открыть файл: {ex.GetBaseException().Message}";
        }
    }

    private async Task<IReadOnlyList<LibraryBlankOption>> LoadLibraryBlankOptionsForCardAsync(long? selectedBlankId, string? blankTypeName)
    {
        var selectedType = BlankTypes.FirstOrDefault(x => string.Equals(x.DisplayName, blankTypeName, StringComparison.OrdinalIgnoreCase))?.Value;
        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);
        if (selectedType is not null && selectedType != BlankType.Unknown)
        {
            query = query.Where(x => x.BlankType == selectedType);
        }

        var options = await query
            .OrderBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Take(5000)
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .ToListAsync();

        if (selectedBlankId is not null && options.All(x => x.Id != selectedBlankId.Value))
        {
            var selected = await dbContext.CanonicalBlanks.AsNoTracking()
                .Where(x => x.Id == selectedBlankId.Value)
                .Select(x => new LibraryBlankOption(
                    x.Id,
                    x.BlankType,
                    x.CanonicalName,
                    x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                    x.Material,
                    x.BaseUnit,
                    x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
                .FirstOrDefaultAsync();
            if (selected is not null)
            {
                options.Insert(0, selected);
            }
        }

        return options;
    }

    private async Task<LibraryBlankOption?> ResolveBlankOptionForCardAsync(string? text, BlankType? selectedType)
    {
        var searchText = text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(searchText))
        {
            return null;
        }

        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive);
        if (selectedType is not null && selectedType != BlankType.Unknown)
        {
            query = query.Where(x => x.BlankType == selectedType.Value);
        }

        var blanks = await query.Take(5000).ToListAsync();
        return blanks
            .Where(x => BlankMatchesSearch(x, searchText))
            .OrderBy(x => x.Aliases.Any(a => UiSearchText.EqualsNormalized(a.OneCCode, searchText)) ? 0 : 1)
            .ThenBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .FirstOrDefault();
    }

    private async Task SaveObjectCardInfoAsync(
        LibraryRow sourceRow,
        string ips,
        string designation,
        string name,
        DisplayOption<BlankType>? blankType,
        LibraryBlankOption? selectedBlank,
        string blankText,
        string quantityText,
        DisplayOption<MeasurementUnit>? unit,
        string leadTimeText,
        bool requiresNitriding,
        bool requiresHeatTreatment,
        bool requiresChemicalOxidation,
        bool requiresKeyway,
        bool supplyRequiresHeatTreatment,
        bool supplyRequiresLaserCutting,
        TextBlock status)
    {
        ips = ips.Trim();
        if (string.IsNullOrWhiteSpace(ips))
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = "Укажите IPS.";
            return;
        }

        if (string.IsNullOrWhiteSpace(name))
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = "Укажите наименование детали.";
            return;
        }

        var duplicatePart = await dbContext.Parts.AsNoTracking()
            .FirstOrDefaultAsync(x => x.Ips == ips && x.Id != sourceRow.PartId);
        if (duplicatePart is not null)
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = $"IPS {ips} уже есть в библиотеке у другой карточки.";
            return;
        }

        var part = await dbContext.Parts
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .FirstOrDefaultAsync(x => x.Id == sourceRow.PartId);
        if (part is null)
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = "Деталь не найдена в библиотеке.";
            return;
        }

        part.Ips = ips;
        part.Designation = NullIfWhiteSpace(designation);
        part.Name = name.Trim();
        part.RequiresNitriding = requiresNitriding;
        part.RequiresHeatTreatment = requiresHeatTreatment;
        part.RequiresChemicalOxidation = requiresChemicalOxidation;
        part.RequiresKeyway = requiresKeyway;
        part.BlankSupplyRequiresHeatTreatment = supplyRequiresHeatTreatment;
        part.BlankSupplyRequiresLaserCutting = supplyRequiresLaserCutting;
        part.UpdatedAt = DateTime.UtcNow;

        var blank = selectedBlank ?? await ResolveBlankOptionForCardAsync(blankText, blankType?.Value);
        if (blank is not null)
        {
            if (!TryParseQuantity(quantityText, out var quantity) || quantity <= 0)
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = "Норма расхода должна быть положительным числом.";
                return;
            }

            if (!int.TryParse(leadTimeText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var leadTimeDays) || leadTimeDays < 0)
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = "Срок заготовки должен быть целым числом дней.";
                return;
            }

            var selectedUnit = unit ?? UnitTypes.FirstOrDefault();
            if (selectedUnit is null)
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = "Укажите единицу измерения.";
                return;
            }

            var storedQuantity = ToStoredConsumptionQuantity(quantity, selectedUnit);
            var sameMap = part.BlankMaps.FirstOrDefault(x => x.CanonicalBlankId == blank.Id && x.IsActive);
            foreach (var activeMap in part.BlankMaps.Where(x => x.IsActive && (sameMap is null || x.Id != sameMap.Id)))
            {
                activeMap.IsActive = false;
                activeMap.IsPrimary = false;
                activeMap.UpdatedAt = DateTime.UtcNow;
            }

            if (sameMap is null)
            {
                dbContext.PartBlankMaps.Add(new PartBlankMap
                {
                    Part = part,
                    CanonicalBlankId = blank.Id,
                    ConsumptionQuantity = storedQuantity,
                    ConsumptionUnit = selectedUnit.Value,
                    BlankLeadTimeDays = leadTimeDays,
                    Source = "Карточка объекта",
                    IsPrimary = true,
                    IsActive = true
                });
            }
            else
            {
                sameMap.ConsumptionQuantity = storedQuantity;
                sameMap.ConsumptionUnit = selectedUnit.Value;
                sameMap.BlankLeadTimeDays = leadTimeDays;
                sameMap.Source = string.IsNullOrWhiteSpace(sameMap.Source) ? "Карточка объекта" : sameMap.Source;
                sameMap.IsPrimary = true;
                sameMap.UpdatedAt = DateTime.UtcNow;
            }
        }

        await dbContext.SaveChangesAsync();
        await LoadAsync();
        status.Foreground = Brushes.SeaGreen;
        status.Text = $"Информация по детали IPS {ips} сохранена и синхронизирована с библиотекой.";
    }

    private async Task<IReadOnlyList<ObjectCardRouteRow>> LoadObjectCardRouteRowsAsync(string ips)
    {
        var rows = await dbContext.ProductionRouteOperations.AsNoTracking()
            .Where(x => x.Ips == ips)
            .OrderBy(x => x.Sequence)
            .ToListAsync();
        return rows
            .Select(x => new ObjectCardRouteRow(
                string.IsNullOrWhiteSpace(x.OperationCode) ? x.Sequence.ToString(CultureInfo.InvariantCulture) : x.OperationCode,
                UiText.Clean(x.Description),
                UiText.Clean(x.EquipmentGroup),
                FormatDecimal(x.MachineMinutes),
                FormatDecimal(x.SetupMinutes),
                FormatDecimal(x.AuxiliaryMinutes)))
            .ToList();
    }

    private async Task SaveObjectCardRouteRowsAsync(string ips, ObservableCollection<ObjectCardRouteRow> rows, TextBlock status)
    {
        ips = ips.Trim();
        if (string.IsNullOrWhiteSpace(ips))
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = "Укажите IPS перед сохранением маршрута.";
            return;
        }

        var existing = await dbContext.ProductionRouteOperations.Where(x => x.Ips == ips).ToListAsync();
        dbContext.ProductionRouteOperations.RemoveRange(existing);
        var usedSequences = new HashSet<int>();
        var index = 0;
        foreach (var row in rows.Where(x => !x.IsEmpty))
        {
            index++;
            var sequence = row.SequenceValue;
            if (sequence <= 0 || !usedSequences.Add(sequence))
            {
                sequence = index * 10;
                while (!usedSequences.Add(sequence))
                {
                    sequence += 10;
                }
            }

            dbContext.ProductionRouteOperations.Add(new ProductionRouteOperation
            {
                Ips = ips,
                Sequence = sequence,
                OperationCode = row.OperationNumber.Trim(),
                Description = row.OperationName.Trim(),
                EquipmentGroup = row.WorkCenter.Trim(),
                MachineMinutes = ParseOptionalDecimal(row.MachineTimeText),
                PieceMinutes = ParseOptionalDecimal(row.MachineTimeText),
                SetupMinutes = ParseOptionalDecimal(row.SetupTimeText),
                AuxiliaryMinutes = ParseOptionalDecimal(row.AuxiliaryTimeText),
                UpdatedAt = DateTime.UtcNow
            });
        }

        await dbContext.SaveChangesAsync();
        status.Foreground = Brushes.SeaGreen;
        status.Text = $"Маршрут обработки IPS {ips} сохранен. Операций: {rows.Count(x => !x.IsEmpty)}.";
    }

    private async Task<IReadOnlyList<ObjectCardToolRow>> LoadObjectCardToolRowsAsync(string ips, IReadOnlyList<ObjectCardRouteRow> routeRows)
    {
        var json = await dbContext.Settings.AsNoTracking()
            .Where(x => x.Key == BuildObjectCardToolsSettingKey(ips))
            .Select(x => x.Value)
            .FirstOrDefaultAsync();
        var rows = string.IsNullOrWhiteSpace(json)
            ? []
            : JsonSerializer.Deserialize<List<ObjectCardToolRowDto>>(json) ?? [];
        var result = rows
            .Select(x => new ObjectCardToolRow(x.OperationNumber, x.OperationName, x.ToolingName, x.ConsumptionRate))
            .ToList();
        if (result.Count == 0)
        {
            result.AddRange(routeRows.Select(x => new ObjectCardToolRow(x.OperationNumber, x.OperationName, string.Empty, string.Empty)));
        }

        return result;
    }

    private async Task SaveObjectCardToolRowsAsync(string ips, ObservableCollection<ObjectCardToolRow> rows, TextBlock status)
    {
        ips = ips.Trim();
        if (string.IsNullOrWhiteSpace(ips))
        {
            status.Foreground = Brushes.Firebrick;
            status.Text = "Укажите IPS перед сохранением СТО и расхода.";
            return;
        }

        var dto = rows
            .Where(x => !x.IsEmpty)
            .Select(x => new ObjectCardToolRowDto(x.OperationNumber.Trim(), x.OperationName.Trim(), x.ToolingName.Trim(), x.ConsumptionRate.Trim()))
            .ToList();
        var json = JsonSerializer.Serialize(dto);
        var key = BuildObjectCardToolsSettingKey(ips);
        var setting = await dbContext.Settings.FirstOrDefaultAsync(x => x.Key == key);
        if (setting is null)
        {
            dbContext.Settings.Add(new AppSetting { Key = key, Value = json });
        }
        else
        {
            setting.Value = json;
        }

        await dbContext.SaveChangesAsync();
        status.Foreground = Brushes.SeaGreen;
        status.Text = $"СТО и расход IPS {ips} сохранены. Строк: {dto.Count}.";
    }

    private static void SyncToolRowsWithRoute(ObservableCollection<ObjectCardToolRow> toolRows, IEnumerable<ObjectCardRouteRow> routeRows)
    {
        foreach (var route in routeRows.Where(x => !x.IsEmpty))
        {
            if (toolRows.Any(x => string.Equals(x.OperationNumber, route.OperationNumber, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            toolRows.Add(new ObjectCardToolRow(route.OperationNumber, route.OperationName, string.Empty, string.Empty));
        }
    }

    private static void RenumberRouteRows(ObservableCollection<ObjectCardRouteRow> rows)
    {
        for (var index = 0; index < rows.Count; index++)
        {
            if (rows[index].SequenceValue <= 0)
            {
                rows[index].OperationNumber = ((index + 1) * 10).ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    private static string BuildObjectCardLaborText(IEnumerable<ObjectCardRouteRow> rows)
    {
        var routeRows = rows.Where(x => !x.IsEmpty).ToList();
        var machine = routeRows.Sum(x => ParseOptionalDecimal(x.MachineTimeText));
        var setup = routeRows.Sum(x => ParseOptionalDecimal(x.SetupTimeText));
        var auxiliary = routeRows.Sum(x => ParseOptionalDecimal(x.AuxiliaryTimeText));
        var total = machine + setup + auxiliary;
        return $"Трудоемкость изготовления: Т маш {FormatDecimal(machine)} мин; Тнал {FormatDecimal(setup)} мин; Твсп {FormatDecimal(auxiliary)} мин; всего {FormatDecimal(total)} мин.";
    }

    private static StackPanel LabeledControl(string label, System.Windows.Controls.Control control) =>
        new()
        {
            Children =
            {
                new TextBlock { Text = label },
                control
            }
        };

    private static void RemoveSelectedRows<T>(DataGrid grid, ObservableCollection<T> rows)
    {
        var selected = grid.SelectedItems.OfType<T>().ToList();
        if (selected.Count == 0 && grid.CurrentItem is T current)
        {
            selected.Add(current);
        }

        foreach (var row in selected)
        {
            rows.Remove(row);
        }
    }

    private static void MoveSelectedRow<T>(DataGrid grid, ObservableCollection<T> rows, int direction)
    {
        if (grid.CurrentItem is not T row)
        {
            return;
        }

        var index = rows.IndexOf(row);
        var target = index + direction;
        if (index < 0 || target < 0 || target >= rows.Count)
        {
            return;
        }

        rows.Move(index, target);
        grid.SelectedItem = row;
    }

    private static void ShowPdfFullScreen(FileInfo drawing, LibraryRow row)
    {
        PdfViewer? fullViewer = null;
        PdfDocument? fullDocument = null;
        WindowsFormsHost? fullHost = null;
        try
        {
            fullDocument = PdfDocument.Load(drawing.FullName);
            fullViewer = new PdfViewer
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                ZoomMode = PdfViewerZoomMode.FitWidth,
                Document = fullDocument
            };
            fullHost = new WindowsFormsHost { Child = fullViewer };
            var window = new Window
            {
                Title = $"Чертеж IPS {row.Ips}",
                Owner = Application.Current?.MainWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowState = WindowState.Maximized,
                Content = fullHost
            };
            window.Closed += (_, _) =>
            {
                if (fullViewer is not null)
                {
                    fullViewer.Document = null;
                }

                if (fullHost is not null)
                {
                    fullHost.Child = null;
                }

                fullViewer?.Dispose();
                fullDocument?.Dispose();
            };
            window.Show();
        }
        catch (Exception ex)
        {
            fullViewer?.Dispose();
            fullDocument?.Dispose();
            MessageBox.Show($"Не удалось открыть чертеж на все окно.\n{ex.GetBaseException().Message}", "Карточка объекта", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string BuildObjectCardToolsSettingKey(string ips) =>
        $"ObjectCard.Tools.{StockCodeNormalizer.NormalizeForComparison(ips)}";

    private static decimal ParseOptionalDecimal(string? value) =>
        TryParseQuantity(value, out var quantity) ? quantity : 0m;

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
        var searchValue = Search.Trim();

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

        HashSet<string>? demandKeys = null;
        if (DemandOnly)
        {
            var latestDemandBatchId = await dbContext.DemandBatches.AsNoTracking()
                .OrderByDescending(x => x.ImportedAt)
                .ThenByDescending(x => x.Id)
                .Select(x => (long?)x.Id)
                .FirstOrDefaultAsync();
            demandKeys = latestDemandBatchId is null
                ? []
                : (await dbContext.DemandItems.AsNoTracking()
                    .Where(x => x.DemandBatchId == latestDemandBatchId.Value && x.Ips != "")
                    .Select(x => x.Ips)
                    .Distinct()
                    .ToListAsync())
                    .Select(StockCodeNormalizer.NormalizeForComparison)
                    .Where(x => x.Length > 0)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var demandPartCodes = demandKeys
                .SelectMany(x => new[] { x, x.All(char.IsDigit) ? x.TrimStart('0') : x })
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            query = query.Where(x => demandPartCodes.Contains(x.Ips));
        }

        var parts = await query
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .OrderBy(x => x.Ips)
            .Take(string.IsNullOrWhiteSpace(searchValue) ? 1000 : 5000)
            .ToListAsync();
        if (!string.IsNullOrWhiteSpace(searchValue))
        {
            parts = parts
                .Where(PartMatchesSearch)
                .Take(1000)
                .ToList();
        }

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
                BuildExternalServiceNote(part),
                part.RequiresNitriding,
                part.RequiresHeatTreatment,
                part.RequiresChemicalOxidation,
                part.RequiresKeyway,
                BuildBlankSupplyRequirement(part),
                part.BlankSupplyRequiresHeatTreatment,
                part.BlankSupplyRequiresLaserCutting,
                blank is null ? null : DisplayBlankType(blank.BlankType),
                UiText.Clean(activeAlias?.SourceName ?? blank?.CanonicalName),
                UiText.Clean(blank?.Material),
                activeAlias?.OneCCode,
                ToDisplayConsumptionQuantity(map?.ConsumptionQuantity, effectiveUnit),
                effectiveUnit,
                FormatDecimal(ToDisplayConsumptionQuantity(map?.ConsumptionQuantity, effectiveUnit)),
                effectiveUnit is null ? string.Empty : DisplayConsumptionUnit(effectiveUnit.Value),
                map?.BlankLeadTimeDays ?? 30,
                (map?.BlankLeadTimeDays ?? 30).ToString(CultureInfo.InvariantCulture),
                UiText.Clean(part.Source),
                part.UpdatedAt.ToLocalTime().ToString("g"),
                map is null ||
                    blank is null ||
                    activeAlias is null ||
                    string.IsNullOrWhiteSpace(activeAlias.OneCCode) ||
                    map.ConsumptionQuantity <= 0));
        }

        SummaryText = DemandOnly
            ? $"Деталей: {Rows.Select(x => x.PartId).Distinct().Count()}; без заготовки: {Rows.Count(x => x.PartBlankMapId is null || x.CanonicalBlankId is null)}; в потребности: {demandKeys?.Count ?? 0}"
            : $"Деталей: {Rows.Select(x => x.PartId).Distinct().Count()}; без заготовки: {Rows.Count(x => x.PartBlankMapId is null || x.CanonicalBlankId is null)}";
        await LoadBlankSuggestionsAsync();
    }

    private bool PartMatchesSearch(Part part)
    {
        var searchValue = Search.Trim();
        var blankValues = part.BlankMaps
            .Where(m => m.IsActive && m.CanonicalBlank is not null)
            .SelectMany(m => new[]
            {
                m.CanonicalBlank!.CanonicalName,
                m.CanonicalBlank.Material
            }.Concat(m.CanonicalBlank.Aliases.SelectMany(a => new[] { a.OneCCode, a.SourceName })));
        return UiSearchText.ContainsAnyField(searchValue,
            [part.Ips, part.Designation, part.Name, BuildExternalServiceNote(part), ..blankValues]);
    }

    private static string BuildExternalServiceNote(Part part)
    {
        var notes = new List<string>(4);
        if (part.RequiresNitriding)
        {
            notes.Add("Азотирование");
        }

        if (part.RequiresHeatTreatment)
        {
            notes.Add("ТО");
        }

        if (part.RequiresChemicalOxidation)
        {
            notes.Add("Хим. окс");
        }

        if (part.RequiresKeyway)
        {
            notes.Add("Шпон паз");
        }

        return string.Join("; ", notes);
    }

    private static string BuildBlankSupplyRequirement(Part part)
    {
        var notes = new List<string>(2);
        if (part.BlankSupplyRequiresHeatTreatment)
        {
            notes.Add("ТО");
        }

        if (part.BlankSupplyRequiresLaserCutting)
        {
            notes.Add("Лазерная резка");
        }

        return string.Join("; ", notes);
    }

    private static string BuildObjectCardText(LibraryRow row) =>
        $"1 Информация по детали\n" +
        $"IPS: {row.Ips}\n" +
        $"Обозначение: {row.Designation}\n" +
        $"Наименование: {row.PartName}\n" +
        $"Вид заготовки: {row.BlankType}\n" +
        $"Заготовка: {row.BlankName}\n" +
        $"Материал: {row.Material}\n" +
        $"Условие поставки заготовки: {row.BlankSupplyRequirement}\n" +
        $"Код УТ: {row.OneCCode}\n" +
        $"Норма расхода: {row.Quantity} {row.UnitName}\n" +
        $"Срок заготовки, дней: {row.BlankLeadTimeDaysText}\n" +
        $"Услуги: {row.Note}\n" +
        $"Источник: {row.Source}\n" +
        $"Обновлено: {row.UpdatedAt}";

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

        var existingParts = await dbContext.Parts.ToListAsync();
        var partsByIps = existingParts
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(p => p.UpdatedAt).First(), StringComparer.OrdinalIgnoreCase);
        var mskRecords = await dbContext.MskRecords.AsNoTracking().ToListAsync();
        var mskByIps = mskRecords
            .Where(x => !string.IsNullOrWhiteSpace(x.Ips))
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(r => r.ImportedAt).First(), StringComparer.OrdinalIgnoreCase);
        var mskDetails = MskViewModel.LoadMskDetailsFromReport();
        var added = 0;
        foreach (var demandPart in demandParts)
        {
            mskByIps.TryGetValue(demandPart.Ips, out var mskRecord);
            var mskDetail = MskViewModel.ResolveMskDetail(mskRecord, mskDetails);
            var (designation, name) = SplitDesignationAndName(FirstNotEmpty(mskRecord?.Name, demandPart.Name));
            if (partsByIps.TryGetValue(demandPart.Ips, out var existingPart))
            {
                if (!IsArchivedLibrarySource(existingPart.Source))
                {
                    continue;
                }

                if (IsManualLibraryDeleteSource(existingPart.Source))
                {
                    continue;
                }

                existingPart.Source = mskRecord is null ? "Потребность" : "Потребность; МСК";
                existingPart.Designation = FirstNotEmpty(mskRecord?.Designation, existingPart.Designation, designation);
                existingPart.Name = FirstNotEmpty(name, existingPart.Name, mskRecord?.Name, demandPart.Name, "Из потребности");
                existingPart.HasMsk = existingPart.HasMsk || mskRecord is not null;
                existingPart.UpdatedAt = DateTime.UtcNow;
                await TryAttachMskBlankAsync(existingPart, mskDetail);
                added++;
                continue;
            }

            var part = new Part
            {
                Ips = demandPart.Ips,
                Designation = FirstNotEmpty(mskRecord?.Designation, designation),
                Name = FirstNotEmpty(name, mskRecord?.Name, demandPart.Name, "Из потребности"),
                HasMsk = mskRecord is not null,
                Source = mskRecord is null ? "Потребность" : "Потребность; МСК"
            };
            dbContext.Parts.Add(part);
            partsByIps[part.Ips] = part;
            await TryAttachMskBlankAsync(part, mskDetail);
            added++;
        }

        if (added > 0)
        {
            await dbContext.SaveChangesAsync();
            EditorStatus = $"Добавлено или возвращено деталей из потребности: {added}.";
        }
    }

    private async Task TryAttachMskBlankAsync(Part part, MskCsvDetail? mskDetail)
    {
        if (mskDetail is null || string.IsNullOrWhiteSpace(mskDetail.OneCCode))
        {
            return;
        }

        var alias = await dbContext.BlankAliases
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive)
            .FirstOrDefaultAsync(x => x.OneCCode == mskDetail.OneCCode);
        if (alias?.CanonicalBlank is null)
        {
            return;
        }

        var quantity = decimal.TryParse(
            (mskDetail.ConsumptionQuantity ?? string.Empty).Trim().Replace('.', ','),
            NumberStyles.Number,
            CultureInfo.GetCultureInfo("ru-RU"),
            out var parsedQuantity)
            ? parsedQuantity
            : 1m;
        dbContext.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = alias.CanonicalBlank,
            ConsumptionQuantity = quantity > 0 ? quantity : 1m,
            ConsumptionUnit = ParseConsumptionUnit(mskDetail.UnitName, alias.CanonicalBlank.BaseUnit),
            BlankLeadTimeDays = 30,
            IsPrimary = true,
            IsActive = true,
            Source = "Потребность; МСК; НСИ"
        });
    }

    [RelayCommand]
    private async Task LoadBlankSuggestionsAsync()
    {
        var query = dbContext.CanonicalBlanks.AsNoTracking()
            .Include(x => x.Aliases)
            .Where(x => x.IsActive && x.Aliases.Any(a => a.IsActive));

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

            if (useDetectedShapeFilters && normalized.Diameter is not null && ShouldApplyExactDimensionFilter(searchText, 'D'))
            {
                query = query.Where(x => x.DiameterMm == normalized.Diameter);
            }

            if (useDetectedShapeFilters && normalized.Width is not null && ShouldApplyExactDimensionFilter(searchText, 'W'))
            {
                query = query.Where(x => x.WidthMm == normalized.Width);
            }

            if (useDetectedShapeFilters && !string.IsNullOrWhiteSpace(normalized.Material))
            {
                var material = NormalizeText(normalized.Material);
                query = query.Where(x => x.Material != null && x.Material.ToUpper().Replace(" ", "").Contains(material));
            }

        }

        var blanks = await query
            .OrderBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Take(string.IsNullOrWhiteSpace(searchText) ? 50 : 5000)
            .ToListAsync();
        if (!string.IsNullOrWhiteSpace(searchText))
        {
            var parts = searchText.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var namePart = parts.FirstOrDefault();
            var codePart = parts.Length > 1 ? parts[^1] : null;
            blanks = blanks
                .Where(x => BlankMatchesSearch(x, searchText, namePart, codePart))
                .Take(50)
                .ToList();
        }

        var options = blanks
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
                RequiresNitriding = EditRequiresNitriding,
                RequiresHeatTreatment = EditRequiresHeatTreatment,
                RequiresChemicalOxidation = EditRequiresChemicalOxidation,
                RequiresKeyway = EditRequiresKeyway,
                BlankSupplyRequiresHeatTreatment = EditBlankSupplyRequiresHeatTreatment,
                BlankSupplyRequiresLaserCutting = EditBlankSupplyRequiresLaserCutting,
                Source = NullIfWhiteSpace(EditSource) ?? "Ручной ввод"
            };
            dbContext.Parts.Add(part);
        }
        else
        {
            part.Designation = NullIfWhiteSpace(EditDesignation);
            part.Name = EditPartName.Trim();
            part.RequiresNitriding = EditRequiresNitriding;
            part.RequiresHeatTreatment = EditRequiresHeatTreatment;
            part.RequiresChemicalOxidation = EditRequiresChemicalOxidation;
            part.RequiresKeyway = EditRequiresKeyway;
            part.BlankSupplyRequiresHeatTreatment = EditBlankSupplyRequiresHeatTreatment;
            part.BlankSupplyRequiresLaserCutting = EditBlankSupplyRequiresLaserCutting;
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

        var storedQuantity = ToStoredConsumptionQuantity(quantity, SelectedUnit);
        var storedUnit = SelectedUnit.Value;
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
                ConsumptionQuantity = storedQuantity,
                ConsumptionUnit = storedUnit,
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

            sameActiveMap.ConsumptionQuantity = storedQuantity;
            sameActiveMap.ConsumptionUnit = storedUnit;
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

        var blanks = await query
            .OrderBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Take(5000)
            .ToListAsync();

        return blanks
            .Where(x => BlankMatchesSearch(x, searchText, namePart, codePart))
            .OrderBy(x => x.Aliases.Any(a => UiSearchText.EqualsNormalized(a.OneCCode, codePart)) ? 0 : 1)
            .ThenBy(x => x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault() ?? x.CanonicalName)
            .Select(x => new LibraryBlankOption(
                x.Id,
                x.BlankType,
                x.CanonicalName,
                x.Aliases.Where(a => a.IsActive).Select(a => a.SourceName).FirstOrDefault(),
                x.Material,
                x.BaseUnit,
                x.Aliases.Where(a => a.IsActive).Select(a => a.OneCCode).FirstOrDefault()))
            .FirstOrDefault();
    }

    private static bool BlankMatchesSearch(CanonicalBlank blank, string searchText, string? namePart = null, string? codePart = null) =>
        UiSearchText.Contains(blank.CanonicalName, searchText) ||
        UiSearchText.Contains(blank.Material, searchText) ||
        blank.Aliases.Any(a => UiSearchText.Contains(a.OneCCode, searchText) || UiSearchText.Contains(a.SourceName, searchText)) ||
        BlankMatchesAllSearchTokens(blank, searchText) ||
        (!string.IsNullOrWhiteSpace(namePart) && (UiSearchText.Contains(blank.CanonicalName, namePart) || blank.Aliases.Any(a => UiSearchText.Contains(a.SourceName, namePart)))) ||
        (!string.IsNullOrWhiteSpace(codePart) && blank.Aliases.Any(a => UiSearchText.Contains(a.OneCCode, codePart) || UiSearchText.Contains(a.SourceName, codePart)));

    private static bool ShouldApplyExactDimensionFilter(string searchText, char marker)
    {
        var match = Regex.Match(searchText, $@"(?<![\p{{L}}\p{{N}}]){marker}\s*(\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return true;
        }

        var value = match.Groups[1].Value;
        return value.Contains('.') || value.Contains(',') || value.Count(char.IsDigit) >= 3;
    }

    private static bool BlankMatchesAllSearchTokens(CanonicalBlank blank, string searchText)
    {
        var tokens = searchText
            .Split([' ', '\t', '\r', '\n', '|', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(x => x.Length > 0)
            .ToArray();

        return tokens.Length > 1 && tokens.All(token =>
            UiSearchText.Contains(blank.CanonicalName, token) ||
            UiSearchText.Contains(blank.Material, token) ||
            blank.Aliases.Any(a => UiSearchText.Contains(a.OneCCode, token) || UiSearchText.Contains(a.SourceName, token)));
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

        SelectedRow = row;
        EditIps = row.Ips;
        EditDesignation = row.Designation ?? string.Empty;
        EditRequiresNitriding = row.RequiresNitriding;
        EditRequiresHeatTreatment = row.RequiresHeatTreatment;
        EditRequiresChemicalOxidation = row.RequiresChemicalOxidation;
        EditRequiresKeyway = row.RequiresKeyway;
        EditBlankSupplyRequiresHeatTreatment = row.BlankSupplyRequiresHeatTreatment;
        EditBlankSupplyRequiresLaserCutting = row.BlankSupplyRequiresLaserCutting;
        EditPartName = row.PartName;
        ConsumptionQuantity = row.ConsumptionQuantity ?? 1m;
        ConsumptionQuantityText = FormatDecimal(ConsumptionQuantity);
        BlankLeadTimeDaysText = row.BlankLeadTimeDays.ToString(CultureInfo.InvariantCulture);
        SelectedUnit = GetConsumptionUnitOption(row.ConsumptionUnit);
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
            foreach (var map in part.BlankMaps.Where(x => x.IsActive))
            {
                map.IsActive = false;
                map.IsPrimary = false;
                map.UpdatedAt = DateTime.UtcNow;
            }

            var demandItems = await dbContext.DemandItems.Where(x => x.PartId == part.Id).ToListAsync();
            foreach (var demandItem in demandItems)
            {
                demandItem.PartId = null;
            }

            part.Source = BuildArchivedSource(part.Source, "[MANUAL_LIBRARY_DELETE] Удалено из библиотеки");
            part.UpdatedAt = DateTime.UtcNow;
            EditorStatus = $"Деталь убрана из библиотеки: {row.Ips}.";
            await dbContext.SaveChangesAsync();
            if (reload)
            {
                await LoadAsync();
            }
            return;
        }
        else if (row.CanonicalBlankId is null)
        {
            var demandItems = await dbContext.DemandItems.Where(x => x.PartId == part.Id).ToListAsync();
            foreach (var demandItem in demandItems)
            {
                demandItem.PartId = null;
            }

            part.Source = BuildArchivedSource(part.Source, "[MANUAL_LIBRARY_DELETE] Удалено из библиотеки");
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
        if (IsManualLibraryDeleteSource(item.Part?.Source))
        {
            return false;
        }

        var unit = UiText.Clean(item.Unit).Trim();
        if (!string.IsNullOrWhiteSpace(unit) && !unit.Contains("шт", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var ips = UiText.Clean(item.Ips).Trim();
        if (!IsValidDemandPartIps(ips))
        {
            return false;
        }

        var name = UiText.Clean(item.SourcePartName);
        if (string.IsNullOrWhiteSpace(name))
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

    private static bool IsManualLibraryDeleteSource(string? source) =>
        !string.IsNullOrWhiteSpace(source) && source.Contains("[MANUAL_LIBRARY_DELETE]", StringComparison.OrdinalIgnoreCase);

    private static bool IsValidDemandPartIps(string? value)
    {
        var text = UiText.Clean(value).Trim();
        return text.Length is >= 6 and <= 11 &&
            text.Any(ch => ch != '0') &&
            text.All(char.IsDigit);
    }

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

    private static DisplayOption<MeasurementUnit> GetConsumptionUnitOption(MeasurementUnit? unit) =>
        unit == MeasurementUnit.Meter
            ? UiText.ConsumptionUnitTypes.FirstOrDefault(UiText.IsMillimeterOption) ?? UiText.ConsumptionUnitTypes.First(x => x.Value == MeasurementUnit.Meter)
            : UiText.ConsumptionUnitTypes.FirstOrDefault(x => x.Value == unit) ?? UiText.ConsumptionUnitTypes[0];

    private static decimal? ToDisplayConsumptionQuantity(decimal? quantity, MeasurementUnit? unit) =>
        quantity is not null && unit == MeasurementUnit.Meter
            ? FromStoredMeterQuantity(quantity.Value)
            : quantity;

    private static decimal FromStoredMeterQuantity(decimal quantity) => quantity * 1000m;

    private static decimal ToStoredConsumptionQuantity(decimal quantity, DisplayOption<MeasurementUnit> unit) =>
        UiText.IsMillimeterOption(unit) ? quantity / 1000m : quantity;

    private static string DisplayConsumptionUnit(MeasurementUnit unit) =>
        unit == MeasurementUnit.Meter ? "мм" : DisplayUnit(unit);

    private static MeasurementUnit ParseConsumptionUnit(string? value, MeasurementUnit fallback)
    {
        var text = UiText.Clean(value).ToLowerInvariant();
        if (text.Contains("пог", StringComparison.Ordinal) || text.Contains("м", StringComparison.Ordinal))
        {
            return MeasurementUnit.Meter;
        }

        if (text.Contains("шт", StringComparison.Ordinal))
        {
            return MeasurementUnit.Piece;
        }

        return fallback;
    }

    private static string DisplayBlankType(BlankType type) => UiText.DisplayBlankType(type);
}

public sealed partial class NormalizationViewModel(
    BlankDemandPlannerDbContext dbContext,
    IExcelImportService excelImportService,
    IFileDialogService fileDialogService,
    IOneCNomenclatureService? oneCNomenclatureService = null) : ObservableObject
{
    private readonly IOneCNomenclatureService oneCService = oneCNomenclatureService ?? new EmptyOneCNomenclatureService();

    private static readonly BlankType[] MeterBasedBlankTypes =
    [
        BlankType.RoundBar,
        BlankType.SquareBar,
        BlankType.HexBar,
        BlankType.PipeRound,
        BlankType.PipeRectangular,
        BlankType.Angle
    ];

    public ObservableCollection<NsiBlankRow> Rows { get; } = [];
    public ObservableCollection<NsiUsageRow> UsageRows { get; } = [];
    public ObservableCollection<string> MaterialSuggestions { get; } = [];
    public ObservableCollection<string> MaterialGostSuggestions { get; } = [];
    public ObservableCollection<string> ProfileGostSuggestions { get; } = [];

    public ObservableCollection<DisplayOption<BlankType>> BlankTypes { get; } = [..UiReferenceData.BlankTypes()];
    public ObservableCollection<DisplayOption<BlankType?>> BlankTypeFilters { get; } = [..UiReferenceData.BlankTypeFilters(includeUnknown: true)];
    public ObservableCollection<DisplayOption<MeasurementUnit>> UnitTypes { get; } = [..UiReferenceData.UnitTypes()];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankTypeFilter = UiReferenceData.BlankTypeFilters(includeUnknown: true)[0];
    [ObservableProperty] private string sizeFilter = string.Empty;
    [ObservableProperty] private string materialFilter = string.Empty;
    [ObservableProperty] private bool duplicatesOnly;
    [ObservableProperty] private string statusText = string.Empty;
    [ObservableProperty] private NsiBlankRow? selectedRow;
    [ObservableProperty] private DisplayOption<BlankType> editBlankType = UiReferenceData.BlankTypes().First(x => x.Value == BlankType.Unknown);
    [ObservableProperty] private DisplayOption<MeasurementUnit> editUnit = UiReferenceData.UnitTypes()[0];
    [ObservableProperty] private string editOneCCode = string.Empty;
    [ObservableProperty] private string editSourceName = string.Empty;
    [ObservableProperty] private string editMaterial = string.Empty;
    [ObservableProperty] private string editSize = string.Empty;
    [ObservableProperty] private string editMaterialGost = string.Empty;
    [ObservableProperty] private string editProfileGost = string.Empty;
    [ObservableProperty] private bool isUsagePanelVisible = true;
    [ObservableProperty] private string usageStatusText = "Выберите заготовку";
    [ObservableProperty] private string summaryText = "Строк НСИ: 0";
    public string UsagePanelButtonText => IsUsagePanelVisible ? "Скрыть применяемость" : "Отобразить применяемость";

    private enum NsiSuggestionKind
    {
        Material,
        MaterialGost,
        ProfileGost
    }

    private bool oneCNamesRefreshAttempted;
    private bool suppressNsiEditorAutoDefaults;
    private string? lastAutoMaterialGost;
    private string? lastAutoProfileGost;

    partial void OnSearchChanged(string value) => _ = LoadAsync();
    partial void OnSelectedBlankTypeFilterChanged(DisplayOption<BlankType?> value) => _ = LoadAsync();
    partial void OnSizeFilterChanged(string value) => _ = LoadAsync();
    partial void OnMaterialFilterChanged(string value) => _ = LoadAsync();
    partial void OnDuplicatesOnlyChanged(bool value) => _ = LoadAsync();
    partial void OnEditMaterialChanged(string value)
    {
        _ = LoadMaterialSuggestionsAsync(value);
        _ = ApplyMaterialGostDefaultAsync(value);
    }
    partial void OnEditMaterialGostChanged(string value) => _ = LoadMaterialGostSuggestionsAsync(value);
    partial void OnEditProfileGostChanged(string value) => _ = LoadProfileGostSuggestionsAsync(value);
    partial void OnEditBlankTypeChanged(DisplayOption<BlankType> value) => _ = ApplyProfileGostDefaultAsync(value.Value);
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

    public void RefreshReferenceLists()
    {
        var editBlankTypeValue = EditBlankType.Value;
        var selectedBlankTypeFilterValue = SelectedBlankTypeFilter.Value;
        var editUnitValue = EditUnit.Value;

        UiReferenceData.ReplaceOptions(BlankTypes, UiReferenceData.BlankTypes());
        UiReferenceData.ReplaceOptions(BlankTypeFilters, UiReferenceData.BlankTypeFilters(includeUnknown: true));
        UiReferenceData.ReplaceOptions(UnitTypes, UiReferenceData.UnitTypes());

        EditBlankType = BlankTypes.FirstOrDefault(x => x.Value == editBlankTypeValue)
            ?? BlankTypes.FirstOrDefault(x => x.Value == BlankType.Unknown)
            ?? BlankTypes.First();
        SelectedBlankTypeFilter = BlankTypeFilters.FirstOrDefault(x => x.Value == selectedBlankTypeFilterValue) ?? BlankTypeFilters.First();
        EditUnit = UnitTypes.FirstOrDefault(x => x.Value == editUnitValue) ?? UnitTypes.First();
    }

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
            var refreshedNames = await RefreshImportedNamesFromOneCAsync(Path.GetFileName(file), CancellationToken.None);
            StatusText = $"Импорт НСИ завершен. Прочитано: {report.ReadRows}; добавлено: {report.AddedRows}; обновлено: {report.UpdatedRows}; актуализировано из 1С: {refreshedNames}; ошибок: {report.ErrorRows}.";
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
    private async Task AddFromOneCAsync()
    {
        var item = await ShowOneCNomenclatureLookupWindowAsync();
        if (item is null)
        {
            StatusText = "Добавление из 1С отменено.";
            return;
        }

        SelectedRow = null;
        EditOneCCode = NormalizeOneCEditorCode(FirstNotEmpty(item.Article, item.Code));
        EditSourceName = item.Name;
        EditUnit = UnitTypes.FirstOrDefault(x => x.Value == ParseUnit(item.Unit)) ?? UnitTypes[0];
        StatusText = $"Позиция 1С выбрана: {EditOneCCode}. Заполните вид, размер и ГОСТы при необходимости, затем сохраните НСИ.";
    }

    [RelayCommand]
    private async Task LoadMaterialSuggestionsAsync(string? value)
    {
        await LoadNsiEditorSuggestionsAsync(value, NsiSuggestionKind.Material, MaterialSuggestions);
    }

    private async Task LoadMaterialGostSuggestionsAsync(string? value) =>
        await LoadNsiEditorSuggestionsAsync(value, NsiSuggestionKind.MaterialGost, MaterialGostSuggestions);

    private async Task LoadProfileGostSuggestionsAsync(string? value) =>
        await LoadNsiEditorSuggestionsAsync(value, NsiSuggestionKind.ProfileGost, ProfileGostSuggestions);

    private async Task LoadNsiEditorSuggestionsAsync(string? value, NsiSuggestionKind kind, ObservableCollection<string> target)
    {
        var text = (value ?? string.Empty).Trim();
        var query = dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive);
        var values = kind switch
        {
            NsiSuggestionKind.Material => await query.Select(x => x.CanonicalBlank!.Material).ToListAsync(),
            NsiSuggestionKind.MaterialGost => await query.Select(x => x.CanonicalBlank!.MaterialGost).ToListAsync(),
            _ => await query.Select(x => x.CanonicalBlank!.ProfileGost).ToListAsync()
        };

        var suggestions = values
            .Select(UiText.Clean)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(x => string.IsNullOrWhiteSpace(text) || UiSearchText.Contains(x, text))
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .Take(40)
            .ToList();

        target.Clear();
        foreach (var suggestion in suggestions)
        {
            target.Add(suggestion);
        }
    }

    private async Task ApplyMaterialGostDefaultAsync(string? material)
    {
        if (suppressNsiEditorAutoDefaults)
        {
            return;
        }

        var text = UiText.Clean(material).Trim();
        if (string.IsNullOrWhiteSpace(text) || (!string.IsNullOrWhiteSpace(EditMaterialGost) && !string.Equals(EditMaterialGost, lastAutoMaterialGost, StringComparison.Ordinal)))
        {
            return;
        }

        var blanks = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive)
            .Select(x => new { x.CanonicalBlank!.Material, x.CanonicalBlank.MaterialGost })
            .ToListAsync();
        var materialGost = blanks
            .Where(x => UiSearchText.EqualsNormalized(x.Material, text))
            .Select(x => UiText.Clean(x.MaterialGost))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .GroupBy(x => x, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(x => x.Count())
            .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.Key)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(materialGost))
        {
            return;
        }

        lastAutoMaterialGost = materialGost;
        EditMaterialGost = materialGost;
    }

    private async Task ApplyProfileGostDefaultAsync(BlankType blankType)
    {
        if (suppressNsiEditorAutoDefaults || blankType is BlankType.Purchased or BlankType.Unknown)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(EditProfileGost) && !string.Equals(EditProfileGost, lastAutoProfileGost, StringComparison.Ordinal))
        {
            return;
        }

        var values = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive && x.CanonicalBlank.BlankType == blankType)
            .Select(x => x.CanonicalBlank!.ProfileGost)
            .ToListAsync();
        var profileGost = values
            .Select(UiText.Clean)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .GroupBy(x => x, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(x => x.Count())
            .ThenBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.Key)
            .FirstOrDefault();
        if (string.IsNullOrWhiteSpace(profileGost))
        {
            return;
        }

        lastAutoProfileGost = profileGost;
        EditProfileGost = profileGost;
    }

    [RelayCommand]
    private void EditNsiRow(object? parameter)
    {
        var row = parameter as NsiBlankRow ?? SelectedRow;
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
        lastAutoMaterialGost = null;
        lastAutoProfileGost = null;
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

        var code = NullIfWhiteSpace(NormalizeOneCEditorCode(EditOneCCode));
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
            await LoadAsync();
            SelectedRow = Rows.FirstOrDefault(x => x.AliasId == aliasId);
            StatusText = $"Деталь добавлена в список: {code}.";
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
    private async Task DeleteSelectedAsync() => await DeleteSelectedRowsAsync(SelectedRow);


    [RelayCommand]
    private async Task DeleteSelectedRowsAsync(object? parameter)
    {
        try
        {
            var rows = GetSelectedNsiRows(parameter).DistinctBy(x => x.AliasId).ToArray();
            if (rows.Length == 0)
            {
                StatusText = "Выберите одну или несколько строк НСИ для удаления.";
                return;
            }

            var aliasIds = rows.Select(x => x.AliasId).Distinct().ToArray();
            var canonicalIds = rows.Select(x => x.CanonicalBlankId).Distinct().ToArray();
            var usageCount = await dbContext.PartBlankMaps.AsNoTracking()
                .CountAsync(x => x.IsActive && canonicalIds.Contains(x.CanonicalBlankId));
            if (usageCount > 0)
            {
                var confirmation = MessageBox.Show(
                    $"По выбранной НСИ есть применяемость: {usageCount} активных связей с деталями.\n\nПосле удаления эти связи будут удалены, и детали останутся без этой заготовки в расчетах. Подтвердить удаление?",
                    "Удаление НСИ с применяемостью",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (confirmation != MessageBoxResult.Yes)
                {
                    StatusText = "Удаление НСИ отменено.";
                    return;
                }
            }

            var aliases = await dbContext.BlankAliases
                .Include(x => x.CanonicalBlank)
                .Where(x => aliasIds.Contains(x.Id))
                .ToListAsync();
            var allAliasesByCanonical = await dbContext.BlankAliases
                .Where(x => canonicalIds.Contains(x.CanonicalBlankId))
                .ToListAsync();
            var canonicalIdsToDelete = allAliasesByCanonical
                .GroupBy(x => x.CanonicalBlankId)
                .Where(x => x.All(a => aliasIds.Contains(a.Id)))
                .Select(x => x.Key)
                .ToArray();
            var undoSnapshot = await SnapshotNsiDeleteAsync(aliasIds, canonicalIdsToDelete);

            var stockLinks = await dbContext.StockItems
                .Where(x => x.BlankAliasId.HasValue && aliasIds.Contains(x.BlankAliasId.Value))
                .ToListAsync();
            foreach (var stock in stockLinks)
            {
                stock.BlankAliasId = null;
                stock.BlankAlias = null;
            }

            if (canonicalIdsToDelete.Length > 0)
            {
                var maps = await dbContext.PartBlankMaps
                    .Where(x => canonicalIdsToDelete.Contains(x.CanonicalBlankId))
                    .ToListAsync();
                dbContext.PartBlankMaps.RemoveRange(maps);

                var calculationItems = await dbContext.CalculationItems
                    .Where(x => x.CanonicalBlankId.HasValue && canonicalIdsToDelete.Contains(x.CanonicalBlankId.Value))
                    .ToListAsync();
                foreach (var item in calculationItems)
                {
                    item.CanonicalBlankId = null;
                    item.CanonicalBlank = null;
                }
            }

            foreach (var alias in aliases)
            {
                dbContext.BlankAliases.Remove(alias);
            }

            if (canonicalIdsToDelete.Length > 0)
            {
                var blanks = await dbContext.CanonicalBlanks
                    .Where(x => canonicalIdsToDelete.Contains(x.Id))
                    .ToListAsync();
                dbContext.CanonicalBlanks.RemoveRange(blanks);
            }

            await dbContext.SaveChangesAsync();
            StatusText = $"Удалено позиций НСИ: {aliases.Count}.";
            UndoCenter.Push("Отмена удаления НСИ", async () => await RestoreNsiDeleteAsync(undoSnapshot));
            await LoadAsync();
        }
        catch (Exception ex)
        {
            try
            {
                dbContext.Logs.Add(new AppLog
                {
                    Level = "Error",
                    Message = "Ошибка удаления НСИ",
                    Exception = ex.ToString()
                });
                await dbContext.SaveChangesAsync();
            }
            catch
            {
                // Keep the UI responsive even if logging fails.
            }

            StatusText = "Не удалось удалить НСИ. Подробности записаны в лог.";
            MessageBox.Show($"Не удалось удалить НСИ.\n{ex.GetBaseException().Message}", "НСИ", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task<NsiDeleteSnapshot> SnapshotNsiDeleteAsync(long[] aliasIds, long[] canonicalIdsToDelete)
    {
        var blanks = await dbContext.CanonicalBlanks.AsNoTracking()
            .Where(x => canonicalIdsToDelete.Contains(x.Id))
            .Select(x => new NsiCanonicalBlankSnapshot(
                x.Id,
                x.CanonicalName,
                x.CanonicalKey,
                x.BlankType,
                x.Material,
                x.MaterialGost,
                x.ProfileGost,
                x.DiameterMm,
                x.WidthMm,
                x.HeightMm,
                x.ThicknessMm,
                x.WallThicknessMm,
                x.LengthMm,
                x.BaseUnit,
                x.I012Status,
                x.I012Section,
                x.CreatedAt,
                x.UpdatedAt,
                x.IsActive))
            .ToListAsync();
        var aliases = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => aliasIds.Contains(x.Id))
            .Select(x => new NsiBlankAliasSnapshot(
                x.Id,
                x.CanonicalBlankId,
                x.OneCCode,
                x.SourceName,
                x.NormalizedSourceName,
                x.Source,
                x.ImportedAt,
                x.UpdatedAt,
                x.IsActive))
            .ToListAsync();
        var maps = await dbContext.PartBlankMaps.AsNoTracking()
            .Where(x => canonicalIdsToDelete.Contains(x.CanonicalBlankId))
            .Select(x => new NsiPartBlankMapSnapshot(
                x.Id,
                x.PartId,
                x.CanonicalBlankId,
                x.ConsumptionQuantity,
                x.ConsumptionUnit,
                x.LossPercent,
                x.BlankLeadTimeDays,
                x.Source,
                x.SourceFile,
                x.UpdatedAt,
                x.IsActive,
                x.IsPrimary))
            .ToListAsync();
        var stockLinks = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.BlankAliasId.HasValue && aliasIds.Contains(x.BlankAliasId.Value))
            .Select(x => new NsiStockLinkSnapshot(x.Id, x.BlankAliasId))
            .ToListAsync();
        var calculationLinks = await dbContext.CalculationItems.AsNoTracking()
            .Where(x => x.CanonicalBlankId.HasValue && canonicalIdsToDelete.Contains(x.CanonicalBlankId.Value))
            .Select(x => new NsiCalculationLinkSnapshot(x.Id, x.CanonicalBlankId))
            .ToListAsync();

        return new NsiDeleteSnapshot(blanks, aliases, maps, stockLinks, calculationLinks);
    }

    private async Task RestoreNsiDeleteAsync(NsiDeleteSnapshot snapshot)
    {
        dbContext.ChangeTracker.Clear();
        foreach (var saved in snapshot.Blanks)
        {
            var blank = await dbContext.CanonicalBlanks.FirstOrDefaultAsync(x => x.Id == saved.Id);
            if (blank is null)
            {
                dbContext.CanonicalBlanks.Add(saved.ToEntity());
            }
            else
            {
                saved.ApplyTo(blank);
            }
        }

        await dbContext.SaveChangesAsync();

        foreach (var saved in snapshot.Aliases)
        {
            var alias = await dbContext.BlankAliases.FirstOrDefaultAsync(x => x.Id == saved.Id);
            if (alias is null)
            {
                dbContext.BlankAliases.Add(saved.ToEntity());
            }
            else
            {
                saved.ApplyTo(alias);
            }
        }

        await dbContext.SaveChangesAsync();

        foreach (var saved in snapshot.Maps)
        {
            var map = await dbContext.PartBlankMaps.FirstOrDefaultAsync(x => x.Id == saved.Id);
            if (map is null)
            {
                dbContext.PartBlankMaps.Add(saved.ToEntity());
            }
            else
            {
                saved.ApplyTo(map);
            }
        }

        foreach (var saved in snapshot.StockLinks)
        {
            var stock = await dbContext.StockItems.FirstOrDefaultAsync(x => x.Id == saved.Id);
            if (stock is not null)
            {
                stock.BlankAliasId = saved.BlankAliasId;
            }
        }

        foreach (var saved in snapshot.CalculationLinks)
        {
            var item = await dbContext.CalculationItems.FirstOrDefaultAsync(x => x.Id == saved.Id);
            if (item is not null)
            {
                item.CanonicalBlankId = saved.CanonicalBlankId;
            }
        }

        await dbContext.SaveChangesAsync();
        await LoadAsync();
        StatusText = $"Восстановлено позиций НСИ: {snapshot.Aliases.Count}.";
    }

    private async Task CreateManualNsiBlankAsync()
    {
        var code = NullIfWhiteSpace(NormalizeOneCEditorCode(EditOneCCode));
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
            await dbContext.SaveChangesAsync();
            await LoadAsync();
            SelectedRow = Rows.FirstOrDefault(x => x.AliasId == existingAlias.Id);
            StatusText = $"Деталь добавлена в список: {code}.";
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
        await LoadAsync();
        SelectedRow = Rows.FirstOrDefault(x => x.AliasId == alias.Id);
        StatusText = $"Деталь добавлена в список: {code}.";
    }

    private async Task<int> RefreshImportedNamesFromOneCAsync(string sourceFileName, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(sourceFileName))
        {
            return 0;
        }

        var aliases = await dbContext.BlankAliases
            .Include(x => x.CanonicalBlank)
            .Where(x => x.Source == sourceFileName && x.OneCCode != "")
            .OrderByDescending(x => x.UpdatedAt)
            .Take(500)
            .ToListAsync(cancellationToken);
        return await RefreshAliasNamesFromOneCAsync(aliases, cancellationToken);
    }

    private async Task<int> RefreshNsiNamesFromOneCOnOpenAsync(CancellationToken cancellationToken)
    {
        if (oneCNamesRefreshAttempted)
        {
            return 0;
        }

        return await RefreshNamesFromOneCAsync(cancellationToken);
    }

    public async Task<int> RefreshNamesFromOneCAsync(CancellationToken cancellationToken)
    {
        oneCNamesRefreshAttempted = true;
        var aliases = await dbContext.BlankAliases
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive && x.OneCCode != "")
            .OrderByDescending(x => x.UpdatedAt)
            .Take(500)
            .ToListAsync(cancellationToken);
        return await RefreshAliasNamesFromOneCAsync(aliases, cancellationToken);
    }

    private async Task<int> RefreshAliasNamesFromOneCAsync(IReadOnlyList<BlankAlias> aliases, CancellationToken cancellationToken)
    {
        if (aliases.Count == 0)
        {
            return 0;
        }

        IReadOnlyList<OneCNomenclatureItem> items;
        try
        {
            items = await oneCService.ResolveByCodesAsync(aliases.Select(x => x.OneCCode), cancellationToken);
        }
        catch
        {
            return 0;
        }

        var itemsByKey = items
            .SelectMany(item => BuildOneCItemKeys(item).Select(key => (Key: key, Item: item)))
            .Where(x => !string.IsNullOrWhiteSpace(x.Key))
            .GroupBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First().Item, StringComparer.OrdinalIgnoreCase);
        var updated = 0;
        foreach (var alias in aliases)
        {
            var aliasKeys = BuildAliasLookupKeys(alias.OneCCode);
            var item = aliasKeys.Select(key => itemsByKey.GetValueOrDefault(key)).FirstOrDefault(x => x is not null);
            if (item is null || string.IsNullOrWhiteSpace(item.Name) || string.Equals(alias.SourceName, item.Name, StringComparison.Ordinal))
            {
                continue;
            }

            alias.SourceName = item.Name.Trim();
            alias.NormalizedSourceName = NormalizeKey(alias.SourceName);
            alias.UpdatedAt = DateTime.UtcNow;
            if (alias.CanonicalBlank is not null)
            {
                alias.CanonicalBlank.CanonicalName = alias.SourceName;
                alias.CanonicalBlank.BaseUnit = ParseUnit(item.Unit);
                alias.CanonicalBlank.UpdatedAt = DateTime.UtcNow;
            }

            updated++;
        }

        if (updated > 0)
        {
            await dbContext.SaveChangesAsync(cancellationToken);
        }

        return updated;
    }

    private static IEnumerable<string> BuildAliasLookupKeys(string? code)
    {
        var text = (code ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            yield break;
        }

        yield return text;
        yield return StockCodeNormalizer.NormalizeForComparison(text);
    }

    private static IEnumerable<string> BuildOneCItemKeys(OneCNomenclatureItem item)
    {
        foreach (var value in new[] { item.Code, item.Article })
        {
            var text = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            yield return text;
            yield return StockCodeNormalizer.NormalizeForComparison(text);
        }
    }

    private Task<OneCNomenclatureItem?> ShowOneCNomenclatureLookupWindowAsync()
    {
        var items = new ObservableCollection<OneCNomenclatureItem>();
        OneCNomenclatureItem? selected = null;
        var searchBox = new System.Windows.Controls.TextBox { Width = 360, MinWidth = 320, Text = EditOneCCode };
        var status = new TextBlock { Text = "Введите УТ-код, IPS или часть наименования.", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0) };
        var grid = new DataGrid
        {
            ItemsSource = items,
            AutoGenerateColumns = false,
            IsReadOnly = true,
            SelectionMode = DataGridSelectionMode.Single,
            Margin = new Thickness(0, 12, 0, 0),
            MinHeight = 260
        };
        grid.Columns.Add(new DataGridTextColumn { Header = "Код 1С", Binding = new System.Windows.Data.Binding(nameof(OneCNomenclatureItem.Code)), Width = 120 });
        grid.Columns.Add(new DataGridTextColumn { Header = "УТ / артикул", Binding = new System.Windows.Data.Binding(nameof(OneCNomenclatureItem.Article)), Width = 120 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new System.Windows.Data.Binding(nameof(OneCNomenclatureItem.Name)), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Ед.", Binding = new System.Windows.Data.Binding(nameof(OneCNomenclatureItem.Unit)), Width = 90 });

        var searchButton = new System.Windows.Controls.Button { Content = "Найти", MinWidth = 90, Margin = new Thickness(8, 0, 0, 0) };
        var selectButton = new System.Windows.Controls.Button { Content = "Добавить в НСИ", MinWidth = 150, IsDefault = true };
        var cancelButton = new System.Windows.Controls.Button { Content = "Отмена", MinWidth = 90, IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
        var top = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
        top.Children.Add(searchBox);
        top.Children.Add(searchButton);
        var bottom = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        bottom.Children.Add(selectButton);
        bottom.Children.Add(cancelButton);
        var root = new DockPanel { Margin = new Thickness(16) };
        DockPanel.SetDock(top, Dock.Top);
        DockPanel.SetDock(status, Dock.Top);
        DockPanel.SetDock(bottom, Dock.Bottom);
        root.Children.Add(top);
        root.Children.Add(status);
        root.Children.Add(bottom);
        root.Children.Add(grid);
        var window = new Window
        {
            Title = "Добавить НСИ из 1С",
            Content = root,
            Width = 860,
            Height = 520,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(x => x.IsActive)
        };

        async Task SearchAsync()
        {
            var query = searchBox.Text.Trim();
            selected = null;
            grid.SelectedItem = null;
            items.Clear();
            if (string.IsNullOrWhiteSpace(query))
            {
                status.Text = "Введите УТ-код, IPS или часть наименования.";
                return;
            }

            status.Text = "Поиск в 1С...";
            try
            {
                var oneCQuery = BuildOneCSearchQuery(query);
                var result = await oneCService.SearchAsync(oneCQuery, 80, CancellationToken.None);
                if (result.Count == 0 && !string.Equals(oneCQuery, query, StringComparison.Ordinal))
                {
                    result = await oneCService.SearchAsync(query, 80, CancellationToken.None);
                }

                foreach (var item in result)
                {
                    items.Add(item);
                }

                status.Text = items.Count == 0 ? "Позиции в 1С не найдены." : $"Найдено позиций: {items.Count}. Выберите нужную строку.";
            }
            catch (Exception ex)
            {
                status.Text = $"Не удалось выполнить поиск в 1С: {ex.GetBaseException().Message}";
            }
        }

        searchButton.Click += async (_, _) => await SearchAsync();
        searchBox.KeyDown += async (_, args) =>
        {
            if (args.Key == System.Windows.Input.Key.Enter)
            {
                args.Handled = true;
                await SearchAsync();
            }
        };
        grid.MouseDoubleClick += (_, _) =>
        {
            if (grid.SelectedItem is OneCNomenclatureItem item)
            {
                selected = item;
                window.DialogResult = true;
            }
        };
        selectButton.Click += (_, _) =>
        {
            if (grid.SelectedItem is OneCNomenclatureItem item)
            {
                selected = item;
                window.DialogResult = true;
            }
            else
            {
                status.Text = "Выберите позицию из списка.";
            }
        };

        _ = SearchAsync();
        window.ShowDialog();
        return Task.FromResult(selected);
    }

    private static string NormalizeOneCEditorCode(string value)
    {
        var code = UiText.Clean(value).Trim();
        if (code.Length == 11 && code.All(char.IsDigit))
        {
            var ips = code.TrimStart('0');
            return ips.Length == 0 ? "0" : ips;
        }

        return code;
    }

    private static string BuildOneCSearchQuery(string value)
    {
        var query = UiText.Clean(value).Trim();
        if (LooksLikeCodeSearch(query) && !query.Any(char.IsWhiteSpace))
        {
            return UiSearchText.Normalize(query);
        }

        return NormalizeOneCTextSearchQuery(query);
    }

    private static string NormalizeOneCTextSearchQuery(string value)
    {
        var text = Regex.Replace(value, @"(?<![\p{L}\p{N}])[Дд]\s*(?=\d)", "D", RegexOptions.CultureInvariant);
        text = Regex.Replace(text, @"(?<=\d)\s*[ХхXx]\s*(?=\d)", "x", RegexOptions.CultureInvariant);
        return text;
    }

    private static bool LooksLikeCodeSearch(string value) =>
        value.Any(char.IsDigit) ||
        value.Any(ch => ch is '-' or '.' or '_' or '/' or '\\') ||
        value.Any(ch => ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z');

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

    private static Task EnsureNsiRowsAvailableAsync() => Task.CompletedTask;
    [RelayCommand]
    public async Task LoadAsync()
    {
        await EnsureNsiRowsAvailableAsync();
        var refreshedFromOneC = await RefreshNsiNamesFromOneCOnOpenAsync(CancellationToken.None);

        var query = dbContext.BlankAliases.AsNoTracking()
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive);

        if (SelectedBlankTypeFilter.Value is not null)
        {
            var blankType = SelectedBlankTypeFilter.Value.Value;
            query = query.Where(x => x.CanonicalBlank != null && x.CanonicalBlank.BlankType == blankType);
        }

        string? textSizeFilter = null;
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
                textSizeFilter = sizeText;
            }
        }

        var aliases = await query
            .OrderBy(x => x.OneCCode)
            .ToListAsync();
        var duplicateKeys = aliases
            .Select(x => NsiDuplicateKey.Build(x.CanonicalBlank, x))
            .Where(x => x.Length > 0)
            .GroupBy(x => x, StringComparer.Ordinal)
            .Where(x => x.Count() > 1)
            .Select(x => x.Key)
            .ToHashSet(StringComparer.Ordinal);
        var duplicateGroupCount = duplicateKeys.Count;
        if (DuplicatesOnly)
        {
            aliases = aliases
                .Where(x => duplicateKeys.Contains(NsiDuplicateKey.Build(x.CanonicalBlank, x)))
                .OrderBy(x => NsiDuplicateKey.Build(x.CanonicalBlank, x), StringComparer.Ordinal)
                .ThenBy(x => x.OneCCode, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        var searchValue = Search.Trim();
        var materialFilter = MaterialFilter.Trim();
        if (!string.IsNullOrWhiteSpace(searchValue))
        {
            aliases = aliases.Where(x => NsiAliasMatchesSearch(x, searchValue)).ToList();
        }

        if (!string.IsNullOrWhiteSpace(materialFilter))
        {
            aliases = aliases
                .Where(x =>
                    UiSearchText.Contains(x.CanonicalBlank?.Material, materialFilter) ||
                    UiSearchText.Contains(x.SourceName, materialFilter) ||
                    UiSearchText.Contains(x.NormalizedSourceName, materialFilter))
                .ToList();
        }

        if (!string.IsNullOrWhiteSpace(textSizeFilter))
        {
            aliases = aliases
                .Where(x =>
                    UiSearchText.Contains(x.SourceName, textSizeFilter) ||
                    UiSearchText.Contains(x.NormalizedSourceName, textSizeFilter) ||
                    UiSearchText.Contains(x.CanonicalBlank?.CanonicalName, textSizeFilter))
                .ToList();
        }

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

        var pricesByCode = await LoadOneCPricesByCodeAsync(aliases);

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
            pricesByCode.TryGetValue(stockCodeKey, out var price);
            Rows.Add(new NsiBlankRow(
                alias.Id,
                alias.CanonicalBlankId,
                alias.OneCCode,
                UiText.Clean(alias.SourceName),
                DisplayUnit(unit),
                blank is null ? "Не распознано" : DisplayBlankType(blank.BlankType),
                FormatSize(blank),
                duplicateKeys.Contains(NsiDuplicateKey.Build(blank, alias)) ? BuildDuplicateDisplayKey(blank, alias) : string.Empty,
                UiText.Clean(blank?.Material),
                UiText.Clean(blank?.MaterialGost),
                UiText.Clean(blank?.ProfileGost),
                hasCmoStock ? FormatStockQuantity(cmoStock.Quantity, cmoStock.Unit) : "0",
                hasWarehouseStock ? FormatStockQuantity(warehouseStock.Quantity, warehouseStock.Unit) : "0",
                hasCmoStock ? DisplayStockUnit(cmoStock.Unit) : hasWarehouseStock ? DisplayStockUnit(warehouseStock.Unit) : string.Empty,
                FormatPrice(price),
                UiText.Clean(alias.SourceName),
                UiText.Clean(alias.Source),
                alias.UpdatedAt.ToLocalTime().ToString("g")));
        }

        SummaryText = DuplicatesOnly
            ? $"Строк НСИ: {Rows.Count}; групп дублей: {duplicateGroupCount}"
            : $"Строк НСИ: {Rows.Count}";
        StatusText = refreshedFromOneC > 0
            ? $"Показано позиций НСИ: {Rows.Count}; наименований обновлено из 1С: {refreshedFromOneC}."
            : DuplicatesOnly
                ? $"Показано дублей НСИ: {Rows.Count}; групп дублей: {duplicateGroupCount}. Строки сгруппированы по сходимости."
                : $"Показано позиций НСИ: {Rows.Count}";
        if (SelectedRow is not null && Rows.All(x => x.AliasId != SelectedRow.AliasId))
        {
            SelectedRow = null;
        }

        await LoadMaterialSuggestionsAsync(EditMaterial);
        await LoadMaterialGostSuggestionsAsync(EditMaterialGost);
        await LoadProfileGostSuggestionsAsync(EditProfileGost);
    }

    public async Task<int> RefreshPricesFromOneCAsync(CancellationToken cancellationToken)
    {
        var codes = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive && x.OneCCode != string.Empty)
            .Select(x => x.OneCCode)
            .Distinct()
            .ToListAsync(cancellationToken);
        var requestedKeys = codes
            .Select(StockCodeNormalizer.NormalizeForComparison)
            .Where(x => x.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (requestedKeys.Count == 0)
        {
            return 0;
        }

        var pricesByKey = new Dictionary<string, OneCNomenclaturePrice>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var chunk in codes.Chunk(500))
            {
                var prices = await oneCService.ResolvePricesByCodesAsync(chunk, cancellationToken);
                foreach (var price in prices.Where(x => x.Price > 0))
                {
                    foreach (var key in BuildOneCPriceKeys(price).Select(StockCodeNormalizer.NormalizeForComparison).Where(x => x.Length > 0))
                    {
                        if (!pricesByKey.TryGetValue(key, out var current) ||
                            IsPreferredPriceType(price.PriceType) && !IsPreferredPriceType(current.PriceType) ||
                            IsPreferredPriceType(price.PriceType) == IsPreferredPriceType(current.PriceType) && price.Price > current.Price)
                        {
                            pricesByKey[key] = price;
                        }
                    }
                }
            }
        }
        catch
        {
            return 0;
        }

        var now = DateTime.UtcNow;
        var relevantKeys = requestedKeys.Concat(pricesByKey.Keys).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var existing = await dbContext.OneCPriceItems
            .Where(x => relevantKeys.Contains(x.LookupKey))
            .ToDictionaryAsync(x => x.LookupKey, StringComparer.OrdinalIgnoreCase, cancellationToken);

        var changed = 0;
        foreach (var key in requestedKeys)
        {
            if (!pricesByKey.ContainsKey(key) && existing.TryGetValue(key, out var stale))
            {
                dbContext.OneCPriceItems.Remove(stale);
                changed++;
            }
        }

        foreach (var (key, price) in pricesByKey)
        {
            if (!existing.TryGetValue(key, out var item))
            {
                dbContext.OneCPriceItems.Add(new OneCPriceItem
                {
                    LookupKey = key,
                    Code = price.Code,
                    Article = price.Article,
                    Price = price.Price,
                    Currency = price.Currency,
                    PriceType = price.PriceType,
                    SyncedAt = now
                });
                changed++;
                continue;
            }

            if (item.Code != price.Code ||
                item.Article != price.Article ||
                item.Price != price.Price ||
                item.Currency != price.Currency ||
                item.PriceType != price.PriceType)
            {
                changed++;
            }

            item.Code = price.Code;
            item.Article = price.Article;
            item.Price = price.Price;
            item.Currency = price.Currency;
            item.PriceType = price.PriceType;
            item.SyncedAt = now;
        }

        await dbContext.SaveChangesAsync(cancellationToken);
        return changed;
    }

    private async Task<Dictionary<string, OneCPriceItem>> LoadOneCPricesByCodeAsync(IReadOnlyList<BlankAlias> aliases)
    {
        var codes = aliases
            .Select(x => x.OneCCode)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(StockCodeNormalizer.NormalizeForComparison)
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (codes.Length == 0)
        {
            return new Dictionary<string, OneCPriceItem>(StringComparer.OrdinalIgnoreCase);
        }

        var items = await dbContext.OneCPriceItems.AsNoTracking()
            .Where(x => codes.Contains(x.LookupKey))
            .ToListAsync();
        return items
            .GroupBy(x => x.LookupKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(p => p.SyncedAt).First(), StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> BuildOneCPriceKeys(OneCNomenclaturePrice price)
    {
        if (!string.IsNullOrWhiteSpace(price.Code))
        {
            yield return price.Code;
        }

        if (!string.IsNullOrWhiteSpace(price.Article))
        {
            yield return price.Article;
        }
    }

    private static string FormatPrice(OneCPriceItem? price)
    {
        if (price is null || price.Price <= 0)
        {
            return string.Empty;
        }

        var currency = FormatCurrency(price.Currency);
        var amount = FormatDecimal(price.Price);
        return string.IsNullOrWhiteSpace(currency) ? amount : $"{amount} {currency}";
    }

    private static bool IsPreferredPriceType(string? value) =>
        UiText.Clean(value).Contains("Стоимость", StringComparison.OrdinalIgnoreCase);

    private static string FormatCurrency(string? currency)
    {
        var value = UiText.Clean(currency).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var upper = value.ToUpperInvariant();
        return upper is "RUB" or "RUR" or "643" || upper.Contains("РУБ", StringComparison.Ordinal)
            ? "₽"
            : value;
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

    private static string BuildDuplicateDisplayKey(CanonicalBlank? blank, BlankAlias alias)
    {
        if (blank is null)
        {
            return string.Empty;
        }

        var size = FormatSize(blank);
        var physicalPart = string.IsNullOrWhiteSpace(size)
            ? UiText.Clean(alias.SourceName)
            : size;
        return string.Join(" | ", new[]
            {
                DisplayBlankType(blank.BlankType),
                DisplayUnit(MeterBasedBlankTypes.Contains(blank.BlankType) ? MeasurementUnit.Meter : blank.BaseUnit),
                UiText.Clean(blank.Material),
                UiText.Clean(blank.MaterialGost),
                UiText.Clean(blank.ProfileGost),
                physicalPart
            }
            .Where(x => !string.IsNullOrWhiteSpace(x)));
    }

    private void LoadSelectedEditor(NsiBlankRow? row)
    {
        suppressNsiEditorAutoDefaults = true;
        if (row is null)
        {
            EditBlankType = BlankTypes.First(x => x.Value == BlankType.Unknown);
            EditMaterial = string.Empty;
            EditSize = string.Empty;
            EditMaterialGost = string.Empty;
            EditProfileGost = string.Empty;
            lastAutoMaterialGost = null;
            lastAutoProfileGost = null;
            suppressNsiEditorAutoDefaults = false;
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
        lastAutoMaterialGost = null;
        lastAutoProfileGost = null;
        suppressNsiEditorAutoDefaults = false;
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
    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;

    private static MeasurementUnit ParseUnit(string? value)
    {
        var text = UiText.Clean(value).ToUpperInvariant();
        if ((text.Contains("ПОГ", StringComparison.Ordinal) || text.Contains("М", StringComparison.Ordinal) || text.Contains("M", StringComparison.Ordinal)) &&
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

    private static bool NsiAliasMatchesSearch(BlankAlias alias, string searchValue) =>
        UiSearchText.Contains(alias.OneCCode, searchValue) ||
        UiSearchText.Contains(alias.SourceName, searchValue) ||
        UiSearchText.Contains(alias.NormalizedSourceName, searchValue) ||
        UiSearchText.Contains(alias.CanonicalBlank?.CanonicalName, searchValue) ||
        UiSearchText.Contains(alias.CanonicalBlank?.Material, searchValue) ||
        UiSearchText.Contains(alias.CanonicalBlank?.MaterialGost, searchValue) ||
        UiSearchText.Contains(alias.CanonicalBlank?.ProfileGost, searchValue);

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

    private static string FormatStockQuantity(decimal quantity, MeasurementUnit unit) =>
        FormatDecimal(unit == MeasurementUnit.Meter ? quantity * 1000m : quantity);

    private static string DisplayStockUnit(MeasurementUnit unit) =>
        unit == MeasurementUnit.Meter ? "мм" : DisplayUnit(unit);

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
        BlankType.WeldingElement => "Сварное изделие",
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
            var searchedItems = await query
                .OrderBy(x => x.OneCCode)
                .Take(10000)
                .ToListAsync();
            searchedItems = searchedItems
                .Where(x =>
                    UiSearchText.Contains(x.OneCCode, searchValue) ||
                    UiSearchText.Contains(x.SourceName, searchValue) ||
                    UiSearchText.Contains(x.Warehouse, searchValue))
                .Take(2000)
                .ToList();

            AddRows(searchedItems, snapshot.SourceFile);
            return;
        }

        var items = await query
            .OrderBy(x => x.OneCCode)
            .Take(2000)
            .ToListAsync();

        AddRows(items, snapshot.SourceFile);
    }

    private void AddRows(IEnumerable<StockItem> items, string? sourceFile)
    {
        foreach (var item in items)
        {
            Rows.Add(new StockRow(
                item.OneCCode,
                UiText.Clean(item.SourceName),
                FormatDecimal(item.Quantity),
                UiText.DisplayUnit(item.Unit),
                UiText.Clean(item.Warehouse)));
        }

        StatusText = $"Показано остатков: {Rows.Count}; источник: {Path.GetFileName(sourceFile ?? string.Empty)}";
    }

    private static string FormatDecimal(decimal quantity) => quantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
}

public sealed partial class MskViewModel(
    BlankDemandPlannerDbContext dbContext,
    ILogger logger,
    IIpsDrawingService? ipsDrawingService = null,
    IOneCProductionLaunchService? oneCProductionLaunchService = null,
    IOneCGoodsTransferService? oneCGoodsTransferService = null,
    bool autoOpenDrawings = false) : ObservableObject
{
    private const decimal ProductionLaunchMeterCutWidth = 0.005m;

    private const string DefaultMskFolder = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК";

    private readonly List<MskLibraryRow> allRows = [];
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();
    private readonly HashSet<string> attemptedDrawingIps = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> demandIps = new(StringComparer.OrdinalIgnoreCase);
    private int detailMetricsVersion;

    public ObservableCollection<MskLibraryRow> Rows { get; } = [];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private bool demandOnly;
    [ObservableProperty] private string statusText = "МСК не загружены.";
    [ObservableProperty] private MskLibraryRow? selectedRow;
    [ObservableProperty] private string detailText = "Выберите деталь для просмотра МСК.";
    [ObservableProperty] private string detailMetricsText = "Выберите деталь для расчета потребности и запуска.";
    [ObservableProperty] private Uri? drawingViewerSource;
    [ObservableProperty] private string drawingStatusText = "PDF-чертеж не выбран.";
    [ObservableProperty] private bool isBusy;
    [ObservableProperty] private bool isDrawingBusy;
    [ObservableProperty] private string productionLaunchStatusText = "Запуск в производство не выполнялся.";

    partial void OnSearchChanged(string value) => ApplyFilter();
    partial void OnDemandOnlyChanged(bool value) => ApplyFilter();

    partial void OnSelectedRowChanged(MskLibraryRow? value)
    {
        DetailText = value is null
            ? "Выберите деталь для просмотра МСК."
            : BuildDetailText(value);
        DetailMetricsText = value is null
            ? "Выберите деталь для расчета потребности и запуска."
            : "Расчет показателей...";
        _ = LoadDetailMetricsAsync(value, Interlocked.Increment(ref detailMetricsVersion));
        if (value is null)
        {
            DrawingViewerSource = null;
            DrawingStatusText = "PDF-чертеж не выбран.";
        }
        else if (!autoOpenDrawings && DrawingViewerSource is null)
        {
            DrawingStatusText = "PDF-чертеж не открыт. Кликните строку МСК, чтобы открыть чертеж.";
        }

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
                    if (IsValidMskLibraryIps(record.Ips))
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

            var removedInvalidLibraryParts = await RemoveInvalidMskLibraryPartsAsync();
            var parts = await dbContext.Parts.ToListAsync();
            var mskIps = latestByIps.Select(x => x.Ips).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var part in parts)
            {
                part.HasMsk = IsValidMskLibraryIps(part.Ips) && mskIps.Contains(part.Ips);
                part.UpdatedAt = DateTime.UtcNow;
            }
            var addedLibraryParts = await AddMissingLibraryPartsFromMskAsync(latestByIps);

            await dbContext.SaveChangesAsync();
            await LoadFromDatabaseAsync();
            StatusText = $"МСК обновлены из X: файлов {files.Count}, записей {latestByIps.Count}, ошибок чтения {errors}; добавлено в библиотеку: {addedLibraryParts}; убрано неверных: {removedInvalidLibraryParts}.";
            if (addedLibraryParts > 0)
            {
                MessageBox.Show($"При обновлении МСК добавлено новых деталей в библиотеку: {addedLibraryParts}.", "Планирование", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        finally
        {
            IsBusy = false;
        }
    }

    public async Task<int> AddMissingLibraryPartsFromMskAsync(IReadOnlyList<MskRecord>? records = null, IReadOnlyDictionary<string, MskCsvDetail>? details = null)
    {
        records ??= await dbContext.MskRecords.AsNoTracking().ToListAsync();
        var latestByIps = records
            .Where(x => IsValidMskLibraryIps(x.Ips))
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .Select(x => x.OrderByDescending(r => r.ImportedAt).First())
            .ToList();
        if (latestByIps.Count == 0)
        {
            return 0;
        }

        var existingParts = await dbContext.Parts
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ToListAsync();
        var partsByIps = existingParts
            .GroupBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(p => p.UpdatedAt).First(), StringComparer.OrdinalIgnoreCase);
        var mskDetails = details ?? LoadMskDetailsFromReport();
        var added = 0;
        foreach (var record in latestByIps)
        {
            var detail = ResolveMskDetail(record, mskDetails);
            if (partsByIps.TryGetValue(record.Ips, out var existingPart))
            {
                existingPart.HasMsk = true;
                if (IsArchivedLibrarySource(existingPart.Source))
                {
                    existingPart.Source = "МСК";
                    existingPart.Designation = FirstNotEmpty(record.Designation, existingPart.Designation);
                    existingPart.Name = FirstNotEmpty(record.Name, existingPart.Name, record.Ips);
                    existingPart.UpdatedAt = DateTime.UtcNow;
                    added++;
                }

                if (!existingPart.BlankMaps.Any(x => x.IsActive))
                {
                    await TryAttachMskBlankAsync(existingPart, detail);
                }

                continue;
            }

            var (designation, name) = SplitDesignationAndName(FirstNotEmpty(record.Name, record.Ips));
            var part = new Part
            {
                Ips = record.Ips.Trim(),
                Designation = FirstNotEmpty(record.Designation, designation),
                Name = FirstNotEmpty(name, record.Name, record.Ips),
                HasMsk = true,
                Source = "МСК"
            };
            dbContext.Parts.Add(part);
            partsByIps[part.Ips] = part;
            await TryAttachMskBlankAsync(part, detail);
            added++;
        }

        return added;
    }

    public async Task<int> RemoveInvalidMskLibraryPartsAsync()
    {
        var candidates = await dbContext.Parts
            .Include(x => x.BlankMaps)
            .Where(x => x.Source == "МСК")
            .ToListAsync();
        var invalidParts = candidates
            .Where(x => !IsValidMskLibraryIps(x.Ips))
            .ToList();
        if (invalidParts.Count == 0)
        {
            return 0;
        }

        var invalidIds = invalidParts.Select(x => x.Id).ToHashSet();
        var demandItems = await dbContext.DemandItems
            .Where(x => x.PartId != null && invalidIds.Contains(x.PartId.Value))
            .ToListAsync();
        foreach (var demandItem in demandItems)
        {
            demandItem.PartId = null;
        }

        dbContext.PartBlankMaps.RemoveRange(invalidParts.SelectMany(x => x.BlankMaps));
        dbContext.Parts.RemoveRange(invalidParts);
        await dbContext.SaveChangesAsync();
        return invalidParts.Count;
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
    private async Task LaunchProductionAsync()
    {
        var row = SelectedRow ?? ResolveRowFromSearch();
        if (row is null)
        {
            StatusText = "Выберите деталь МСК для запуска в производство.";
            return;
        }

        SelectedRow = row;
        if (oneCProductionLaunchService is null)
        {
            MessageBox.Show("Сервис запуска в производство 1С не подключен.", "Планирование", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        ProductionLaunchPreview preview;
        try
        {
            preview = await BuildProductionLaunchPreviewAsync(row);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not build production launch preview for {Ips}", row.Ips);
            MessageBox.Show($"Не удалось подготовить запуск в производство.\n{ex.GetBaseException().Message}", "Запуск в производство", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (preview.MaxQuantity <= 0)
        {
            MessageBox.Show($"Материала для запуска детали недостаточно.\nДоступно к запуску: 0\n{preview.MaterialLine}", "Запуск в производство", MessageBoxButton.OK, MessageBoxImage.Information);
            ProductionLaunchStatusText = "Запуск невозможен: материала недостаточно.";
            return;
        }

        var launch = ShowProductionLaunchDialog(preview);
        if (launch is null)
        {
            return;
        }

        var requests = BuildProductionLaunchRequests(preview, launch.Quantity, launch.Comment, launch.Piecewise);
        try
        {
            IsBusy = true;
            ProductionLaunchStatusText = launch.Piecewise
                ? $"Создание комплектаций в 1С: {requests.Count} шт..."
                : "Создание комплектации в 1С...";
            var results = new List<OneCProductionLaunchResult>();
            foreach (var request in requests)
            {
                results.Add(await oneCProductionLaunchService.CreateAssemblyAsync(request, CancellationToken.None));
            }

            var numbers = string.Join(", ", results.Select(x => x.Number).Where(x => !string.IsNullOrWhiteSpace(x)));
            ProductionLaunchStatusText = launch.Piecewise
                ? $"Создано комплектаций 1С: {results.Count}; номера: {numbers}."
                : $"Создана комплектация 1С {results[0].Number}; проведено: {(results[0].Posted ? "да" : "нет")}.";
            StatusText = ProductionLaunchStatusText;
            var message = launch.Piecewise
                ? $"Комплектации 1С созданы.\nДокументов: {results.Count}\nНомера: {numbers}\nКоличество деталей: {FormatDecimal(requests.Sum(x => x.Quantity))}\nДокументы не проведены."
                : $"Комплектация 1С создана.\nНомер: {results[0].Number}\nКоличество деталей: {FormatDecimal(requests[0].Quantity)}\nДокумент не проведен.\n\nКомментарий:\n{results[0].Comment}";
            MessageBox.Show(message, "Запуск в производство", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Could not create 1C production launch for {Ips}", row.Ips);
            ProductionLaunchStatusText = $"Ошибка запуска: {ex.GetBaseException().Message}";
            MessageBox.Show($"Не удалось создать комплектацию 1С.\n{ex.GetBaseException().Message}", "Запуск в производство", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task ExternalServicesAsync()
    {
        if (oneCGoodsTransferService is null)
        {
            MessageBox.Show("Сервис перемещения 1С не подключен.", "Услуги на стороне", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        IReadOnlyList<ExternalServiceWipRow> rows;
        try
        {
            rows = await LoadExternalServiceWipRowsAsync();
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not load WIP rows for external services");
            MessageBox.Show($"Не удалось загрузить детали из НЗП.\n{ex.GetBaseException().Message}", "Услуги на стороне", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (rows.Count == 0)
        {
            MessageBox.Show("В 44 секции НЗП не найдены детали в единице измерения шт.", "Услуги на стороне", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var request = ShowExternalServicesDialog(rows);
        if (request is null)
        {
            return;
        }

        try
        {
            IsBusy = true;
            ProductionLaunchStatusText = "Создание перемещения 1С для услуг на стороне...";
            var result = await oneCGoodsTransferService.CreateTransferAsync(request, CancellationToken.None);
            ProductionLaunchStatusText = $"Создано перемещение 1С {result.Number}; проведено: {(result.Posted ? "да" : "нет")}.";
            StatusText = ProductionLaunchStatusText;
            MessageBox.Show(
                $"Перемещение 1С создано.\nНомер: {result.Number}\nСтрок: {request.Items.Count}\nДокумент не проведен.\n\nКомментарий:\n{result.Comment}",
                "Услуги на стороне",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Could not create 1C goods transfer for external services");
            ProductionLaunchStatusText = $"Ошибка перемещения: {ex.GetBaseException().Message}";
            MessageBox.Show($"Не удалось создать перемещение 1С.\n{ex.GetBaseException().Message}", "Услуги на стороне", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
        }
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

    public async Task<ProductionLaunchPreview> BuildProductionLaunchPreviewAsync(MskLibraryRow row)
    {
        var part = await dbContext.Parts
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .FirstOrDefaultAsync(x => x.Ips == row.Ips);
        if (part is null)
        {
            throw new InvalidOperationException($"Деталь IPS {row.Ips} не найдена в библиотеке.");
        }

        var map = part.BlankMaps.FirstOrDefault(x => x.IsActive && x.IsPrimary) ?? part.BlankMaps.FirstOrDefault(x => x.IsActive);
        if (map?.CanonicalBlank is null)
        {
            throw new InvalidOperationException($"Для IPS {row.Ips} не задана активная заготовка в библиотеке.");
        }

        var alias = map.CanonicalBlank.Aliases.FirstOrDefault(x => x.IsActive) ?? map.CanonicalBlank.Aliases.FirstOrDefault();
        if (alias is null || string.IsNullOrWhiteSpace(alias.OneCCode))
        {
            throw new InvalidOperationException($"Для заготовки IPS {row.Ips} не найден УТ-код 1С.");
        }

        var snapshot = await dbContext.StockSnapshots
            .AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .ThenByDescending(x => x.Id)
            .FirstOrDefaultAsync();
        if (snapshot is null)
        {
            throw new InvalidOperationException("Нет снимка остатков 1С. Нажмите \"Обновить данные 1С\".");
        }

        var stockItems = await dbContext.StockItems
            .AsNoTracking()
            .Where(x => x.StockSnapshotId == snapshot.Id)
            .ToListAsync();
        var blankCodeKey = StockCodeNormalizer.NormalizeForComparison(alias.OneCCode);
        var materialStock = stockItems
            .Where(x =>
                StockWarehouseRules.IsProductionLaunchMaterialWarehouse(x.Warehouse) &&
                StockCodeNormalizer.NormalizeForComparison(x.OneCCode) == blankCodeKey &&
                x.Unit == map.ConsumptionUnit)
            .Sum(x => x.Quantity);
        var consumption = map.ConsumptionQuantity <= 0 ? 0m : map.ConsumptionQuantity;
        var maxQuantity = CalculateProductionLaunchMaxQuantity(map.ConsumptionUnit, consumption, materialStock);
        var shortageRows = await BuildDemandShortageRowsAsync(row.Ips, stockItems);
        var defaultQuantity = Math.Min(maxQuantity, Math.Max(1m, decimal.Floor(shortageRows.Sum(x => x.Quantity))));
        if (defaultQuantity <= 0)
        {
            defaultQuantity = maxQuantity;
        }

        var componentName = FirstNotEmpty(alias.SourceName, map.CanonicalBlank.CanonicalName);
        var materialLine = $"Материал: {componentName}; код {alias.OneCCode}; остаток склада: {FormatDecimal(materialStock)} {UiText.DisplayUnit(map.ConsumptionUnit)}; норма: {FormatDecimal(consumption)} {UiText.DisplayUnit(map.ConsumptionUnit)} на деталь.";
        return new ProductionLaunchPreview(
            row.Ips,
            row.Designation,
            FirstNotEmpty(part.Name, row.Name),
            alias.OneCCode,
            componentName,
            consumption,
            map.ConsumptionUnit,
            materialStock,
            maxQuantity,
            defaultQuantity,
            materialLine,
            shortageRows);
    }

    private async Task LoadDetailMetricsAsync(MskLibraryRow? row, int version)
    {
        if (row is null)
        {
            return;
        }

        try
        {
            var batchId = await dbContext.DemandBatches
                .AsNoTracking()
                .OrderByDescending(x => x.ImportedAt)
                .ThenByDescending(x => x.Id)
                .Select(x => (long?)x.Id)
                .FirstOrDefaultAsync();
            var demandQuantities = batchId is null
                ? []
                : await dbContext.DemandItems
                    .AsNoTracking()
                    .Where(x => x.DemandBatchId == batchId.Value && x.Ips == row.Ips)
                    .Select(x => x.Quantity)
                    .ToListAsync();
            var demandQuantity = demandQuantities.Sum();

            var snapshotId = await dbContext.StockSnapshots
                .AsNoTracking()
                .OrderByDescending(x => x.ImportedAt)
                .ThenByDescending(x => x.Id)
                .Select(x => (long?)x.Id)
                .FirstOrDefaultAsync();
            var inProduction = 0m;
            if (snapshotId is not null)
            {
                var ipsKey = StockCodeNormalizer.NormalizeForComparison(row.Ips);
                var stockItems = await dbContext.StockItems
                    .AsNoTracking()
                    .Where(x => x.StockSnapshotId == snapshotId.Value && x.Unit == MeasurementUnit.Piece)
                    .ToListAsync();
                inProduction = stockItems
                    .Where(x =>
                        StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse) &&
                        StockCodeNormalizer.NormalizeForComparison(x.OneCCode) == ipsKey)
                    .Sum(x => x.Quantity);
            }

            var availableLaunchText = "нет данных";
            try
            {
                var preview = await BuildProductionLaunchPreviewAsync(row);
                availableLaunchText = $"{FormatDecimal(preview.MaxQuantity)} шт";
            }
            catch (Exception ex)
            {
                availableLaunchText = $"не рассчитано: {ex.GetBaseException().Message}";
            }

            if (version == Volatile.Read(ref detailMetricsVersion) && SelectedRow?.Ips == row.Ips)
            {
                DetailMetricsText =
                    $"В потребности сейчас: {FormatDecimal(demandQuantity)} шт\n" +
                    $"В производстве: {FormatDecimal(inProduction)} шт\n" +
                    $"Доступно для запуска: {availableLaunchText}";
            }
        }
        catch (Exception ex)
        {
            if (version == Volatile.Read(ref detailMetricsVersion) && SelectedRow?.Ips == row.Ips)
            {
                DetailMetricsText = $"Показатели не рассчитаны: {ex.GetBaseException().Message}";
            }
        }
    }

    private async Task<IReadOnlyList<ProductionLaunchShortageRow>> BuildDemandShortageRowsAsync(string ips, IReadOnlyList<StockItem> stockItems)
    {
        var batch = await dbContext.DemandBatches
            .AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .ThenByDescending(x => x.Id)
            .FirstOrDefaultAsync();
        if (batch is null)
        {
            return [];
        }

        var ipsKey = StockCodeNormalizer.NormalizeForComparison(ips);
        var wip = stockItems
            .Where(x =>
                StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse) &&
                StockCodeNormalizer.NormalizeForComparison(x.OneCCode) == ipsKey &&
                x.Unit == MeasurementUnit.Piece)
            .Sum(x => x.Quantity);
        var demand = await dbContext.DemandItems
            .AsNoTracking()
            .Where(x => x.DemandBatchId == batch.Id && x.Ips == ips)
            .OrderBy(x => x.DemandDate ?? DateTime.MaxValue)
            .ThenBy(x => x.Id)
            .ToListAsync();
        var rows = new List<ProductionLaunchShortageRow>();
        foreach (var item in demand)
        {
            var coveredByWip = Math.Min(item.Quantity, Math.Max(0m, wip));
            wip -= coveredByWip;
            var deficit = Math.Max(0m, item.Quantity - coveredByWip);
            if (deficit <= 0)
            {
                continue;
            }

            rows.Add(new ProductionLaunchShortageRow(
                deficit,
                FirstNotEmpty(item.Project, "проект не указан"),
                FirstNotEmpty(item.ProductionSystem, item.SerialNumber, "№ станка не указан"),
                item.DemandDate));
        }

        return rows;
    }

    public async Task<IReadOnlyList<ExternalServiceWipRow>> LoadExternalServiceWipRowsAsync()
    {
        var snapshot = await dbContext.StockSnapshots
            .AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.Id)
            .FirstOrDefaultAsync();
        if (snapshot is null)
        {
            throw new InvalidOperationException("Нет снимка остатков 1С. Нажмите \"Обновить данные 1С\".");
        }

        var stockRows = await dbContext.StockItems
            .AsNoTracking()
            .Where(x => x.StockSnapshotId == snapshot.Id && x.Unit == MeasurementUnit.Piece && x.Quantity > 0)
            .ToListAsync();
        stockRows = stockRows
            .Where(x => IsExternalServiceSourceWarehouse(x.Warehouse))
            .ToList();

        var parts = await dbContext.Parts
            .AsNoTracking()
            .ToListAsync();
        var partsByCode = parts
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.Ips), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.OrderByDescending(p => p.UpdatedAt).First(), StringComparer.OrdinalIgnoreCase);

        var demandBatchId = await dbContext.DemandBatches
            .AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .ThenByDescending(x => x.Id)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        var demandByIps = new Dictionary<string, List<DemandItem>>(StringComparer.OrdinalIgnoreCase);
        if (demandBatchId is not null)
        {
            var demandItems = await dbContext.DemandItems
                .AsNoTracking()
                .Where(x => x.DemandBatchId == demandBatchId.Value && x.Quantity > 0)
                .OrderBy(x => x.DemandDate ?? DateTime.MaxValue)
                .ThenBy(x => x.Id)
                .ToListAsync();
            demandByIps = demandItems
                .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.Ips), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.ToList(), StringComparer.OrdinalIgnoreCase);
        }

        var result = new List<ExternalServiceWipRow>();
        foreach (var group in stockRows
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .Where(x => !string.IsNullOrWhiteSpace(x.Key)))
        {
            var first = group.OrderBy(x => x.SourceName).First();
            var quantity = group.Sum(x => x.Quantity);
            partsByCode.TryGetValue(group.Key, out var part);
            if (part is null || !HasAnyExternalServiceRequirement(part))
            {
                continue;
            }

            var requiredPart = part;
            var ips = requiredPart.Ips;
            var designation = requiredPart.Designation ?? string.Empty;
            var name = FirstNotEmpty(requiredPart.Name, first.SourceName, ips);
            demandByIps.TryGetValue(group.Key, out var demandRows);
            var hasDemandRequirement = demandRows?.Any(x => x.Quantity > 0) == true;
            result.Add(new ExternalServiceWipRow(
                first.OneCCode,
                ips,
                designation,
                name,
                quantity,
                string.Empty,
                string.Empty,
                null,
                BuildExternalServiceDemandAllocations(quantity, demandRows ?? []),
                requiredPart.RequiresNitriding,
                requiredPart.RequiresHeatTreatment,
                requiredPart.RequiresChemicalOxidation,
                requiredPart.RequiresKeyway,
                hasDemandRequirement));
        }

        return result
            .OrderBy(x => x.Ips, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool HasAnyExternalServiceRequirement(Part part) =>
        part.RequiresNitriding ||
        part.RequiresHeatTreatment ||
        part.RequiresChemicalOxidation ||
        part.RequiresKeyway;

    public static bool MatchesExternalServiceType(ExternalServiceWipRow row, string? serviceType)
    {
        var service = UiSearchText.Normalize(serviceType);
        if (service.Length == 0)
        {
            return row.HasExternalServiceRequirement;
        }

        if (service.Contains(UiSearchText.Normalize("Азотирование"), StringComparison.Ordinal))
        {
            return row.RequiresNitriding;
        }

        if (service.Contains(UiSearchText.Normalize("Термообработка"), StringComparison.Ordinal) ||
            service == UiSearchText.Normalize("ТО"))
        {
            return row.RequiresHeatTreatment;
        }

        if (service.Contains(UiSearchText.Normalize("Хим"), StringComparison.Ordinal) ||
            service.Contains(UiSearchText.Normalize("Окс"), StringComparison.Ordinal))
        {
            return row.RequiresChemicalOxidation;
        }

        if (service.Contains(UiSearchText.Normalize("Шпон"), StringComparison.Ordinal) ||
            service.Contains(UiSearchText.Normalize("Паз"), StringComparison.Ordinal))
        {
            return row.RequiresKeyway;
        }

        return row.HasExternalServiceRequirement;
    }

    public static decimal CalculateExternalServiceDemandQuantity(ExternalServiceWipRow row)
    {
        var stockMarker = UiSearchText.Normalize("на склад");
        var demandQuantity = row.DemandAllocations
            .Where(x => x.DemandDate is not null && UiSearchText.Normalize(x.MachineNumber) != stockMarker)
            .Sum(x => x.Quantity);
        if (demandQuantity <= 0)
        {
            return 0m;
        }

        return Math.Min(row.AvailableQuantity, demandQuantity);
    }

    public static IReadOnlyList<ExternalServiceDemandAllocation> BuildExternalServiceDemandAllocations(decimal availableQuantity, IReadOnlyList<DemandItem> demandRows)
    {
        var result = new List<ExternalServiceDemandAllocation>();
        var remaining = Math.Max(0m, availableQuantity);
        foreach (var demand in demandRows
            .Where(x => x.Quantity > 0)
            .OrderBy(x => x.DemandDate ?? DateTime.MaxValue)
            .ThenBy(x => x.Id))
        {
            if (remaining <= 0)
            {
                break;
            }

            var quantity = Math.Min(remaining, demand.Quantity);
            if (quantity <= 0)
            {
                continue;
            }

            result.Add(new ExternalServiceDemandAllocation(
                FirstNotEmpty(demand.Project, "на склад"),
                FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber, "на склад"),
                quantity,
                demand.DemandDate));
            remaining -= quantity;
        }

        if (remaining > 0)
        {
            result.Add(new ExternalServiceDemandAllocation("на склад", "на склад", remaining, null));
        }

        return result;
    }

    private OneCGoodsTransferRequest? ShowExternalServicesDialog(IReadOnlyList<ExternalServiceWipRow> sourceRows)
    {
        var rows = new ObservableCollection<ExternalServiceWipRow>(sourceRows);
        var serviceTypes = UiReferenceData.ServiceTypes();
        var window = new Window
        {
            Title = "Услуги на стороне",
            Width = 980,
            Height = 720,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Application.Current?.MainWindow,
            ResizeMode = ResizeMode.CanResize
        };

        var root = new DockPanel { Margin = new Thickness(18) };
        var buttons = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        DockPanel.SetDock(buttons, Dock.Bottom);
        root.Children.Add(buttons);

        var content = new DockPanel();
        root.Children.Add(content);
        var top = new StackPanel { Margin = new Thickness(0, 0, 0, 10) };
        DockPanel.SetDock(top, Dock.Top);
        content.Children.Add(top);

        top.Children.Add(new TextBlock { Text = "Выберите продукцию и вид услуги для отправки на сторону", FontSize = 18, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 10) });
        var firstLine = new WrapPanel();
        firstLine.Children.Add(new TextBlock { Text = "Вид услуги:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 6) });
        var serviceTypeBox = new ComboBox { Width = 180, ItemsSource = serviceTypes, SelectedIndex = 0, Margin = new Thickness(0, 0, 14, 6) };
        firstLine.Children.Add(serviceTypeBox);
        firstLine.Children.Add(new TextBlock { Text = "Поиск:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 6) });
        var searchBox = new TextBox { Width = 260, Margin = new Thickness(0, 0, 14, 6) };
        firstLine.Children.Add(searchBox);
        top.Children.Add(firstLine);
        var demandOptionsLine = new StackPanel { Margin = new Thickness(0, 0, 0, 2) };
        var onlyInDemandBox = new System.Windows.Controls.CheckBox { Content = "Только в потребности", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 14, 4) };
        var autoSelectBox = new System.Windows.Controls.CheckBox { Content = "Авто выбор", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 14, 6) };
        demandOptionsLine.Children.Add(onlyInDemandBox);
        demandOptionsLine.Children.Add(autoSelectBox);
        top.Children.Add(demandOptionsLine);

        top.Children.Add(new TextBlock { Text = "Комментарий для 1С:", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 4) });
        var commentBox = new TextBox { Height = 60, TextWrapping = TextWrapping.Wrap, AcceptsReturn = true };
        top.Children.Add(commentBox);
        var commentEdited = false;
        var updatingComment = false;
        bool MatchesDialogFilters(ExternalServiceWipRow row) =>
            MatchesExternalServiceType(row, serviceTypeBox.SelectedItem as string) &&
            (!onlyInDemandBox.IsChecked.GetValueOrDefault() || row.HasDemandRequirement);

        IReadOnlyList<ExternalServiceWipRow> SelectedRows() => rows
            .Where(x => x.IsSelected && MatchesDialogFilters(x))
            .ToList();

        void ApplyDemandQuantities(bool selectRows)
        {
            foreach (var row in rows)
            {
                var demandQuantity = CalculateExternalServiceDemandQuantity(row);
                if (demandQuantity > 0 && MatchesExternalServiceType(row, serviceTypeBox.SelectedItem as string))
                {
                    row.QuantityText = FormatDecimal(demandQuantity);
                    if (selectRows)
                    {
                        row.IsSelected = MatchesDialogFilters(row);
                    }
                }
                else if (selectRows)
                {
                    row.IsSelected = false;
                }
            }
        }

        void RefreshExternalServiceComment(bool force)
        {
            if (!force && commentEdited)
            {
                return;
            }

            updatingComment = true;
            commentBox.Text = BuildExternalServiceComment(serviceTypeBox.SelectedItem as string ?? serviceTypes[0], SelectedRows());
            updatingComment = false;
        }

        commentBox.TextChanged += (_, _) =>
        {
            if (!updatingComment)
            {
                commentEdited = true;
            }
        };
        serviceTypeBox.SelectionChanged += (_, _) =>
        {
            if (onlyInDemandBox.IsChecked.GetValueOrDefault())
            {
                ApplyDemandQuantities(selectRows: autoSelectBox.IsChecked.GetValueOrDefault());
            }

            foreach (var row in rows.Where(x => !MatchesDialogFilters(x)))
            {
                row.IsSelected = false;
            }

            CollectionViewSource.GetDefaultView(rows)?.Refresh();
            RefreshExternalServiceComment(force: false);
        };
        onlyInDemandBox.Checked += (_, _) =>
        {
            ApplyDemandQuantities(selectRows: autoSelectBox.IsChecked.GetValueOrDefault());
            foreach (var row in rows.Where(x => !MatchesDialogFilters(x)))
            {
                row.IsSelected = false;
            }

            CollectionViewSource.GetDefaultView(rows)?.Refresh();
            RefreshExternalServiceComment(force: false);
        };
        onlyInDemandBox.Unchecked += (_, _) =>
        {
            CollectionViewSource.GetDefaultView(rows)?.Refresh();
            RefreshExternalServiceComment(force: false);
        };
        autoSelectBox.Checked += (_, _) =>
        {
            ApplyDemandQuantities(selectRows: true);
            CollectionViewSource.GetDefaultView(rows)?.Refresh();
            RefreshExternalServiceComment(force: false);
        };
        autoSelectBox.Unchecked += (_, _) =>
        {
            foreach (var row in rows.Where(x => x.HasDemandRequirement && MatchesExternalServiceType(x, serviceTypeBox.SelectedItem as string)))
            {
                row.IsSelected = false;
            }

            CollectionViewSource.GetDefaultView(rows)?.Refresh();
            RefreshExternalServiceComment(force: false);
        };
        foreach (var row in rows)
        {
            row.PropertyChanged += (_, args) =>
            {
                if (args.PropertyName is nameof(ExternalServiceWipRow.IsSelected) or nameof(ExternalServiceWipRow.QuantityText))
                {
                    RefreshExternalServiceComment(force: false);
                }
            };
        }

        var status = new TextBlock { Foreground = Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 0) };
        top.Children.Add(status);

        var grid = new DataGrid
        {
            AutoGenerateColumns = false,
            ItemsSource = rows,
            SelectionUnit = DataGridSelectionUnit.Cell,
            SelectionMode = DataGridSelectionMode.Extended,
            CanUserAddRows = false,
            HeadersVisibility = DataGridHeadersVisibility.Column,
            IsReadOnly = false
        };
        var selectCellFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.CheckBox));
        selectCellFactory.SetBinding(System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty, new Binding(nameof(ExternalServiceWipRow.IsSelected)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
        selectCellFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        selectCellFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
        selectCellFactory.SetValue(UIElement.FocusableProperty, false);
        grid.Columns.Add(new DataGridTemplateColumn { Header = "", CellTemplate = new DataTemplate { VisualTree = selectCellFactory }, Width = 36 });
        grid.Columns.Add(new DataGridTextColumn { Header = "IPS/код", Binding = new Binding(nameof(ExternalServiceWipRow.Ips)), IsReadOnly = true, Width = 110 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Обозначение", Binding = new Binding(nameof(ExternalServiceWipRow.Designation)), IsReadOnly = true, Width = 180 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Наименование", Binding = new Binding(nameof(ExternalServiceWipRow.Name)), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        grid.Columns.Add(new DataGridTextColumn { Header = "Остаток", Binding = new Binding(nameof(ExternalServiceWipRow.AvailableText)), IsReadOnly = true, Width = 90 });
        grid.Columns.Add(new DataGridTextColumn { Header = "Переместить", Binding = new Binding(nameof(ExternalServiceWipRow.QuantityText)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged }, Width = 110 });
        content.Children.Add(grid);

        var view = CollectionViewSource.GetDefaultView(rows);
        view.Filter = item =>
        {
            if (item is not ExternalServiceWipRow row)
            {
                return false;
            }

            var search = searchBox.Text.Trim();
            if (!MatchesDialogFilters(row))
            {
                return false;
            }

            return string.IsNullOrWhiteSpace(search) ||
                UiSearchText.Contains(row.Ips, search) ||
                UiSearchText.Contains(row.OneCCode, search) ||
                UiSearchText.Contains(row.Designation, search) ||
                UiSearchText.Contains(row.Name, search) ||
                UiSearchText.Contains(row.Project, search) ||
                UiSearchText.Contains(row.MachineNumber, search);
        };
        searchBox.TextChanged += (_, _) => view.Refresh();

        var createButton = new Button { Content = "Создать перемещение 1С", MinWidth = 180, Margin = new Thickness(6, 0, 0, 0) };
        var exportRequestButton = new Button { Content = "Сформировать заявку", MinWidth = 150, Margin = new Thickness(6, 0, 0, 0) };
        var cancelButton = new Button { Content = "Отмена", MinWidth = 90, Margin = new Thickness(6, 0, 0, 0) };
        buttons.Children.Add(cancelButton);
        buttons.Children.Add(exportRequestButton);
        buttons.Children.Add(createButton);

        OneCGoodsTransferRequest? result = null;
        cancelButton.Click += (_, _) => window.Close();
        exportRequestButton.Click += (_, _) =>
        {
            var serviceType = serviceTypeBox.SelectedItem as string ?? serviceTypes[0];
            var selectedRows = SelectedRows();
            if (!TryBuildExternalServiceTransferRequest(serviceType, commentBox.Text, selectedRows, out _, out var error))
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = error;
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Сформировать заявку на услугу",
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"Заявка на услугу ЦМО от {DateTime.Now:dd.MM.yyyy}.xlsx"
            };
            if (dialog.ShowDialog(window) != true)
            {
                return;
            }

            try
            {
                ExportExternalServiceRequestForm(dialog.FileName, serviceType, selectedRows);
                status.Foreground = Brushes.SeaGreen;
                status.Text = $"Заявка сформирована: {dialog.FileName}";
                Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = $"Не удалось сформировать заявку: {ex.GetBaseException().Message}";
            }
        };
        createButton.Click += async (_, _) =>
        {
            var serviceType = serviceTypeBox.SelectedItem as string ?? serviceTypes[0];
            var selectedRows = SelectedRows();
            RefreshExternalServiceComment(force: string.IsNullOrWhiteSpace(commentBox.Text));
            if (!TryBuildExternalServiceTransferRequest(serviceType, commentBox.Text, selectedRows, out var request, out var error))
            {
                status.Foreground = Brushes.Firebrick;
                status.Text = error;
                return;
            }

            try
            {
                createButton.IsEnabled = false;
                IsBusy = true;
                ProductionLaunchStatusText = "Создание перемещения 1С для услуг на стороне...";
                status.Foreground = Brushes.DimGray;
                status.Text = ProductionLaunchStatusText;
                var transferResult = await oneCGoodsTransferService!.CreateTransferAsync(request!, CancellationToken.None);
                ProductionLaunchStatusText = $"Создано перемещение 1С {transferResult.Number}; проведено: {(transferResult.Posted ? "да" : "нет")}.";
                StatusText = ProductionLaunchStatusText;
                status.Foreground = Brushes.SeaGreen;
                status.Text = $"{ProductionLaunchStatusText} Окно оставлено открытым, можно сформировать заявку.";
                MessageBox.Show(
                    $"Перемещение 1С создано.\nНомер: {transferResult.Number}\nСтрок: {request!.Items.Count}\nДокумент не проведен.\n\nКомментарий:\n{transferResult.Comment}",
                    "Услуги на стороне",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Could not create 1C goods transfer for external services");
                ProductionLaunchStatusText = $"Ошибка перемещения: {ex.GetBaseException().Message}";
                status.Foreground = Brushes.Firebrick;
                status.Text = ProductionLaunchStatusText;
                MessageBox.Show($"Не удалось создать перемещение 1С.\n{ex.GetBaseException().Message}", "Услуги на стороне", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                createButton.IsEnabled = true;
            }
        };

        window.Content = root;
        RefreshExternalServiceComment(force: true);
        window.ShowDialog();
        return result;
    }

    public static bool TryBuildExternalServiceTransferRequest(
        string serviceType,
        string comment,
        IReadOnlyList<ExternalServiceWipRow> selectedRows,
        out OneCGoodsTransferRequest? request,
        out string error)
    {
        request = null;
        error = string.Empty;
        var items = new List<OneCGoodsTransferItem>();
        foreach (var row in selectedRows)
        {
            if (!TryParseExternalServiceQuantity(row, out var quantity, out error))
            {
                return false;
            }

            items.Add(new OneCGoodsTransferItem(row.OneCCode, row.Name, quantity));
        }

        if (items.Count == 0)
        {
            error = "Выберите хотя бы одну деталь из НЗП.";
            return false;
        }

        var service = string.IsNullOrWhiteSpace(serviceType) ? "Услуги на стороне" : serviceType.Trim();
        request = new OneCGoodsTransferRequest(
            service,
            string.IsNullOrWhiteSpace(comment) ? BuildExternalServiceComment(service, selectedRows) : comment.Trim(),
            items
                .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
                .Select(x => new OneCGoodsTransferItem(
                    x.First().OneCCode,
                    x.First().Name,
                    x.Sum(i => i.Quantity)))
                .ToList());
        return true;
    }

    public static string BuildExternalServiceComment(string serviceType, IReadOnlyList<ExternalServiceWipRow> selectedRows)
    {
        var service = string.IsNullOrWhiteSpace(serviceType) ? "Услуги на стороне" : serviceType.Trim();
        var parts = BuildExternalServiceRequestRows(service, selectedRows)
            .GroupBy(row => row.Ips, StringComparer.OrdinalIgnoreCase)
            .Select(group =>
            {
                var machines = group
                    .Select(row => $"{FirstNotEmpty(row.MachineNumber, "на склад")} - {FormatDecimal(row.Quantity)} шт")
                    .ToList();
                return $"IPS {group.Key} {string.Join("; ", machines)}";
            })
            .ToList();
        return parts.Count == 0
            ? $"Услуги на стороне: {service}"
            : $"Услуги на стороне: {service} {string.Join(". ", parts)}";
    }

    public static DateTime CalculateExternalServiceReadyDate(string serviceType, DateTime formedAt) =>
        formedAt.Date.AddDays(GetExternalServiceLeadTimeDays(serviceType));

    public static int GetExternalServiceLeadTimeDays(string serviceType)
    {
        var service = UiSearchText.Normalize(serviceType);
        return service.Contains(UiSearchText.Normalize("Азотирование"), StringComparison.Ordinal)
            ? 10
            : 7;
    }

    public static IReadOnlyList<ExternalServiceRequestRow> BuildExternalServiceRequestRows(string serviceType, IReadOnlyList<ExternalServiceWipRow> selectedRows, DateTime? formedAt = null)
    {
        var service = string.IsNullOrWhiteSpace(serviceType) ? "Услуги на стороне" : serviceType.Trim();
        var readyDate = CalculateExternalServiceReadyDate(service, formedAt ?? DateTime.Today);
        var result = new List<ExternalServiceRequestRow>();
        foreach (var row in selectedRows.Where(x => x.IsSelected))
        {
            if (!TryParseExternalServiceQuantity(row, out var selectedQuantity, out _) || selectedQuantity <= 0)
            {
                continue;
            }

            var remaining = selectedQuantity;
            var allocations = row.DemandAllocations.Count > 0
                ? row.DemandAllocations
                : [new ExternalServiceDemandAllocation("на склад", "на склад", row.AvailableQuantity, null)];
            foreach (var allocation in allocations)
            {
                if (remaining <= 0)
                {
                    break;
                }

                var quantity = Math.Min(remaining, allocation.Quantity);
                if (quantity <= 0)
                {
                    continue;
                }

                var isUrgent = allocation.DemandDate is not null && readyDate > allocation.DemandDate.Value.Date;
                result.Add(new ExternalServiceRequestRow(
                    allocation.Project,
                    allocation.MachineNumber,
                    row.Ips,
                    row.Nomenclature,
                    "шт",
                    quantity,
                    service,
                    allocation.DemandDate,
                    readyDate,
                    isUrgent,
                    isUrgent ? "СРОЧНО!" : string.Empty));
                remaining -= quantity;
            }

            if (remaining > 0)
            {
                result.Add(new ExternalServiceRequestRow("на склад", "на склад", row.Ips, row.Nomenclature, "шт", remaining, service, null, readyDate, false, string.Empty));
            }
        }

        return result;
    }

    public static void ExportExternalServiceRequestForm(string filePath, string serviceType, IReadOnlyList<ExternalServiceWipRow> selectedRows, DateTime? formedAt = null)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Заявка");
        var headers = new[]
        {
            "№",
            "ПС",
            "№ станка",
            "Код IPS",
            "Номенклатура",
            "Ед. измерения",
            "Кол-во",
            "Обоснование",
            "Дата потребности",
            "Примечание",
            "Цена, руб",
            "Стоимость, руб"
        };
        for (var column = 0; column < headers.Length; column++)
        {
            sheet.Cells[1, column + 1].Value = headers[column];
        }

        using (var range = sheet.Cells[1, 1, 1, headers.Length])
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(230, 236, 245));
            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        var service = string.IsNullOrWhiteSpace(serviceType) ? "Услуги на стороне" : serviceType.Trim();
        var expandedRowNumber = 2;
        var expandedIndex = 1;
        foreach (var requestRow in BuildExternalServiceRequestRows(serviceType, selectedRows, formedAt))
        {
            sheet.Cells[expandedRowNumber, 1].Value = expandedIndex++;
            sheet.Cells[expandedRowNumber, 2].Value = requestRow.Project;
            sheet.Cells[expandedRowNumber, 3].Value = requestRow.MachineNumber;
            sheet.Cells[expandedRowNumber, 4].Value = requestRow.Ips;
            sheet.Cells[expandedRowNumber, 5].Value = requestRow.Nomenclature;
            sheet.Cells[expandedRowNumber, 6].Value = requestRow.UnitName;
            sheet.Cells[expandedRowNumber, 7].Value = (double)requestRow.Quantity;
            sheet.Cells[expandedRowNumber, 8].Value = requestRow.Justification;
            sheet.Cells[expandedRowNumber, 9].Value = requestRow.ServiceReadyDate;
            sheet.Cells[expandedRowNumber, 9].Style.Numberformat.Format = "dd.mm.yyyy";
            if (requestRow.IsUrgent)
            {
                sheet.Cells[expandedRowNumber, 9].Style.Font.Bold = true;
                sheet.Cells[expandedRowNumber, 9].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            }

            sheet.Cells[expandedRowNumber, 10].Value = requestRow.Note;
            if (requestRow.IsUrgent)
            {
                sheet.Cells[expandedRowNumber, 10].Style.Font.Bold = true;
                sheet.Cells[expandedRowNumber, 10].Style.Font.Color.SetColor(System.Drawing.Color.Red);
            }

            sheet.Cells[expandedRowNumber, 11].Value = string.Empty;
            sheet.Cells[expandedRowNumber, 12].Value = string.Empty;
            expandedRowNumber++;
        }

        sheet.View.FreezePanes(2, 1);
        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        package.SaveAs(new FileInfo(filePath));
        return;

        var rowNumber = 2;
        var index = 1;
        foreach (var row in selectedRows.Where(x => x.IsSelected))
        {
            if (!TryParseExternalServiceQuantity(row, out var quantity, out _) || quantity <= 0)
            {
                continue;
            }

            sheet.Cells[rowNumber, 1].Value = index++;
            sheet.Cells[rowNumber, 2].Value = row.Project;
            sheet.Cells[rowNumber, 3].Value = row.MachineNumber;
            sheet.Cells[rowNumber, 4].Value = row.Ips;
            sheet.Cells[rowNumber, 5].Value = row.Nomenclature;
            sheet.Cells[rowNumber, 6].Value = "шт";
            sheet.Cells[rowNumber, 7].Value = (double)quantity;
            sheet.Cells[rowNumber, 8].Value = service;
            if (row.DemandDate is not null)
            {
                sheet.Cells[rowNumber, 9].Value = row.DemandDate.Value;
                sheet.Cells[rowNumber, 9].Style.Numberformat.Format = "dd.mm.yyyy";
            }

            sheet.Cells[rowNumber, 10].Value = string.Empty;
            sheet.Cells[rowNumber, 11].Value = string.Empty;
            rowNumber++;
        }

        sheet.View.FreezePanes(2, 1);
        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        package.SaveAs(new FileInfo(filePath));
    }

    private static bool TryParseExternalServiceQuantity(ExternalServiceWipRow row, out decimal quantity, out string error)
    {
        quantity = 0m;
        error = string.Empty;
        if (!TryParseProductionLaunchQuantity(row.QuantityText, out quantity) || quantity <= 0)
        {
            error = $"Некорректное количество для {row.Ips}.";
            return false;
        }

        if (quantity != decimal.Floor(quantity))
        {
            error = $"Для услуг на стороне количество должно быть целым в шт: {row.Ips}.";
            return false;
        }

        if (quantity > row.AvailableQuantity)
        {
            error = $"Нельзя переместить больше остатка НЗП для {row.Ips}: доступно {FormatDecimal(row.AvailableQuantity)} шт.";
            return false;
        }

        return true;
    }

    private static int SelectExternalServiceRowsByText(IEnumerable<ExternalServiceWipRow> rows, string text)
    {
        var keys = Regex.Split(text, @"[\s,;]+")
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .SelectMany(x => new[] { UiSearchText.Normalize(x), StockCodeNormalizer.NormalizeForComparison(x) })
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (keys.Count == 0)
        {
            return 0;
        }

        var count = 0;
        foreach (var row in rows)
        {
            var matches = keys.Contains(UiSearchText.Normalize(row.Ips)) ||
                keys.Contains(StockCodeNormalizer.NormalizeForComparison(row.Ips)) ||
                keys.Contains(StockCodeNormalizer.NormalizeForComparison(row.OneCCode)) ||
                keys.Contains(UiSearchText.Normalize(row.Designation));
            row.IsSelected = matches;
            if (matches)
            {
                count++;
            }
        }

        return count;
    }

    private static bool IsExternalServiceSourceWarehouse(string? warehouse)
    {
        var text = (warehouse ?? string.Empty).Trim();
        return text.StartsWith("44", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("44 секция НЗП", StringComparison.OrdinalIgnoreCase);
    }

    private static string DisplayIpsFromOneCCode(string oneCCode)
    {
        var normalized = StockCodeNormalizer.NormalizeForComparison(oneCCode);
        return normalized.Length == 11 && normalized.All(char.IsDigit)
            ? normalized.TrimStart('0')
            : oneCCode;
    }

    private ProductionLaunchDialogResult? ShowProductionLaunchDialog(ProductionLaunchPreview preview)
    {
        var window = new Window
        {
            Title = "Запустить в производство",
            Width = 640,
            Height = 560,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = Application.Current?.MainWindow,
            ResizeMode = ResizeMode.NoResize
        };
        var root = new DockPanel { Margin = new Thickness(18) };
        var title = new TextBlock
        {
            Text = $"{preview.Ips} {preview.Designation} {preview.PartName}".Trim(),
            FontSize = 18,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 12)
        };
        DockPanel.SetDock(title, Dock.Top);
        root.Children.Add(title);

        var buttons = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        DockPanel.SetDock(buttons, Dock.Bottom);
        root.Children.Add(buttons);

        var content = new StackPanel();
        root.Children.Add(content);
        content.Children.Add(new TextBlock { Text = preview.MaterialLine, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
        content.Children.Add(new TextBlock { Text = $"Доступно деталей для запуска: {FormatDecimal(preview.MaxQuantity)}", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 8) });

        var makeMax = new System.Windows.Controls.CheckBox { Content = "Сделать сколько можно", IsChecked = true, Margin = new Thickness(0, 4, 0, 8) };
        content.Children.Add(makeMax);
        var piecewise = new System.Windows.Controls.CheckBox { Content = "Запуск по 1 шт", IsChecked = false, Margin = new Thickness(0, 0, 0, 8) };
        content.Children.Add(piecewise);
        var quantityPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
        quantityPanel.Children.Add(new TextBlock { Text = "Количество деталей:", VerticalAlignment = VerticalAlignment.Center, Width = 150 });
        var quantityBox = new System.Windows.Controls.TextBox { Width = 120, Text = FormatDecimal(preview.MaxQuantity), IsEnabled = false };
        quantityPanel.Children.Add(quantityBox);
        content.Children.Add(quantityPanel);
        var status = new TextBlock { Foreground = System.Windows.Media.Brushes.Firebrick, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) };
        content.Children.Add(status);

        content.Children.Add(new TextBlock { Text = "Комментарий для 1С:", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 4, 0, 4) });
        var commentBox = new System.Windows.Controls.TextBox
        {
            IsReadOnly = false,
            TextWrapping = TextWrapping.Wrap,
            AcceptsReturn = true,
            Height = 185,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled
        };
        content.Children.Add(commentBox);

        var launchButton = new System.Windows.Controls.Button { Content = "Запустить в производство", MinWidth = 180, Margin = new Thickness(6, 0, 0, 0) };
        var cancelButton = new System.Windows.Controls.Button { Content = "Отмена", MinWidth = 90, Margin = new Thickness(6, 0, 0, 0) };
        buttons.Children.Add(cancelButton);
        buttons.Children.Add(launchButton);

        ProductionLaunchDialogResult? result = null;
        decimal CurrentQuantity()
        {
            if (makeMax.IsChecked == true)
            {
                return preview.MaxQuantity;
            }

            return TryParseProductionLaunchQuantity(quantityBox.Text, out var quantity) ? quantity : 0m;
        }

        void RefreshComment()
        {
            var quantity = CurrentQuantity();
            quantityBox.IsEnabled = makeMax.IsChecked != true;
            commentBox.Text = piecewise.IsChecked == true
                ? BuildPiecewiseProductionLaunchCommentPreview(preview, quantity, ProductionLaunchCommentPrefix)
                : BuildProductionLaunchComment(preview, quantity);
            status.Text = quantity <= 0
                ? "Укажите количество больше 0."
                : quantity > preview.MaxQuantity
                    ? $"Нельзя запустить больше {FormatDecimal(preview.MaxQuantity)} деталей: не хватает материала."
                    : string.Empty;
        }

        makeMax.Checked += (_, _) =>
        {
            quantityBox.Text = FormatDecimal(preview.MaxQuantity);
            RefreshComment();
        };
        makeMax.Unchecked += (_, _) =>
        {
            quantityBox.Text = FormatDecimal(preview.DefaultQuantity);
            RefreshComment();
        };
        quantityBox.TextChanged += (_, _) => RefreshComment();
        piecewise.Checked += (_, _) => RefreshComment();
        piecewise.Unchecked += (_, _) => RefreshComment();
        cancelButton.Click += (_, _) => window.Close();
        launchButton.Click += (_, _) =>
        {
            var quantity = CurrentQuantity();
            if (quantity <= 0 || quantity > preview.MaxQuantity)
            {
                RefreshComment();
                return;
            }

            result = new ProductionLaunchDialogResult(decimal.Floor(quantity), commentBox.Text.Trim(), piecewise.IsChecked == true);
            window.DialogResult = true;
        };

        window.Content = root;
        RefreshComment();
        window.ShowDialog();
        return result;
    }

    public static IReadOnlyList<OneCProductionLaunchRequest> BuildProductionLaunchRequests(ProductionLaunchPreview preview, decimal quantity, string comment, bool piecewise)
    {
        var pieceCount = (int)Math.Max(0m, decimal.Floor(quantity));
        if (!piecewise || pieceCount <= 1)
        {
            return [BuildProductionLaunchRequest(preview, pieceCount, comment)];
        }

        var comments = BuildPiecewiseProductionLaunchComments(preview, pieceCount, comment);
        return comments
            .Select(commentLine => BuildProductionLaunchRequest(preview, 1m, commentLine))
            .ToList();
    }

    private static OneCProductionLaunchRequest BuildProductionLaunchRequest(ProductionLaunchPreview preview, decimal quantity, string comment)
    {
        var totalComponentQuantity = CalculateProductionLaunchMaterialQuantity(preview.ConsumptionUnit, preview.ConsumptionQuantity, quantity);
        return new OneCProductionLaunchRequest(
            preview.Ips,
            preview.Designation,
            preview.PartName,
            quantity,
            string.IsNullOrWhiteSpace(comment) ? BuildProductionLaunchComment(preview, quantity) : comment.Trim(),
            [new OneCProductionLaunchComponent(preview.ComponentOneCCode, preview.ComponentName, totalComponentQuantity, preview.ConsumptionUnit)]);
    }

    private static decimal CalculateProductionLaunchMaxQuantity(MeasurementUnit unit, decimal consumption, decimal materialStock)
    {
        if (consumption <= 0 || materialStock < consumption)
        {
            return 0m;
        }

        if (unit != MeasurementUnit.Meter)
        {
            return decimal.Floor(materialStock / consumption);
        }

        return decimal.Floor((materialStock + ProductionLaunchMeterCutWidth) / (consumption + ProductionLaunchMeterCutWidth));
    }

    private static decimal CalculateProductionLaunchMaterialQuantity(MeasurementUnit unit, decimal consumption, decimal quantity)
    {
        var pieceCount = Math.Max(0m, decimal.Floor(quantity));
        if (pieceCount <= 0 || consumption <= 0)
        {
            return 0m;
        }

        var baseQuantity = consumption * pieceCount;
        if (unit != MeasurementUnit.Meter)
        {
            return baseQuantity;
        }

        return baseQuantity + Math.Max(0m, pieceCount - 1m) * ProductionLaunchMeterCutWidth;
    }

    private static string BuildProductionLaunchComment(ProductionLaunchPreview preview, decimal quantity)
    {
        return ProductionLaunchCommentPrefix + " " + BuildCoverageComment(preview.ShortageRows, quantity);
    }

    private const string ProductionLaunchCommentPrefix = "Создано ПО Планирование";

    private static string BuildPiecewiseProductionLaunchCommentPreview(ProductionLaunchPreview preview, decimal quantity, string comment)
    {
        var pieceCount = (int)Math.Max(0m, decimal.Floor(quantity));
        return string.Join(Environment.NewLine, BuildPiecewiseProductionLaunchComments(preview, pieceCount, comment));
    }

    private static IReadOnlyList<string> BuildPiecewiseProductionLaunchComments(ProductionLaunchPreview preview, int quantity, string comment)
    {
        var enteredLines = SplitCommentLines(comment);
        if (enteredLines.Count > 1)
        {
            return BuildPiecewiseLaunchRows(preview.ShortageRows, quantity)
                .Select((row, index) => index < enteredLines.Count
                    ? enteredLines[index]
                    : BuildPiecewiseProductionLaunchComment(ProductionLaunchCommentPrefix, row))
                .ToList();
        }

        var prefix = enteredLines.Count == 0 ? ProductionLaunchCommentPrefix : enteredLines[0].Trim().TrimEnd('.', ';');
        return BuildPiecewiseLaunchRows(preview.ShortageRows, quantity)
            .Select(row => BuildPiecewiseProductionLaunchComment(prefix, row))
            .ToList();
    }

    private static IReadOnlyList<string> SplitCommentLines(string comment)
    {
        return comment
            .Split(["\r\n", "\n", "\r"], StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();
    }

    private static IReadOnlyList<ProductionLaunchShortageRow> BuildPiecewiseLaunchRows(IReadOnlyList<ProductionLaunchShortageRow> rows, int quantity)
    {
        var result = new List<ProductionLaunchShortageRow>(quantity);
        foreach (var row in rows)
        {
            var count = (int)Math.Max(0m, decimal.Floor(row.Quantity));
            for (var i = 0; i < count && result.Count < quantity; i++)
            {
                result.Add(row with { Quantity = 1m });
            }

            if (result.Count >= quantity)
            {
                break;
            }
        }

        while (result.Count < quantity)
        {
            result.Add(new ProductionLaunchShortageRow(1m, string.Empty, string.Empty, null));
        }

        return result;
    }

    private static string BuildPiecewiseProductionLaunchComment(string prefix, ProductionLaunchShortageRow row)
    {
        return string.IsNullOrWhiteSpace(row.MachineNumber)
            ? $"{prefix} 1 шт"
            : $"{prefix} {row.MachineNumber}; 1 шт";
    }

    private static string BuildCoverageComment(IReadOnlyList<ProductionLaunchShortageRow> rows, decimal quantity)
    {
        if (quantity <= 0)
        {
            return "закрытие дефицита: количество запуска не указано.";
        }

        var remaining = quantity;
        var parts = new List<string>();
        foreach (var row in rows)
        {
            if (remaining <= 0)
            {
                break;
            }

            var covered = Math.Min(row.Quantity, remaining);
            remaining -= covered;
            parts.Add($"{row.MachineNumber}; {FormatDecimal(covered)} шт");
        }

        return parts.Count == 0
            ? $"{FormatDecimal(quantity)} шт"
            : string.Join(" ", parts);
    }

    private async Task LoadFromDatabaseAsync()
    {
        await RemoveInvalidMskLibraryPartsAsync();
        demandIps.Clear();
        var latestDemandBatchId = await dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        if (latestDemandBatchId is not null)
        {
            foreach (var ips in await dbContext.DemandItems.AsNoTracking()
                .Where(x => x.DemandBatchId == latestDemandBatchId.Value && !string.IsNullOrWhiteSpace(x.Ips))
                .Select(x => x.Ips)
                .Distinct()
                .ToListAsync())
            {
                demandIps.Add(UiText.Clean(ips));
            }
        }

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
            .Where(x => IsValidMskLibraryIps(x.Ips))
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
            : $"Показано {Rows.Count} из {allRows.Count}; записей МСК: {recordsByIps.Count}.";
    }

    private void ApplyFilter()
    {
        var searchValue = UiText.Clean(Search).Trim();
        IEnumerable<MskLibraryRow> filtered = allRows;
        if (DemandOnly)
        {
            filtered = filtered.Where(x => demandIps.Contains(UiText.Clean(x.Ips)));
        }

        filtered = string.IsNullOrWhiteSpace(searchValue)
            ? filtered
            : filtered.Where(x => UiSearchText.ContainsAnyField(searchValue,
                x.Ips,
                x.Designation,
                x.Name,
                x.BlankType,
                x.BlankName,
                x.Material,
                x.OneCCode,
                x.FileName));

        Rows.Clear();
        foreach (var row in filtered.Take(1000))
        {
            Rows.Add(row);
        }

        SelectedRow = ResolvePreferredRow(searchValue);
        StatusText = DemandOnly
            ? $"Показано {Rows.Count} из {allRows.Count}; в потребности: {demandIps.Count}."
            : $"Показано {Rows.Count} из {allRows.Count}.";
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
            : Rows.FirstOrDefault(x => UiSearchText.EqualsNormalized(x.Ips, searchValue)) ?? Rows[0];
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

    public static bool IsValidMskLibraryIps(string? value)
    {
        var text = UiText.Clean(value).Trim();
        return (text.Length == 7 || text.Length == 11) &&
            text.Any(ch => ch != '0') &&
            text.All(char.IsDigit);
    }

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

    private static bool TryParseProductionLaunchQuantity(string value, out decimal quantity) =>
        decimal.TryParse((value ?? string.Empty).Replace(',', '.'), NumberStyles.Number, CultureInfo.InvariantCulture, out quantity);

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
        $"Есть МСК: {row.HasMsk}";

    private async Task TryOpenDrawingAsync(MskLibraryRow row, bool forceRetry)
    {
        if (string.IsNullOrWhiteSpace(row.Ips))
        {
            return;
        }

        var outputDirectory = GetDrawingCacheDirectory();
        var cachedDrawing = GetStableCachedDrawing(row, outputDirectory);
        if (IsReadablePdf(cachedDrawing))
        {
            ShowDrawing(row, cachedDrawing, fromCache: true);
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
            IsDrawingBusy = true;
            DrawingStatusText = $"Поиск PDF-чертежа IPS {row.Ips} в локальных папках...";
            var drawing = FindLocalDrawingPdf(row);
            if (drawing is not null)
            {
                var localStableDrawing = CopyToStableCache(row, drawing, outputDirectory);
                ShowDrawing(row, localStableDrawing, fromCache: false);
                return;
            }

            DrawingStatusText = $"PDF-чертеж IPS {row.Ips} локально не найден. Поиск через IPS Bridge...";
            foreach (var query in BuildDrawingQueries(row))
            {
                var bridgeDrawing = await drawingService.FindDrawingPdfAsync(query, outputDirectory, CancellationToken.None);
                if (bridgeDrawing is not null && IsReadablePdf(bridgeDrawing))
                {
                    var stableDrawing = CopyToStableCache(row, bridgeDrawing, outputDirectory);
                    ShowDrawing(row, stableDrawing, fromCache: false);
                    return;
                }
            }

            DrawingViewerSource = null;
            DrawingStatusText = $"PDF-чертеж IPS {row.Ips} не найден локально и через IPS Bridge.";
            StatusText = DrawingStatusText;
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Could not open local drawing for {Ips}", row.Ips);
            DrawingStatusText = $"PDF-чертеж IPS {row.Ips} не открыт: {ex.GetBaseException().Message}";
            StatusText = DrawingStatusText;
        }
        finally
        {
            IsDrawingBusy = false;
        }
    }

    internal static bool IsReadablePdf(FileInfo file)
    {
        if (!file.Exists || file.Length < 4)
        {
            return false;
        }

        using var stream = file.OpenRead();
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 &&
            header[0] == '%' &&
            header[1] == 'P' &&
            header[2] == 'D' &&
            header[3] == 'F';
    }

    internal static FileInfo GetStableCachedDrawing(MskLibraryRow row, DirectoryInfo outputDirectory) =>
        new(Path.Combine(outputDirectory.FullName, MakeStableDrawingFileName(row.Ips)));

    internal static FileInfo? FindLocalDrawingPdf(MskLibraryRow row)
    {
        foreach (var candidate in EnumerateLocalDrawingCandidates(row))
        {
            if (IsReadablePdf(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    internal static IEnumerable<string> BuildDrawingQueries(MskLibraryRow row)
    {
        var values = new List<string?>
        {
            row.Ips,
            row.Designation,
            row.Name
        };

        if (!string.IsNullOrWhiteSpace(row.Ips) && row.Ips.All(char.IsDigit) && row.Ips.Length < 11)
        {
            values.Insert(1, row.Ips.PadLeft(11, '0'));
        }

        return values
            .Select(x => UiText.Clean(x).Trim())
            .Where(x => x.Length >= 3)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<FileInfo> EnumerateLocalDrawingCandidates(MskLibraryRow row)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in EnumerateExactDrawingPaths(row))
        {
            if (seen.Add(path))
            {
                yield return new FileInfo(path);
            }
        }

        var tokens = new[]
            {
                row.Ips,
                row.Designation,
                row.Name,
                Path.GetFileNameWithoutExtension(row.FilePath)
            }
            .Select(NormalizeDrawingSearchText)
            .Where(x => x.Length >= 4)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var directory in EnumerateDrawingDirectories(row))
        {
            if (!directory.Exists)
            {
                continue;
            }

            IEnumerable<FileInfo> files;
            try
            {
                files = directory.EnumerateFiles("*.pdf", SearchOption.TopDirectoryOnly)
                    .Concat(directory.EnumerateFiles("*.PDF", SearchOption.TopDirectoryOnly));
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                if (!seen.Add(file.FullName))
                {
                    continue;
                }

                var normalizedName = NormalizeDrawingSearchText(Path.GetFileNameWithoutExtension(file.Name));
                if (tokens.Any(token => normalizedName.Contains(token, StringComparison.OrdinalIgnoreCase) ||
                    token.Contains(normalizedName, StringComparison.OrdinalIgnoreCase)))
                {
                    yield return file;
                }
            }
        }
    }

    private static IEnumerable<string> EnumerateExactDrawingPaths(MskLibraryRow row)
    {
        if (!string.IsNullOrWhiteSpace(row.FilePath))
        {
            var directory = Path.GetDirectoryName(row.FilePath);
            var fileName = Path.GetFileNameWithoutExtension(row.FilePath);
            if (!string.IsNullOrWhiteSpace(directory) && !string.IsNullOrWhiteSpace(fileName))
            {
                yield return Path.Combine(directory, $"{fileName}.pdf");
                yield return Path.Combine(directory, $"{fileName}.PDF");
            }
        }
    }

    private static IEnumerable<DirectoryInfo> EnumerateDrawingDirectories(MskLibraryRow row)
    {
        yield return new DirectoryInfo(Path.Combine(Directory.GetCurrentDirectory(), "Данные для работы", "Чертежи"));
        yield return new DirectoryInfo(Path.Combine(AppContext.BaseDirectory, "Данные для работы", "Чертежи"));
        yield return new DirectoryInfo(Path.Combine(AppContext.BaseDirectory, "drawings"));
        yield return new DirectoryInfo(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "Интеграция с сервисами",
            "factory_ai_assistant",
            "data",
            "drawings"));

        var configured = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_DIR");
        if (!string.IsNullOrWhiteSpace(configured))
        {
            foreach (var path in configured.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                yield return new DirectoryInfo(path);
            }
        }
    }

    private static string NormalizeDrawingSearchText(string? value)
    {
        var text = UiText.Clean(value);
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        return Regex.Replace(text.ToUpperInvariant(), @"[^0-9A-ZА-Я]+", string.Empty);
    }

    private static string MakeStableDrawingFileName(string ips)
    {
        var safeIps = Regex.Replace(ips, @"[^A-Za-zА-Яа-я0-9_. -]+", "_").Trim(' ', '_', '.');
        return string.IsNullOrWhiteSpace(safeIps) ? "IPS.pdf" : $"IPS_{safeIps}.pdf";
    }

    internal static FileInfo CopyToStableCache(MskLibraryRow row, FileInfo drawing, DirectoryInfo outputDirectory)
    {
        var target = GetStableCachedDrawing(row, outputDirectory);
        if (!string.Equals(drawing.FullName, target.FullName, StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(drawing.FullName, target.FullName, overwrite: true);
        }

        return target;
    }

    private void ShowDrawing(MskLibraryRow row, FileInfo drawing, bool fromCache)
    {
        DrawingViewerSource = new Uri(drawing.FullName);
        DrawingStatusText = fromCache
            ? $"PDF-чертеж IPS {row.Ips} открыт из кэша: {drawing.Name}"
            : $"PDF-чертеж IPS {row.Ips} загружен и открыт: {drawing.Name}";
        StatusText = DrawingStatusText;
    }

    internal static DirectoryInfo GetDrawingCacheDirectory()
    {
        var configured = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR");
        var path = string.IsNullOrWhiteSpace(configured)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BlankDemandPlanner",
                "IpsDrawings")
            : configured;
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

    internal static MskCsvDetail? ResolveMskDetail(MskRecord? record, IReadOnlyDictionary<string, MskCsvDetail> details)
    {
        if (record is null)
        {
            return null;
        }

        return details.TryGetValue(record.FileName, out var byFile)
            ? byFile
            : details.TryGetValue($"IPS:{record.Ips}", out var byIps) ? byIps : null;
    }

    private async Task TryAttachMskBlankAsync(Part part, MskCsvDetail? mskDetail)
    {
        if (mskDetail is null || string.IsNullOrWhiteSpace(mskDetail.OneCCode))
        {
            return;
        }

        var code = mskDetail.OneCCode.Trim();
        var normalizedCode = StockCodeNormalizer.NormalizeForComparison(code);
        var aliases = await dbContext.BlankAliases
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive)
            .ToListAsync();
        var alias = aliases.FirstOrDefault(x =>
            string.Equals(x.OneCCode, code, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(StockCodeNormalizer.NormalizeForComparison(x.OneCCode), normalizedCode, StringComparison.OrdinalIgnoreCase));
        if (alias?.CanonicalBlank is null)
        {
            return;
        }

        var quantity = decimal.TryParse(
            (mskDetail.ConsumptionQuantity ?? string.Empty).Trim().Replace('.', ','),
            NumberStyles.Number,
            CultureInfo.GetCultureInfo("ru-RU"),
            out var parsedQuantity)
            ? parsedQuantity
            : 1m;
        dbContext.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = alias.CanonicalBlank,
            ConsumptionQuantity = quantity > 0 ? quantity : 1m,
            ConsumptionUnit = ParseMskConsumptionUnit(mskDetail.UnitName, alias.CanonicalBlank.BaseUnit),
            BlankLeadTimeDays = 30,
            IsPrimary = true,
            IsActive = true,
            Source = "МСК; НСИ"
        });
    }

    private static MeasurementUnit ParseMskConsumptionUnit(string? value, MeasurementUnit fallback)
    {
        var text = UiText.Clean(value).ToLowerInvariant();
        if (text.Contains("пог", StringComparison.Ordinal) || text.Contains("м", StringComparison.Ordinal))
        {
            return MeasurementUnit.Meter;
        }

        if (text.Contains("шт", StringComparison.Ordinal))
        {
            return MeasurementUnit.Piece;
        }

        return fallback;
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

    private static bool IsArchivedLibrarySource(string? source) =>
        !string.IsNullOrWhiteSpace(source) && source.Contains("[ARCHIVED_LIBRARY]", StringComparison.OrdinalIgnoreCase);

    internal static IReadOnlyDictionary<string, MskCsvDetail> LoadMskDetailsFromReport()
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
        else
        {
            var role = await GetFirstLoginRoleAsync(http, settings, sessionId, tools, cancellationToken);
            if (role.RoleId is not null)
            {
                args["role_id"] = role.RoleId.Value;
            }
            else if (!string.IsNullOrWhiteSpace(role.RoleName))
            {
                args["role_name"] = role.RoleName;
            }
        }

        using var document = await CallToolAsync(http, settings, sessionId, tools["login"], args, cancellationToken);
        var payload = ExtractPayload(document);
        var session = FindString(payload, "session_id") ?? FindString(payload, "sessionId") ?? FindString(payload, "id");
        return string.IsNullOrWhiteSpace(session)
            ? throw new InvalidOperationException("IPS Bridge login не вернул session_id.")
            : session;
    }

    private static async Task<(int? RoleId, string? RoleName)> GetFirstLoginRoleAsync(
        HttpClient http,
        IpsBridgeSettings settings,
        string? sessionId,
        IReadOnlyDictionary<string, JsonElement> tools,
        CancellationToken cancellationToken)
    {
        var tool = tools.TryGetValue("list_login_roles", out var listLoginRoles)
            ? listLoginRoles
            : tools.TryGetValue("get_login_roles", out var getLoginRoles) ? getLoginRoles : default;
        if (tool.ValueKind == JsonValueKind.Undefined)
        {
            return (null, null);
        }

        try
        {
            using var document = await CallToolAsync(http, settings, sessionId, tool, new Dictionary<string, object?>
            {
                ["username"] = settings.Username
            }, cancellationToken);
            var payload = ExtractPayload(document);
            foreach (var nested in Walk(payload))
            {
                if (nested.ValueKind == JsonValueKind.Array)
                {
                    var first = nested.EnumerateArray().FirstOrDefault(x => x.ValueKind == JsonValueKind.Object);
                    if (first.ValueKind == JsonValueKind.Object)
                    {
                        return (
                            ReadInt(first, "RoleId") ?? ReadInt(first, "role_id") ?? ReadInt(first, "id"),
                            ReadString(first, "RoleName") ?? ReadString(first, "role_name") ?? ReadString(first, "name"));
                    }
                }
            }
        }
        catch
        {
            // Some bridge builds do not expose role discovery; login without role can still work there.
        }

        return (null, null);
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

public sealed partial class BlankSelectionViewModel(
    BlankDemandPlannerDbContext dbContext,
    IIpsDrawingService? ipsDrawingService = null,
    ILogger? logger = null) : ObservableObject
{
    private bool suppressPartSearchChanged;
    private bool suppressAutoFind;
    private bool isAutoFinding;
    private bool rerunAutoFind;
    private int autoFindVersion;
    private int partDetailVersion;
    private int partDrawingVersion;
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();

    public ObservableCollection<BlankSelectionRow> Rows { get; } = [];
    public ObservableCollection<PartWithoutBlankOption> PartSuggestions { get; } = [];
    public ObservableCollection<DisplayOption<MeasurementUnit>> UnitTypes { get; } = [..UiReferenceData.ConsumptionUnitTypes()];
    public ObservableCollection<DisplayOption<BlankType?>> BlankTypeFilters { get; } = [..UiReferenceData.SuggestedBlankTypeFilters()];

    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankType = UiReferenceData.SuggestedBlankTypeFilters()[0];
    [ObservableProperty] private string partSearch = string.Empty;
    [ObservableProperty] private bool isPartSuggestionsOpen;
    [ObservableProperty] private PartWithoutBlankOption? selectedPart;
    [ObservableProperty] private BlankSelectionRow? selectedBlankRow;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiReferenceData.ConsumptionUnitTypes()[0];
    [ObservableProperty] private string material = string.Empty;
    [ObservableProperty] private bool positiveStockOnly;
    [ObservableProperty] private string diameterText = string.Empty;
    [ObservableProperty] private string widthText = string.Empty;
    [ObservableProperty] private string heightText = string.Empty;
    [ObservableProperty] private string thicknessText = string.Empty;
    [ObservableProperty] private string wallThicknessText = string.Empty;
    [ObservableProperty] private string lengthText = string.Empty;
    [ObservableProperty] private string statusText = "Введите вид, материал или размеры детали для подбора заготовки.";
    [ObservableProperty] private string partDetailText = "Выберите деталь без заготовки для просмотра информации.";
    [ObservableProperty] private Uri? drawingViewerSource;
    [ObservableProperty] private string drawingStatusText = "PDF-чертеж не выбран.";
    [ObservableProperty] private bool isDrawingBusy;

    public void RefreshReferenceLists()
    {
        var selectedBlankTypeValue = SelectedBlankType.Value;
        var selectedUnitValue = SelectedUnit.Value;

        UiReferenceData.ReplaceOptions(UnitTypes, UiReferenceData.ConsumptionUnitTypes());
        UiReferenceData.ReplaceOptions(BlankTypeFilters, UiReferenceData.SuggestedBlankTypeFilters());

        SelectedBlankType = BlankTypeFilters.FirstOrDefault(x => x.Value == selectedBlankTypeValue) ?? BlankTypeFilters.First();
        SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == selectedUnitValue) ?? UnitTypes.First();
    }

    partial void OnSelectedBlankTypeChanged(DisplayOption<BlankType?> value) => QueueAutoFind();

    partial void OnMaterialChanged(string value) => QueueAutoFind();

    partial void OnPositiveStockOnlyChanged(bool value) => QueueAutoFind();

    partial void OnDiameterTextChanged(string value) => QueueAutoFind();

    partial void OnWidthTextChanged(string value) => QueueAutoFind();

    partial void OnHeightTextChanged(string value) => QueueAutoFind();

    partial void OnThicknessTextChanged(string value) => QueueAutoFind();

    partial void OnWallThicknessTextChanged(string value) => QueueAutoFind();

    partial void OnLengthTextChanged(string value) => QueueAutoFind();

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
        PartDetailText = BuildPartDetailText(value, null);
        _ = LoadSelectedPartDetailAsync(value, Interlocked.Increment(ref partDetailVersion));
        _ = TryOpenSelectedPartDrawingAsync(value, Interlocked.Increment(ref partDrawingVersion));
        StatusText = $"Выбрана деталь без заготовки: {value.Ips}. Подберите заготовку и нажмите \"Добавить в библиотеку\".";
    }

    partial void OnSelectedBlankRowChanged(BlankSelectionRow? value)
    {
        if (value is not null)
        {
            SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == value.BaseUnit) ?? SelectedUnit;
        }
    }

    private void QueueAutoFind()
    {
        if (suppressAutoFind)
        {
            return;
        }

        var version = Interlocked.Increment(ref autoFindVersion);
        _ = RunAutoFindAsync(version);
    }

    private async Task RunAutoFindAsync(int version)
    {
        var started = false;
        try
        {
            await Task.Delay(300);
            if (version != Volatile.Read(ref autoFindVersion))
            {
                return;
            }

            if (isAutoFinding)
            {
                rerunAutoFind = true;
                return;
            }

            isAutoFinding = true;
            started = true;
            await FindAsync();
        }
        finally
        {
            if (started)
            {
                isAutoFinding = false;
                if (rerunAutoFind || version != Volatile.Read(ref autoFindVersion))
                {
                    rerunAutoFind = false;
                    QueueAutoFind();
                }
            }
        }
    }

    public async Task LoadAsync() => await LoadPartSuggestionsAsync(openDropDown: false);

    [RelayCommand]
    public async Task LoadPartSuggestionsAsync() => await LoadPartSuggestionsAsync(openDropDown: false);

    private async Task LoadPartSuggestionsAsync(bool openDropDown)
    {
        var search = (PartSearch ?? string.Empty).Trim();
        var query = ActiveLibraryPartsWithoutBlankQuery();

        if (SelectedPart is not null && !string.Equals(search, SelectedPart.DisplayName, StringComparison.Ordinal))
        {
            SelectedPart = null;
        }

        if (SelectedPart is not null)
        {
            var stillVisible = PartSuggestions.Any(x => x.PartId == SelectedPart.PartId) ||
                await query.AnyAsync(x => x.Id == SelectedPart.PartId);
            if (!stillVisible)
            {
                SelectedPart = null;
            }
        }

        query = ActiveLibraryPartsWithoutBlankQuery();

        if (SelectedPart is not null && string.Equals(search, SelectedPart.DisplayName, StringComparison.Ordinal))
        {
            PartSuggestions.Clear();
            PartSuggestions.Add(SelectedPart);
            IsPartSuggestionsOpen = false;
            return;
        }

        if (!string.IsNullOrWhiteSpace(search))
        {
            var normalized = ExtractIps(search);
            var searchedParts = await query
                .OrderBy(x => x.Ips)
                .Take(5000)
                .ToListAsync();
            searchedParts = searchedParts
                .Where(x => UiSearchText.ContainsAnyField(search,
                    x.Ips,
                    normalized,
                    x.Designation,
                    x.Name))
                .Take(80)
                .ToList();

            PartSuggestions.Clear();
            foreach (var part in searchedParts.Select(x => new PartWithoutBlankOption(x.Id, x.Ips, x.Designation, x.Name)))
            {
                PartSuggestions.Add(part);
            }

            IsPartSuggestionsOpen = openDropDown && PartSuggestions.Count > 0;
            return;
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

    private IQueryable<Part> ActiveLibraryPartsWithoutBlankQuery() =>
        dbContext.Parts.AsNoTracking()
            .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
            .Where(x => !x.BlankMaps.Any(m => m.IsActive));

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
        blanks = blanks
            .Where(x => x.Aliases.Any(a => a.IsActive))
            .ToList();
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
        var stockItemsForSelection = latestSnapshotId is null
            ? []
            : await dbContext.StockItems.AsNoTracking()
                .Where(x => x.StockSnapshotId == latestSnapshotId.Value)
                .ToListAsync();
        var stockByCode = stockItemsForSelection
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.Sum(i => i.Quantity), StringComparer.OrdinalIgnoreCase);

        var selectedBlankId = SelectedBlankRow?.BlankId;
        var matches = blanks
            .Select(blank => TryBuildMatch(blank, request, stockByCode))
            .Where(x => x is not null)
            .Select(x => x!)
            .Where(x => !PositiveStockOnly || x.StockValue > 0)
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

        SelectedBlankRow = selectedBlankId is null
            ? Rows.FirstOrDefault()
            : Rows.FirstOrDefault(x => x.BlankId == selectedBlankId.Value) ?? Rows.FirstOrDefault();
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
        PartDetailText = "Выберите деталь без заготовки для просмотра информации.";
        Interlocked.Increment(ref partDetailVersion);
        DrawingViewerSource = null;
        DrawingStatusText = "PDF-чертеж не выбран.";
        Interlocked.Increment(ref partDrawingVersion);
        IsPartSuggestionsOpen = false;
        PartSearch = string.Empty;
        await LoadPartSuggestionsAsync();
    }

    [RelayCommand]
    private void Clear()
    {
        suppressAutoFind = true;
        try
        {
            SelectedBlankType = BlankTypeFilters[0];
            PartSearch = string.Empty;
            SelectedPart = null;
            PartDetailText = "Выберите деталь без заготовки для просмотра информации.";
            Interlocked.Increment(ref partDetailVersion);
            DrawingViewerSource = null;
            DrawingStatusText = "PDF-чертеж не выбран.";
            Interlocked.Increment(ref partDrawingVersion);
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
        finally
        {
            suppressAutoFind = false;
            Interlocked.Increment(ref autoFindVersion);
        }
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
            .Where(x => x.Source == null || !x.Source.Contains("[ARCHIVED_LIBRARY]"))
            .Where(x => !x.BlankMaps.Any(m => m.IsActive))
            .FirstOrDefaultAsync(x => x.Ips == ips);
    }

    private async Task LoadSelectedPartDetailAsync(PartWithoutBlankOption option, int version)
    {
        try
        {
            var record = await dbContext.MskRecords.AsNoTracking()
                .Where(x => x.Ips == option.Ips)
                .OrderByDescending(x => x.ImportedAt)
                .FirstOrDefaultAsync();
            var detail = MskViewModel.ResolveMskDetail(record, MskViewModel.LoadMskDetailsFromReport());
            if (version == Volatile.Read(ref partDetailVersion) && SelectedPart?.PartId == option.PartId)
            {
                PartDetailText = BuildPartDetailText(option, detail);
            }
        }
        catch
        {
            if (version == Volatile.Read(ref partDetailVersion) && SelectedPart?.PartId == option.PartId)
            {
                PartDetailText = BuildPartDetailText(option, null);
            }
        }
    }

    private async Task TryOpenSelectedPartDrawingAsync(PartWithoutBlankOption option, int version)
    {
        if (string.IsNullOrWhiteSpace(option.Ips))
        {
            return;
        }

        var row = new MskLibraryRow(
            option.Ips,
            UiText.Clean(option.Designation),
            UiText.Clean(option.PartName),
            string.Empty,
            "Да",
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty,
            "30");
        var outputDirectory = MskViewModel.GetDrawingCacheDirectory();
        var cachedDrawing = MskViewModel.GetStableCachedDrawing(row, outputDirectory);
        if (MskViewModel.IsReadablePdf(cachedDrawing))
        {
            ShowSelectedPartDrawing(option, cachedDrawing, fromCache: true, version);
            return;
        }

        try
        {
            IsDrawingBusy = true;
            DrawingStatusText = $"Поиск PDF-чертежа IPS {option.Ips}...";
            var drawing = MskViewModel.FindLocalDrawingPdf(row);
            if (drawing is not null)
            {
                var localStableDrawing = MskViewModel.CopyToStableCache(row, drawing, outputDirectory);
                ShowSelectedPartDrawing(option, localStableDrawing, fromCache: false, version);
                return;
            }

            foreach (var query in MskViewModel.BuildDrawingQueries(row))
            {
                var bridgeDrawing = await drawingService.FindDrawingPdfAsync(query, outputDirectory, CancellationToken.None);
                if (bridgeDrawing is not null && MskViewModel.IsReadablePdf(bridgeDrawing))
                {
                    var stableDrawing = MskViewModel.CopyToStableCache(row, bridgeDrawing, outputDirectory);
                    ShowSelectedPartDrawing(option, stableDrawing, fromCache: false, version);
                    return;
                }
            }

            if (version == Volatile.Read(ref partDrawingVersion) && SelectedPart?.PartId == option.PartId)
            {
                DrawingViewerSource = null;
                DrawingStatusText = $"PDF-чертеж IPS {option.Ips} не найден.";
            }
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex, "Could not open drawing for blank selection part {Ips}", option.Ips);
            if (version == Volatile.Read(ref partDrawingVersion) && SelectedPart?.PartId == option.PartId)
            {
                DrawingStatusText = $"PDF-чертеж IPS {option.Ips} не открыт: {ex.GetBaseException().Message}";
            }
        }
        finally
        {
            IsDrawingBusy = false;
        }
    }

    private void ShowSelectedPartDrawing(PartWithoutBlankOption option, FileInfo drawing, bool fromCache, int version)
    {
        if (version != Volatile.Read(ref partDrawingVersion) || SelectedPart?.PartId != option.PartId)
        {
            return;
        }

        DrawingViewerSource = new Uri(drawing.FullName);
        DrawingStatusText = fromCache
            ? $"PDF-чертеж IPS {option.Ips} открыт из кэша: {drawing.Name}"
            : $"PDF-чертеж IPS {option.Ips} открыт: {drawing.Name}";
    }

    private static string BuildPartDetailText(PartWithoutBlankOption option, MskCsvDetail? detail) =>
        $"IPS: {option.Ips}\n" +
        $"Обозначение: {UiText.Clean(option.Designation)}\n" +
        $"Наименование: {UiText.Clean(option.PartName)}\n" +
        $"Вид заготовки: {FirstNotEmpty(detail?.BlankType, "Не подобрана")}\n" +
        $"Заготовка: {FirstNotEmpty(detail?.BlankName, "Не подобрана")}\n" +
        $"Материал: {UiText.Clean(detail?.Material)}\n" +
        $"Код УТ: {UiText.Clean(detail?.OneCCode)}\n" +
        $"Норма расхода: {FirstNotEmpty(detail?.ConsumptionQuantity, "1")} {FirstNotEmpty(detail?.UnitName, "шт")}\n" +
        "Срок заготовки, дней: 30\n" +
        "Есть в библиотеке: Да\n" +
        $"Есть МСК: {(detail is null ? "Нет" : "Да")}";

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

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

        var alias = blank.Aliases.FirstOrDefault(x => x.IsActive);
        if (alias is null)
        {
            return null;
        }

        var stock = blank.Aliases
            .Where(x => !string.IsNullOrWhiteSpace(x.OneCCode) && stockByCode.TryGetValue(StockCodeNormalizer.NormalizeForComparison(x.OneCCode), out _))
            .Sum(x => stockByCode.TryGetValue(StockCodeNormalizer.NormalizeForComparison(x.OneCCode), out var quantity) ? quantity : 0m);
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
            stock,
            FormatDecimal(stock),
            FormatDecimal(allowance),
            allowance,
            details);
    }

    private static IEnumerable<DimensionComparison> BuildComparisons(CanonicalBlank blank, BlankSelectionRequest request)
    {
        if (blank.BlankType is BlankType.RoundBar or BlankType.BronzeBar or BlankType.PipeRound or BlankType.Forging)
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
    IFileDialogService fileDialogService,
    IIpsDrawingService? ipsDrawingService = null,
    ILogger? logger = null) : ObservableObject
{
    public ObservableCollection<CalculationMaterialRow> Rows { get; } = [];
    public ObservableCollection<LibraryBlankOption> BlankSuggestions { get; } = [];
    public ObservableCollection<DisplayOption<MeasurementUnit>> UnitTypes { get; } = [..UiReferenceData.ConsumptionUnitTypes()];
    public ObservableCollection<DisplayOption<BlankType?>> BlankTypeFilters { get; } = [..UiReferenceData.CalculationBlankTypeFilters()];
    private const string OneTimeBlankComment = "Разовая заготовка из расчета";
    private readonly IIpsDrawingService drawingService = ipsDrawingService ?? new IpsBridgeDrawingService();
    [ObservableProperty] private string statusText = string.Empty;
    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private CalculationMaterialRow? selectedRow;
    [ObservableProperty] private string blankSearch = string.Empty;
    [ObservableProperty] private DisplayOption<BlankType?> selectedBlankType = UiReferenceData.CalculationBlankTypeFilters()[0];
    [ObservableProperty] private string material = string.Empty;
    [ObservableProperty] private string diameterText = string.Empty;
    [ObservableProperty] private string widthText = string.Empty;
    [ObservableProperty] private string heightText = string.Empty;
    [ObservableProperty] private string thicknessText = string.Empty;
    [ObservableProperty] private string wallThicknessText = string.Empty;
    [ObservableProperty] private string lengthText = string.Empty;
    [ObservableProperty] private LibraryBlankOption? selectedBlank;
    [ObservableProperty] private string selectedOneCCode = string.Empty;
    [ObservableProperty] private string consumptionQuantityText = "1";
    [ObservableProperty] private DisplayOption<MeasurementUnit> selectedUnit = UiReferenceData.ConsumptionUnitTypes()[0];
    [ObservableProperty] private string assignmentStatus = string.Empty;
    private long? _lastRunId;
    private long? _lastBatchId;

    public void RefreshReferenceLists()
    {
        var selectedBlankTypeValue = SelectedBlankType.Value;
        var selectedUnitValue = SelectedUnit.Value;

        UiReferenceData.ReplaceOptions(UnitTypes, UiReferenceData.ConsumptionUnitTypes());
        UiReferenceData.ReplaceOptions(BlankTypeFilters, UiReferenceData.CalculationBlankTypeFilters());

        SelectedBlankType = BlankTypeFilters.FirstOrDefault(x => x.Value == selectedBlankTypeValue) ?? BlankTypeFilters.First();
        SelectedUnit = UnitTypes.FirstOrDefault(x => x.Value == selectedUnitValue) ?? UnitTypes.First();
    }

    partial void OnSearchChanged(string value) => _ = LoadLastRunAsync();
    partial void OnBlankSearchChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnSelectedBlankTypeChanged(DisplayOption<BlankType?> value) => _ = LoadBlankSuggestionsAsync();
    partial void OnMaterialChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnDiameterTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnWidthTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnHeightTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnThicknessTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnWallThicknessTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
    partial void OnLengthTextChanged(string value) => _ = LoadBlankSuggestionsAsync();
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
        StatusText = string.IsNullOrWhiteSpace(Search)
            ? $"Расчет материалов #{run.Id}, строк заявки: {Rows.Count}"
            : $"Расчет материалов #{run.Id}, показано строк: {Rows.Count}";
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
    private async Task OpenDrawingAsync(object? parameter)
    {
        var row = parameter as CalculationMaterialRow ?? SelectedRow;
        if (row is null)
        {
            AssignmentStatus = "Выберите строку расчета для открытия чертежа.";
            return;
        }

        SelectedRow = row;
        AssignmentStatus = await DrawingPdfOpener.OpenExternalAsync(
            new DrawingLookupRequest(row.Ips, null, row.Name, null),
            drawingService,
            "Расчет материалов",
            logger,
            CancellationToken.None);
    }

    [RelayCommand]
    private async Task LoadBlankSuggestionsAsync()
    {
        var query = dbContext.BlankAliases.AsNoTracking()
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive);

        var searchText = BlankSearch.Trim();
        var request = new BlankSelectionRequest(
            ParseCalculationNullable(DiameterText),
            ParseCalculationNullable(WidthText),
            ParseCalculationNullable(HeightText),
            ParseCalculationNullable(ThicknessText),
            ParseCalculationNullable(WallThicknessText),
            ParseCalculationNullable(LengthText));
        if (SelectedBlankType.Value is not null)
        {
            var type = SelectedBlankType.Value.Value;
            query = query.Where(x => x.CanonicalBlank != null && x.CanonicalBlank.BlankType == type);
        }

        var aliases = await query
            .OrderBy(x => x.SourceName)
            .Take(20000)
            .ToListAsync();
        if (!string.IsNullOrWhiteSpace(Material))
        {
            var materialKey = NormalizeCalculationMaterial(Material);
            aliases = aliases
                .Where(x => CalculationMaterialMatches(x, materialKey))
                .ToList();
        }

        if (request.HasAnySize)
        {
            aliases = aliases
                .Where(x => x.CanonicalBlank is not null && CalculationSizeMatches(x.CanonicalBlank, request))
                .OrderBy(x => x.CanonicalBlank is null ? decimal.MaxValue : CalculationAllowanceScore(x.CanonicalBlank, request))
                .ThenBy(x => x.SourceName)
                .ToList();
        }

        if (!string.IsNullOrWhiteSpace(searchText))
        {
            aliases = aliases
                .Where(x =>
                    ContainsSearch(x.CanonicalBlank?.CanonicalName, searchText) ||
                    ContainsSearch(x.CanonicalBlank?.Material, searchText) ||
                    ContainsSearch(x.OneCCode, searchText) ||
                    ContainsSearch(x.SourceName, searchText))
                .Take(50)
                .ToList();
        }
        else
        {
            aliases = aliases.Take(50).ToList();
        }

        var blankOptions = aliases
            .Where(x => x.CanonicalBlank is not null)
            .Select(x => new LibraryBlankOption(
                x.CanonicalBlankId,
                x.CanonicalBlank!.BlankType,
                x.CanonicalBlank.CanonicalName,
                x.SourceName,
                x.CanonicalBlank.Material,
                x.CanonicalBlank.BaseUnit,
                x.OneCCode))
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

    private static bool CalculationSizeMatches(CanonicalBlank blank, BlankSelectionRequest request)
    {
        var comparisons = BuildCalculationComparisons(blank, request).ToArray();
        return !request.HasAnySize || (comparisons.Length > 0 && comparisons.All(x => x.BlankValue >= x.RequiredValue));
    }

    private static decimal CalculationAllowanceScore(CanonicalBlank blank, BlankSelectionRequest request) =>
        BuildCalculationComparisons(blank, request).Sum(x => x.BlankValue - x.RequiredValue);

    private static IEnumerable<DimensionComparison> BuildCalculationComparisons(CanonicalBlank blank, BlankSelectionRequest request)
    {
        if (blank.BlankType is BlankType.RoundBar or BlankType.BronzeBar or BlankType.PipeRound or BlankType.Forging)
        {
            var requiredDiameter = request.Diameter ?? MaxCalculation(request.Width, request.Height, request.Thickness);
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
            var requiredWidth = request.Width ?? request.Diameter ?? MaxCalculation(request.Height, request.Thickness);
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

    private static decimal? MaxCalculation(params decimal?[] values)
    {
        var present = values.Where(x => x is not null).Select(x => x!.Value).ToArray();
        return present.Length == 0 ? null : present.Max();
    }

    private static decimal? ParseCalculationNullable(string? value) =>
        decimal.TryParse((value ?? string.Empty).Trim().Replace('.', ','), NumberStyles.Number, CultureInfo.GetCultureInfo("ru-RU"), out var result)
            ? result
            : null;

    private static string NormalizeCalculationMaterial(string value) => UiSearchText.Normalize(value).Replace(" ", string.Empty, StringComparison.Ordinal);

    private static bool CalculationMaterialMatches(BlankAlias alias, string materialKey)
    {
        if (string.IsNullOrWhiteSpace(materialKey))
        {
            return true;
        }

        var blank = alias.CanonicalBlank;
        return NormalizeCalculationMaterial(blank?.Material ?? string.Empty).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
            NormalizeCalculationMaterial(blank?.CanonicalName ?? string.Empty).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
            NormalizeCalculationMaterial(alias.SourceName).Contains(materialKey, StringComparison.OrdinalIgnoreCase) ||
            NormalizeCalculationMaterial(alias.NormalizedSourceName).Contains(materialKey, StringComparison.OrdinalIgnoreCase);
    }

    private async Task LoadMaterialRowsAsync(CalculationRun run)
    {
        var stockRemaining = run.Items
            .Where(x => x.CanonicalBlankId is not null && x.Status == CalculationStatus.Ok)
            .GroupBy(x => (BlankId: x.CanonicalBlankId!.Value, x.Unit))
            .ToDictionary(x => x.Key, x => x.Sum(i => i.TotalStock));

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

        var rows = new List<CalculationMaterialRow>();
        var number = 1;
        var meterCutSourceCounts = new Dictionary<(long BlankId, MeasurementUnit Unit), int>();
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
                rows.Add(new CalculationMaterialRow(
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
                rows.Add(new CalculationMaterialRow(
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
                var key = (map.CanonicalBlankId, map.ConsumptionUnit);
                meterCutSourceCounts.TryGetValue(key, out var sourceCount);
                var requiredWithCut = required + CalculateMeterCutAllowance(map.ConsumptionUnit, effectiveDemandQuantity, sourceCount);
                meterCutSourceCounts[key] = sourceCount + 1;
                var purchaseQuantity = AllocateUncoveredQuantity(stockRemaining, map.CanonicalBlankId, map.ConsumptionUnit, requiredWithCut);
                if (purchaseQuantity <= 0)
                {
                    continue;
                }

                var displayQuantity = Math.Min(required, purchaseQuantity);
                rows.Add(new CalculationMaterialRow(
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
                    FormatDecimal(displayQuantity),
                    FormatBlankDemandDate(demand.DemandDate, map.BlankLeadTimeDays),
                    false,
                    demand.Id));
            }
        }

        var searchValue = Search.Trim();
        var visibleRows = string.IsNullOrWhiteSpace(searchValue)
            ? rows
            : rows.Where(x => CalculationRowMatchesSearch(x, searchValue)).ToList();

        Rows.Clear();
        foreach (var row in visibleRows)
        {
            Rows.Add(row);
        }
    }

    private static bool CalculationRowMatchesSearch(CalculationMaterialRow row, string searchValue) =>
        UiSearchText.ContainsAnyField(searchValue,
            row.ProductionSystem,
            row.MachineNumber,
            row.Ips,
            row.Name,
            row.OneCCode,
            row.BlankType,
            row.Nomenclature,
            row.UnitName,
            row.DemandDate);

    private static decimal AllocateUncoveredQuantity(IDictionary<(long BlankId, MeasurementUnit Unit), decimal> stockRemaining, long blankId, MeasurementUnit unit, decimal required)
    {
        if (required <= 0)
        {
            return 0m;
        }

        if (!stockRemaining.TryGetValue((blankId, unit), out var stock) || stock <= 0)
        {
            return required;
        }

        var covered = Math.Min(required, stock);
        stockRemaining[(blankId, unit)] = stock - covered;
        return required - covered;
    }

    private static decimal CalculateMeterCutAllowance(MeasurementUnit unit, decimal effectiveDemandQuantity, int existingSourceCount)
    {
        if (unit != MeasurementUnit.Meter || effectiveDemandQuantity <= 0)
        {
            return 0m;
        }

        var pieceCount = (int)Math.Ceiling(effectiveDemandQuantity);
        return (Math.Max(0, pieceCount - 1) + (existingSourceCount > 0 ? 1 : 0)) * 0.005m;
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
        UiSearchText.Contains(value, searchText);
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

public sealed partial class SettingsViewModel(BlankDemandPlannerDbContext dbContext, IAppAuthService authService) : ObservableObject
{
    private bool settingsLoaded;
    private bool isLoadingSettings;

    public IReadOnlyList<string> FontFamilies { get; } = ["Segoe UI", "Arial", "Calibri", "Tahoma", "Times New Roman"];
    public IReadOnlyList<double> FontSizes { get; } = [12, 13, 14, 15, 16, 18, 20];
    public IReadOnlyList<DisplayOption<HorizontalAlignment>> ToolbarPlacements { get; } =
    [
        new(HorizontalAlignment.Left, "Слева"),
        new(HorizontalAlignment.Center, "По центру"),
        new(HorizontalAlignment.Right, "Справа")
    ];
    public IReadOnlyList<DisplayOption<Dock>> NavigationPlacements { get; } =
    [
        new(Dock.Left, "Слева"),
        new(Dock.Right, "Справа"),
        new(Dock.Top, "Сверху"),
        new(Dock.Bottom, "Снизу")
    ];

    [ObservableProperty] private string selectedFontFamily = "Segoe UI";
    [ObservableProperty] private double selectedFontSize = 13;
    [ObservableProperty] private DisplayOption<HorizontalAlignment> selectedToolbarPlacement = new(HorizontalAlignment.Left, "Слева");
    [ObservableProperty] private HorizontalAlignment toolbarHorizontalAlignment = HorizontalAlignment.Left;
    [ObservableProperty] private DisplayOption<Dock> selectedNavigationPlacement = new(Dock.Left, "Слева");
    [ObservableProperty] private Dock navigationDock = Dock.Left;
    [ObservableProperty] private System.Windows.Controls.Orientation navigationOrientation = System.Windows.Controls.Orientation.Vertical;
    [ObservableProperty] private double navigationPanelWidth = 168;
    [ObservableProperty] private double navigationPanelHeight = double.NaN;
    [ObservableProperty] private Thickness navigationListMargin = new(0, 22, 0, 0);
    [ObservableProperty] private string iconPath = string.Empty;
    [ObservableProperty] private string photoPath = string.Empty;
    [ObservableProperty] private AppUserAdminRow? selectedUser;
    [ObservableProperty] private string userName = string.Empty;
    [ObservableProperty] private string userDisplayName = string.Empty;
    [ObservableProperty] private string generatedPassword = string.Empty;
    [ObservableProperty] private bool userIsAdmin;
    [ObservableProperty] private bool userIsActive = true;
    [ObservableProperty] private string currentPassword = string.Empty;
    [ObservableProperty] private string newPassword = string.Empty;
    [ObservableProperty] private string repeatPassword = string.Empty;
    public bool IsAdmin => authService.CurrentUser?.IsAdmin == true;
    public string SignedInUserText => authService.CurrentUser is null
        ? "Пользователь не определен"
        : $"{authService.CurrentUser.DisplayName} ({authService.CurrentUser.UserName})";
    public ObservableCollection<AppUserAdminRow> Users { get; } = [];
    public ObservableCollection<PermissionEditRow> PermissionRows { get; } = [];
    [ObservableProperty] private string statusText = "Настройки готовы";

    [RelayCommand]
    public async Task LoadAsync()
    {
        if (!settingsLoaded)
        {
            isLoadingSettings = true;
            try
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
                var navigationPlacement = await ReadSettingAsync("UI.NavigationPlacement", "Left");
                SelectedNavigationPlacement = NavigationPlacements.FirstOrDefault(x => string.Equals(x.Value.ToString(), navigationPlacement, StringComparison.OrdinalIgnoreCase))
                    ?? NavigationPlacements[0];
                ApplyNavigationPlacement();
                IconPath = await ReadSettingAsync("UI.IconPath", string.Empty);
                PhotoPath = await ReadSettingAsync("UI.PhotoPath", string.Empty);
                ApplyToWindow();
                settingsLoaded = true;
            }
            finally
            {
                isLoadingSettings = false;
            }
        }

        await LoadUsersAsync();
    }

    private async Task LoadUsersAsync()
    {
        PermissionRows.Clear();
        foreach (var definition in authService.PageDefinitions)
        {
            PermissionRows.Add(new PermissionEditRow(definition.PageKey, definition.Title, definition.PageKey == "Settings", false));
        }

        Users.Clear();
        if (!IsAdmin)
        {
            return;
        }

        foreach (var user in await authService.GetUsersAsync())
        {
            Users.Add(user);
        }
    }

    [RelayCommand]
    private async Task ApplyAsync()
    {
        ToolbarHorizontalAlignment = SelectedToolbarPlacement.Value;
        ApplyNavigationPlacement();
        ApplyToWindow();
        await WriteSettingAsync("UI.FontFamily", SelectedFontFamily);
        await WriteSettingAsync("UI.FontSize", SelectedFontSize.ToString(CultureInfo.InvariantCulture));
        await WriteSettingAsync("UI.ToolbarPlacement", SelectedToolbarPlacement.Value.ToString());
        await WriteSettingAsync("UI.NavigationPlacement", SelectedNavigationPlacement.Value.ToString());
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
        SelectedNavigationPlacement = NavigationPlacements[0];
        IconPath = string.Empty;
        PhotoPath = string.Empty;
        await ApplyAsync();
    }

    [RelayCommand]
    private void GeneratePassword()
    {
        GeneratedPassword = authService.GeneratePassword();
        StatusText = "Пароль сгенерирован. В базе хранится только bcrypt-хэш.";
    }

    [RelayCommand]
    private void NewUser()
    {
        SelectedUser = null;
        UserName = string.Empty;
        UserDisplayName = string.Empty;
        GeneratedPassword = authService.GeneratePassword();
        UserIsAdmin = false;
        UserIsActive = true;
        foreach (var row in PermissionRows)
        {
            row.CanRead = row.PageKey == "Settings";
            row.CanEdit = false;
        }
    }

    [RelayCommand]
    private async Task SaveUserAsync()
    {
        if (!IsAdmin)
        {
            StatusText = "Недостаточно прав.";
            return;
        }

        var savedUserName = UserName.Trim();
        var result = await authService.SaveUserAsync(new AppUserEditRequest(
            SelectedUser?.Id,
            UserName,
            UserDisplayName,
            UserIsAdmin,
            UserIsActive,
            string.IsNullOrWhiteSpace(GeneratedPassword) ? null : GeneratedPassword,
            PermissionRows.Select(x => new AppUserPermissionRow(x.PageKey, x.CanRead, x.CanEdit)).ToList()));
        StatusText = result.Message;
        await LoadUsersAsync();
        if (result.IsSuccess)
        {
            SelectedUser = Users.FirstOrDefault(x => string.Equals(x.UserName, savedUserName, StringComparison.OrdinalIgnoreCase));
        }
    }

    [RelayCommand]
    private async Task ResetUserPasswordAsync()
    {
        if (!IsAdmin || SelectedUser is null)
        {
            StatusText = "Выберите пользователя.";
            return;
        }

        if (string.IsNullOrWhiteSpace(GeneratedPassword))
        {
            GeneratedPassword = authService.GeneratePassword();
        }

        var result = await authService.ResetPasswordAsync(SelectedUser.Id, GeneratedPassword);
        StatusText = result.Message;
        await LoadUsersAsync();
    }

    [RelayCommand]
    private async Task DeleteUserAsync()
    {
        if (!IsAdmin || SelectedUser is null)
        {
            StatusText = "Выберите пользователя.";
            return;
        }

        var userName = SelectedUser.UserName;
        var result = await authService.DeleteUserAsync(SelectedUser.Id);
        StatusText = result.Message;
        await LoadUsersAsync();
        if (result.IsSuccess)
        {
            SelectedUser = null;
            NewUser();
            StatusText = $"Пользователь {userName} удален из активного списка.";
        }
    }

    [RelayCommand]
    private async Task ChangeOwnPasswordAsync()
    {
        if (!string.Equals(NewPassword, RepeatPassword, StringComparison.Ordinal))
        {
            StatusText = "Новый пароль и повтор не совпадают.";
            return;
        }

        var result = await authService.ChangeOwnPasswordAsync(CurrentPassword, NewPassword);
        StatusText = result.Message;
        if (result.IsSuccess)
        {
            CurrentPassword = string.Empty;
            NewPassword = string.Empty;
            RepeatPassword = string.Empty;
        }
    }

    partial void OnSelectedUserChanged(AppUserAdminRow? value)
    {
        if (value is null)
        {
            return;
        }

        UserName = value.UserName;
        UserDisplayName = value.DisplayName;
        UserIsAdmin = value.IsAdmin;
        UserIsActive = value.IsActive;
        GeneratedPassword = string.Empty;
        foreach (var row in PermissionRows)
        {
            var permission = value.Permissions.FirstOrDefault(x => x.PageKey == row.PageKey);
            row.CanRead = value.IsAdmin || permission?.CanRead == true || row.PageKey == "Settings";
            row.CanEdit = value.IsAdmin || permission?.CanEdit == true;
        }
    }

    partial void OnSelectedFontFamilyChanged(string value)
    {
        if (!isLoadingSettings)
        {
            ApplyToWindow();
        }
    }

    partial void OnSelectedFontSizeChanged(double value)
    {
        if (!isLoadingSettings)
        {
            ApplyToWindow();
        }
    }

    partial void OnSelectedToolbarPlacementChanged(DisplayOption<HorizontalAlignment> value) => ToolbarHorizontalAlignment = value.Value;
    partial void OnSelectedNavigationPlacementChanged(DisplayOption<Dock> value) => ApplyNavigationPlacement();

    private void ApplyNavigationPlacement()
    {
        NavigationDock = SelectedNavigationPlacement.Value;
        var isHorizontal = NavigationDock is Dock.Top or Dock.Bottom;
        NavigationOrientation = isHorizontal ? System.Windows.Controls.Orientation.Horizontal : System.Windows.Controls.Orientation.Vertical;
        NavigationPanelWidth = isHorizontal ? double.NaN : 168;
        NavigationPanelHeight = isHorizontal ? 48 : double.NaN;
        NavigationListMargin = isHorizontal ? new Thickness(10, 4, 10, 4) : new Thickness(0, 22, 0, 0);
    }

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

public sealed partial class PermissionEditRow(string pageKey, string title, bool canRead, bool canEdit) : ObservableObject
{
    public string PageKey { get; } = pageKey;
    public string Title { get; } = title;
    [ObservableProperty] private bool canRead = canRead;
    [ObservableProperty] private bool canEdit = canEdit;
}

public sealed partial class DeveloperModeViewModel(
    BlankDemandPlannerDbContext dbContext,
    IAppAuthService authService,
    Func<Task>? afterSave = null) : ObservableObject
{
    public ObservableCollection<DeveloperOptionRow> UnitRows { get; } = [];
    public ObservableCollection<DeveloperOptionRow> ServiceRows { get; } = [];
    public ObservableCollection<DeveloperOptionRow> SupplyRows { get; } = [];
    public ObservableCollection<DeveloperOptionRow> BlankTypeRows { get; } = [];
    [ObservableProperty] private string statusText = "Режим разработчика доступен только администратору.";
    public bool IsAdmin => authService.CurrentUser?.IsAdmin == true;

    [RelayCommand]
    public async Task LoadAsync()
    {
        if (!IsAdmin)
        {
            StatusText = "Недостаточно прав.";
            return;
        }

        await UiReferenceData.LoadAsync(dbContext);
        ReplaceRows(UnitRows, UiReferenceData.UnitEditorRows());
        ReplaceRows(ServiceRows, UiReferenceData.ServiceEditorRows());
        ReplaceRows(SupplyRows, UiReferenceData.SupplyEditorRows());
        ReplaceRows(BlankTypeRows, UiReferenceData.BlankTypeEditorRows());
        StatusText = "Справочники загружены.";
    }

    [RelayCommand]
    private void AddUnit() => UnitRows.Add(new DeveloperOptionRow("Meter", "Новая ед.", true));

    [RelayCommand]
    private void AddService() => ServiceRows.Add(new DeveloperOptionRow("Custom", "Новая услуга", true));

    [RelayCommand]
    private void AddSupply() => SupplyRows.Add(new DeveloperOptionRow("Custom", "Новое условие", true));

    [RelayCommand]
    private void AddBlankType() => BlankTypeRows.Add(new DeveloperOptionRow(nameof(BlankType.CustomBlank), "Новый вид", true));

    [RelayCommand]
    private void RemoveUnit(DeveloperOptionRow? row) => RemoveRow(UnitRows, row);

    [RelayCommand]
    private void RemoveService(DeveloperOptionRow? row) => RemoveRow(ServiceRows, row);

    [RelayCommand]
    private void RemoveSupply(DeveloperOptionRow? row) => RemoveRow(SupplyRows, row);

    [RelayCommand]
    private void RemoveBlankType(DeveloperOptionRow? row) => RemoveRow(BlankTypeRows, row);

    [RelayCommand]
    private async Task SaveAsync()
    {
        if (!IsAdmin)
        {
            StatusText = "Недостаточно прав.";
            return;
        }

        await UiReferenceData.SaveAsync(dbContext, new DeveloperReferenceLists(
            UnitRows.Select(x => x.ToDto()).ToList(),
            ServiceRows.Select(x => x.ToDto()).ToList(),
            SupplyRows.Select(x => x.ToDto()).ToList(),
            BlankTypeRows.Select(x => x.ToDto()).ToList()));
        if (afterSave is not null)
        {
            await afterSave();
        }

        StatusText = "Справочники сохранены и применены к открытым вкладкам.";
    }

    private static void ReplaceRows(ObservableCollection<DeveloperOptionRow> target, IEnumerable<DeveloperOptionRow> source)
    {
        target.Clear();
        foreach (var row in source)
        {
            target.Add(row);
        }
    }

    private static void RemoveRow(ObservableCollection<DeveloperOptionRow> rows, DeveloperOptionRow? row)
    {
        if (row is not null)
        {
            rows.Remove(row);
        }
    }
}

public sealed partial class DeveloperOptionRow(string code, string displayName, bool isActive) : ObservableObject
{
    [ObservableProperty] private string code = code;
    [ObservableProperty] private string displayName = displayName;
    [ObservableProperty] private bool isActive = isActive;

    public DeveloperReferenceOption ToDto() => new(Code.Trim(), DisplayName.Trim(), IsActive);
}

public sealed record DeveloperReferenceLists(
    List<DeveloperReferenceOption> Units,
    List<DeveloperReferenceOption> Services,
    List<DeveloperReferenceOption> SupplyConditions,
    List<DeveloperReferenceOption> BlankTypes);

public sealed record DeveloperReferenceOption(string Code, string DisplayName, bool IsActive);

public static class UiReferenceData
{
    private const string SettingKey = "Developer.ReferenceLists";
    private static DeveloperReferenceLists current = CreateDefaults();

    public static async Task LoadAsync(BlankDemandPlannerDbContext dbContext)
    {
        var json = await dbContext.Settings
            .AsNoTracking()
            .Where(x => x.Key == SettingKey)
            .Select(x => x.Value)
            .FirstOrDefaultAsync();
        current = string.IsNullOrWhiteSpace(json)
            ? CreateDefaults()
            : JsonSerializer.Deserialize<DeveloperReferenceLists>(json) ?? CreateDefaults();
    }

    public static async Task SaveAsync(BlankDemandPlannerDbContext dbContext, DeveloperReferenceLists lists)
    {
        current = Normalize(lists);
        var json = JsonSerializer.Serialize(current);
        var setting = await dbContext.Settings.FirstOrDefaultAsync(x => x.Key == SettingKey);
        if (setting is null)
        {
            dbContext.Settings.Add(new AppSetting { Key = SettingKey, Value = json });
        }
        else
        {
            setting.Value = json;
        }

        await dbContext.SaveChangesAsync();
    }

    public static IReadOnlyList<DisplayOption<MeasurementUnit>> UnitTypes() =>
        current.Units
            .Where(x => x.IsActive)
            .Select(x => new DisplayOption<MeasurementUnit>(ParseUnitCode(x.Code), CleanName(x.DisplayName, x.Code)))
            .ToList();

    public static IReadOnlyList<DisplayOption<MeasurementUnit>> ConsumptionUnitTypes() =>
        UnitTypes()
            .Where(x => x.Value != MeasurementUnit.Kilogram)
            .ToList();

    public static IReadOnlyList<DisplayOption<MeasurementUnit?>> UnitFilters()
    {
        var seen = new HashSet<MeasurementUnit>();
        return
        [
            new((MeasurementUnit?)null, "Все ед."),
            ..UnitTypes()
                .Where(x => seen.Add(x.Value))
                .Select(x => new DisplayOption<MeasurementUnit?>(x.Value, x.Value == MeasurementUnit.Meter ? "пог. м" : x.DisplayName))
        ];
    }

    public static IReadOnlyList<DisplayOption<MeasurementUnit?>> ConsumptionUnitFilters() =>
    [
        new((MeasurementUnit?)null, "Все ед."),
        ..UnitFilters().Where(x => x.Value is MeasurementUnit.Piece or MeasurementUnit.Meter)
    ];

    public static IReadOnlyList<DisplayOption<BlankType>> BlankTypes() =>
        current.BlankTypes
            .Where(x => x.IsActive)
            .Select(x => new DisplayOption<BlankType>(ParseBlankTypeCode(x.Code), CleanName(x.DisplayName, x.Code)))
            .ToList();

    public static IReadOnlyList<DisplayOption<BlankType?>> BlankTypeFilters(bool includeUnknown)
    {
        var types = BlankTypes()
            .Where(x => includeUnknown || x.Value is not BlankType.Unknown and not BlankType.CustomBlank)
            .Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName));
        return [new((BlankType?)null, "Все виды"), ..types];
    }

    public static IReadOnlyList<DisplayOption<BlankType?>> SuggestedBlankTypeFilters() =>
    [
        new((BlankType?)null, "Все подходящие"),
        ..BlankTypes()
            .Where(x => x.Value is not BlankType.Unknown and not BlankType.CustomBlank and not BlankType.Purchased)
            .Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName))
    ];

    public static IReadOnlyList<DisplayOption<BlankType?>> CalculationBlankTypeFilters() =>
    [
        new((BlankType?)null, "Все подходящие"),
        ..BlankTypes()
            .Where(x => x.Value is not BlankType.Unknown and not BlankType.CustomBlank)
            .Select(x => new DisplayOption<BlankType?>(x.Value, x.DisplayName))
    ];

    public static IReadOnlyList<string> ServiceTypes() =>
        current.Services
            .Where(x => x.IsActive && !string.IsNullOrWhiteSpace(x.DisplayName))
            .Select(x => x.DisplayName.Trim())
            .DefaultIfEmpty("Хим окс")
            .ToList();

    public static string ServiceName(string code, string fallback) => LookupName(current.Services, code, fallback);
    public static string SupplyConditionName(string code, string fallback) => LookupName(current.SupplyConditions, code, fallback);
    public static bool IsServiceActive(string code) => IsActive(current.Services, code);
    public static bool IsSupplyConditionActive(string code) => IsActive(current.SupplyConditions, code);

    public static IReadOnlyList<DeveloperOptionRow> UnitEditorRows() => current.Units.Select(ToRow).ToList();
    public static IReadOnlyList<DeveloperOptionRow> ServiceEditorRows() => current.Services.Select(ToRow).ToList();
    public static IReadOnlyList<DeveloperOptionRow> SupplyEditorRows() => current.SupplyConditions.Select(ToRow).ToList();
    public static IReadOnlyList<DeveloperOptionRow> BlankTypeEditorRows() => current.BlankTypes.Select(ToRow).ToList();

    public static string DisplayUnit(MeasurementUnit unit)
    {
        var options = UnitTypes();
        if (unit == MeasurementUnit.Meter)
        {
            return options.FirstOrDefault(x => x.Value == unit && !UiText.IsMillimeterOption(x))?.DisplayName ?? "пог. м";
        }

        return options.FirstOrDefault(x => x.Value == unit)?.DisplayName ?? unit.ToString();
    }

    public static string DisplayBlankType(BlankType type) =>
        type == BlankType.CustomBlank ? "Прочее" : BlankTypes().FirstOrDefault(x => x.Value == type)?.DisplayName ?? UiText.DisplayHiddenBlankType(type);

    public static void ReplaceOptions<T>(ObservableCollection<DisplayOption<T>> target, IEnumerable<DisplayOption<T>> source)
    {
        target.Clear();
        foreach (var option in source)
        {
            target.Add(option);
        }
    }

    private static DeveloperOptionRow ToRow(DeveloperReferenceOption option) => new(option.Code, option.DisplayName, option.IsActive);

    private static string LookupName(IEnumerable<DeveloperReferenceOption> source, string code, string fallback) =>
        source.FirstOrDefault(x => string.Equals(x.Code, code, StringComparison.OrdinalIgnoreCase))?.DisplayName.Trim()
        ?? fallback;

    private static bool IsActive(IEnumerable<DeveloperReferenceOption> source, string code) =>
        source.FirstOrDefault(x => string.Equals(x.Code, code, StringComparison.OrdinalIgnoreCase))?.IsActive ?? true;

    private static DeveloperReferenceLists Normalize(DeveloperReferenceLists lists) => new(
        NormalizeList(lists.Units, CreateDefaults().Units),
        NormalizeList(lists.Services, CreateDefaults().Services),
        NormalizeList(lists.SupplyConditions, CreateDefaults().SupplyConditions),
        NormalizeList(lists.BlankTypes, CreateDefaults().BlankTypes));

    private static List<DeveloperReferenceOption> NormalizeList(IEnumerable<DeveloperReferenceOption> rows, IEnumerable<DeveloperReferenceOption> fallback)
    {
        var result = rows
            .Select(x => new DeveloperReferenceOption(x.Code.Trim(), x.DisplayName.Trim(), x.IsActive))
            .Where(x => !string.IsNullOrWhiteSpace(x.Code) && !string.IsNullOrWhiteSpace(x.DisplayName))
            .ToList();
        return result.Count == 0 ? fallback.ToList() : result;
    }

    private static DeveloperReferenceLists CreateDefaults() => new(
        [
            new(nameof(MeasurementUnit.Piece), "шт", true),
            new(nameof(MeasurementUnit.Meter), "мм", true),
            new(nameof(MeasurementUnit.Meter), "пог. м", true),
            new(nameof(MeasurementUnit.Kilogram), "кг", true)
        ],
        [
            new("ChemicalOxidation", "Хим окс", true),
            new("Nitriding", "Азотирование", true),
            new("Keyway", "Шпоночный паз", true),
            new("HeatTreatment", "Термообработка", true)
        ],
        [
            new("HeatTreatment", "ТО", true),
            new("LaserCutting", "Лазерная резка", true)
        ],
        [
            new(nameof(BlankType.RoundBar), "Круг", true),
            new(nameof(BlankType.SquareBar), "Квадрат", true),
            new(nameof(BlankType.HexBar), "Шестигранник", true),
            new(nameof(BlankType.Sheet), "Лист", true),
            new(nameof(BlankType.Plate), "Плита", true),
            new(nameof(BlankType.PipeRound), "Труба профильная круглая", true),
            new(nameof(BlankType.PipeRectangular), "Труба профильная прямоугольная", true),
            new(nameof(BlankType.Angle), "Уголок", true),
            new(nameof(BlankType.WeldingElement), "Сварное изделие", true),
            new(nameof(BlankType.Purchased), "Покупная", true),
            new(nameof(BlankType.Casting), "Литье", true),
            new(nameof(BlankType.Forging), "Поковка", true),
            new(nameof(BlankType.Unknown), "Не распознано", true)
        ]);

    private static MeasurementUnit ParseUnitCode(string code)
    {
        if (Enum.TryParse<MeasurementUnit>(code, ignoreCase: true, out var parsed))
        {
            return parsed;
        }

        var normalized = UiText.Clean(code);
        if (normalized.Contains("кг", StringComparison.OrdinalIgnoreCase)) return MeasurementUnit.Kilogram;
        if (normalized.Contains("м", StringComparison.OrdinalIgnoreCase)) return MeasurementUnit.Meter;
        return MeasurementUnit.Piece;
    }

    private static BlankType ParseBlankTypeCode(string code) =>
        Enum.TryParse<BlankType>(code, ignoreCase: true, out var parsed) ? parsed : BlankType.CustomBlank;

    private static string CleanName(string name, string fallback) =>
        string.IsNullOrWhiteSpace(name) ? fallback.Trim() : name.Trim();
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
        new(BlankType.WeldingElement, "Сварное изделие"),
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
        new(MeasurementUnit.Meter, "мм"),
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

    public static string DisplayUnit(MeasurementUnit unit) => UiReferenceData.DisplayUnit(unit);
    public static bool IsMillimeterOption(DisplayOption<MeasurementUnit>? option) =>
        option is not null &&
        option.Value == MeasurementUnit.Meter &&
        string.Equals(Clean(option.DisplayName), "мм", StringComparison.OrdinalIgnoreCase);

    public static string DisplayBlankType(BlankType type) => UiReferenceData.DisplayBlankType(type);

    public static string DisplayHiddenBlankType(BlankType type) => type switch
    {
        BlankType.Channel => "Швеллер",
        BlankType.IBeam => "Двутавр",
        BlankType.BronzeBar => "Пруток бронзовый",
        BlankType.BronzeSheet => "Лист бронзовый",
        _ => type.ToString()
    };

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

public sealed record NsiDeleteSnapshot(
    List<NsiCanonicalBlankSnapshot> Blanks,
    List<NsiBlankAliasSnapshot> Aliases,
    List<NsiPartBlankMapSnapshot> Maps,
    List<NsiStockLinkSnapshot> StockLinks,
    List<NsiCalculationLinkSnapshot> CalculationLinks);

public sealed record NsiCanonicalBlankSnapshot(
    long Id,
    string CanonicalName,
    string CanonicalKey,
    BlankType BlankType,
    string? Material,
    string? MaterialGost,
    string? ProfileGost,
    decimal? DiameterMm,
    decimal? WidthMm,
    decimal? HeightMm,
    decimal? ThicknessMm,
    decimal? WallThicknessMm,
    decimal? LengthMm,
    MeasurementUnit BaseUnit,
    I012Status I012Status,
    string? I012Section,
    DateTime CreatedAt,
    DateTime UpdatedAt,
    bool IsActive)
{
    public CanonicalBlank ToEntity()
    {
        var blank = new CanonicalBlank();
        ApplyTo(blank);
        return blank;
    }

    public void ApplyTo(CanonicalBlank blank)
    {
        blank.Id = Id;
        blank.CanonicalName = CanonicalName;
        blank.CanonicalKey = CanonicalKey;
        blank.BlankType = BlankType;
        blank.Material = Material;
        blank.MaterialGost = MaterialGost;
        blank.ProfileGost = ProfileGost;
        blank.DiameterMm = DiameterMm;
        blank.WidthMm = WidthMm;
        blank.HeightMm = HeightMm;
        blank.ThicknessMm = ThicknessMm;
        blank.WallThicknessMm = WallThicknessMm;
        blank.LengthMm = LengthMm;
        blank.BaseUnit = BaseUnit;
        blank.I012Status = I012Status;
        blank.I012Section = I012Section;
        blank.CreatedAt = CreatedAt;
        blank.UpdatedAt = UpdatedAt;
        blank.IsActive = IsActive;
    }
}

public sealed record NsiBlankAliasSnapshot(
    long Id,
    long CanonicalBlankId,
    string OneCCode,
    string SourceName,
    string NormalizedSourceName,
    string Source,
    DateTime ImportedAt,
    DateTime UpdatedAt,
    bool IsActive)
{
    public BlankAlias ToEntity()
    {
        var alias = new BlankAlias();
        ApplyTo(alias);
        return alias;
    }

    public void ApplyTo(BlankAlias alias)
    {
        alias.Id = Id;
        alias.CanonicalBlankId = CanonicalBlankId;
        alias.OneCCode = OneCCode;
        alias.SourceName = SourceName;
        alias.NormalizedSourceName = NormalizedSourceName;
        alias.Source = Source;
        alias.ImportedAt = ImportedAt;
        alias.UpdatedAt = UpdatedAt;
        alias.IsActive = IsActive;
    }
}

public sealed record NsiPartBlankMapSnapshot(
    long Id,
    long PartId,
    long CanonicalBlankId,
    decimal ConsumptionQuantity,
    MeasurementUnit ConsumptionUnit,
    decimal LossPercent,
    int BlankLeadTimeDays,
    string Source,
    string? SourceFile,
    DateTime UpdatedAt,
    bool IsActive,
    bool IsPrimary)
{
    public PartBlankMap ToEntity()
    {
        var map = new PartBlankMap();
        ApplyTo(map);
        return map;
    }

    public void ApplyTo(PartBlankMap map)
    {
        map.Id = Id;
        map.PartId = PartId;
        map.CanonicalBlankId = CanonicalBlankId;
        map.ConsumptionQuantity = ConsumptionQuantity;
        map.ConsumptionUnit = ConsumptionUnit;
        map.LossPercent = LossPercent;
        map.BlankLeadTimeDays = BlankLeadTimeDays;
        map.Source = Source;
        map.SourceFile = SourceFile;
        map.UpdatedAt = UpdatedAt;
        map.IsActive = IsActive;
        map.IsPrimary = IsPrimary;
    }
}

public sealed record NsiStockLinkSnapshot(long Id, long? BlankAliasId);
public sealed record NsiCalculationLinkSnapshot(long Id, long? CanonicalBlankId);

public sealed record LibraryRow(long PartId, long? PartBlankMapId, long? CanonicalBlankId, string Ips, string? Designation, string PartName, string Note, bool RequiresNitriding, bool RequiresHeatTreatment, bool RequiresChemicalOxidation, bool RequiresKeyway, string BlankSupplyRequirement, bool BlankSupplyRequiresHeatTreatment, bool BlankSupplyRequiresLaserCutting, string? BlankType, string? BlankName, string? Material, string? OneCCode, decimal? ConsumptionQuantity, MeasurementUnit? ConsumptionUnit, string Quantity, string UnitName, int BlankLeadTimeDays, string BlankLeadTimeDaysText, string? Source, string UpdatedAt, bool IsIncomplete);

public sealed record LibraryBlankOption(long Id, BlankType BlankType, string CanonicalName, string? SourceName, string? Material, MeasurementUnit BaseUnit, string? OneCCode)
{
    public string DisplayName => $"{UiText.Clean(string.IsNullOrWhiteSpace(SourceName) ? CanonicalName : SourceName)} | {OneCCode ?? "без УТ"}";
}

public sealed partial class ObjectCardRouteRow(
    string operationNumber,
    string operationName,
    string workCenter,
    string machineTimeText,
    string setupTimeText,
    string auxiliaryTimeText) : ObservableObject
{
    [ObservableProperty] private string operationNumber = operationNumber;
    [ObservableProperty] private string operationName = operationName;
    [ObservableProperty] private string workCenter = workCenter;
    [ObservableProperty] private string machineTimeText = machineTimeText;
    [ObservableProperty] private string setupTimeText = setupTimeText;
    [ObservableProperty] private string auxiliaryTimeText = auxiliaryTimeText;

    public int SequenceValue =>
        int.TryParse(OperationNumber.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(OperationNumber) &&
        string.IsNullOrWhiteSpace(OperationName) &&
        string.IsNullOrWhiteSpace(WorkCenter);
}

public sealed partial class ObjectCardToolRow(
    string operationNumber,
    string operationName,
    string toolingName,
    string consumptionRate) : ObservableObject
{
    [ObservableProperty] private string operationNumber = operationNumber;
    [ObservableProperty] private string operationName = operationName;
    [ObservableProperty] private string toolingName = toolingName;
    [ObservableProperty] private string consumptionRate = consumptionRate;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(OperationNumber) &&
        string.IsNullOrWhiteSpace(OperationName) &&
        string.IsNullOrWhiteSpace(ToolingName) &&
        string.IsNullOrWhiteSpace(ConsumptionRate);
}

public sealed record ObjectCardToolRowDto(string OperationNumber, string OperationName, string ToolingName, string ConsumptionRate);

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

public sealed record BlankSelectionRow(long BlankId, string OneCCode, string SourceName, string BlankType, string Material, string Size, string UnitName, MeasurementUnit BaseUnit, decimal StockValue, string StockQuantity, string Allowance, decimal AllowanceScore, string Details);

public sealed record BlankSelectionRequest(decimal? Diameter, decimal? Width, decimal? Height, decimal? Thickness, decimal? WallThickness, decimal? Length)
{
    public bool HasAnySize => Diameter is not null || Width is not null || Height is not null || Thickness is not null || WallThickness is not null || Length is not null;
}

public sealed record DimensionComparison(string Label, decimal RequiredValue, decimal BlankValue);

public sealed record NsiBlankRow(long AliasId, long CanonicalBlankId, string OneCCode, string SourceName, string UnitName, string BlankType, string Size, string DuplicateKey, string Material, string MaterialGost, string ProfileGost, string CmoStockQuantity, string WarehouseStockQuantity, string StockUnitName, string Price, string BlankName, string Source, string UpdatedAt)
{
    public string StockQuantity => CmoStockQuantity;
}

public sealed record NsiUsageRow(string Ips, string Designation, string PartName, string Quantity, string UnitName, string Source);

public sealed record DrawingLookupRequest(string? Ips, string? Designation, string? Name, string? SourceFilePath);

public static class DrawingPdfOpener
{
    public static async Task<FileInfo?> ResolveDrawingPdfAsync(
        DrawingLookupRequest request,
        IIpsDrawingService drawingService,
        ILogger? logger,
        CancellationToken cancellationToken)
    {
        var ips = UiText.Clean(request.Ips).Trim();
        if (string.IsNullOrWhiteSpace(ips) &&
            string.IsNullOrWhiteSpace(request.Designation) &&
            string.IsNullOrWhiteSpace(request.Name))
        {
            return null;
        }

        try
        {
            var outputDirectory = GetDrawingCacheDirectory();
            var cachedDrawing = GetStableCachedDrawing(request, outputDirectory);
            if (cachedDrawing is not null && IsReadablePdf(cachedDrawing))
            {
                return cachedDrawing;
            }

            var localDrawing = FindLocalDrawingPdf(request);
            if (localDrawing is not null)
            {
                return CopyToStableCache(request, localDrawing, outputDirectory);
            }

            foreach (var query in BuildDrawingQueries(request))
            {
                var bridgeDrawing = await drawingService.FindDrawingPdfAsync(query, outputDirectory, cancellationToken);
                if (bridgeDrawing is not null && IsReadablePdf(bridgeDrawing))
                {
                    return CopyToStableCache(request, bridgeDrawing, outputDirectory);
                }
            }
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex, "Could not resolve drawing for {Ips}", ips);
        }

        return null;
    }

    public static async Task<string> OpenExternalAsync(
        DrawingLookupRequest request,
        IIpsDrawingService drawingService,
        string ownerTitle,
        ILogger? logger,
        CancellationToken cancellationToken)
    {
        var ips = UiText.Clean(request.Ips).Trim();
        if (string.IsNullOrWhiteSpace(ips) &&
            string.IsNullOrWhiteSpace(request.Designation) &&
            string.IsNullOrWhiteSpace(request.Name))
        {
            var emptyMessage = "PDF-чертеж не выбран: в строке нет IPS, обозначения или наименования.";
            MessageBox.Show(emptyMessage, ownerTitle, MessageBoxButton.OK, MessageBoxImage.Information);
            return emptyMessage;
        }

        try
        {
            var outputDirectory = GetDrawingCacheDirectory();
            var cachedDrawing = GetStableCachedDrawing(request, outputDirectory);
            if (cachedDrawing is not null && IsReadablePdf(cachedDrawing))
            {
                OpenFile(cachedDrawing);
                return $"PDF-чертеж открыт из кэша: {cachedDrawing.Name}";
            }

            var localDrawing = FindLocalDrawingPdf(request);
            if (localDrawing is not null)
            {
                var stableDrawing = CopyToStableCache(request, localDrawing, outputDirectory);
                OpenFile(stableDrawing);
                return $"PDF-чертеж открыт: {stableDrawing.Name}";
            }

            foreach (var query in BuildDrawingQueries(request))
            {
                var bridgeDrawing = await drawingService.FindDrawingPdfAsync(query, outputDirectory, cancellationToken);
                if (bridgeDrawing is not null && IsReadablePdf(bridgeDrawing))
                {
                    var stableDrawing = CopyToStableCache(request, bridgeDrawing, outputDirectory);
                    OpenFile(stableDrawing);
                    return $"PDF-чертеж открыт через IPS Bridge: {stableDrawing.Name}";
                }
            }

            var notFoundMessage = string.IsNullOrWhiteSpace(ips)
                ? "PDF-чертеж не найден локально и через IPS Bridge."
                : $"PDF-чертеж IPS {ips} не найден локально и через IPS Bridge.";
            MessageBox.Show(notFoundMessage, ownerTitle, MessageBoxButton.OK, MessageBoxImage.Information);
            return notFoundMessage;
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex, "Could not open drawing for {Ips}", ips);
            var message = string.IsNullOrWhiteSpace(ips)
                ? $"PDF-чертеж не открыт: {ex.GetBaseException().Message}"
                : $"PDF-чертеж IPS {ips} не открыт: {ex.GetBaseException().Message}";
            MessageBox.Show(message, ownerTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            return message;
        }
    }

    private static void OpenFile(FileInfo file) =>
        Process.Start(new ProcessStartInfo(file.FullName) { UseShellExecute = true });

    private static FileInfo? GetStableCachedDrawing(DrawingLookupRequest request, DirectoryInfo outputDirectory)
    {
        var key = FirstNotEmpty(request.Ips, request.Designation, request.Name);
        return string.IsNullOrWhiteSpace(key)
            ? null
            : new FileInfo(Path.Combine(outputDirectory.FullName, MakeStableDrawingFileName(key)));
    }

    private static FileInfo? FindLocalDrawingPdf(DrawingLookupRequest request)
    {
        foreach (var candidate in EnumerateLocalDrawingCandidates(request))
        {
            if (IsReadablePdf(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static FileInfo CopyToStableCache(DrawingLookupRequest request, FileInfo drawing, DirectoryInfo outputDirectory)
    {
        outputDirectory.Create();
        var target = GetStableCachedDrawing(request, outputDirectory) ??
            new FileInfo(Path.Combine(outputDirectory.FullName, MakeStableDrawingFileName(Path.GetFileNameWithoutExtension(drawing.Name))));
        if (!string.Equals(drawing.FullName, target.FullName, StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(drawing.FullName, target.FullName, overwrite: true);
        }

        return target;
    }

    private static IEnumerable<FileInfo> EnumerateLocalDrawingCandidates(DrawingLookupRequest request)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in EnumerateExactDrawingPaths(request))
        {
            if (seen.Add(path))
            {
                yield return new FileInfo(path);
            }
        }

        var tokens = new[]
            {
                request.Ips,
                NormalizeNumericIps(request.Ips),
                request.Designation,
                request.Name,
                Path.GetFileNameWithoutExtension(request.SourceFilePath)
            }
            .Select(NormalizeDrawingSearchText)
            .Where(x => x.Length >= 4)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var directory in EnumerateDrawingDirectories(request))
        {
            if (!directory.Exists)
            {
                continue;
            }

            IEnumerable<FileInfo> files;
            try
            {
                files = directory.EnumerateFiles("*.pdf", SearchOption.TopDirectoryOnly)
                    .Concat(directory.EnumerateFiles("*.PDF", SearchOption.TopDirectoryOnly));
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                if (!seen.Add(file.FullName))
                {
                    continue;
                }

                var normalizedName = NormalizeDrawingSearchText(Path.GetFileNameWithoutExtension(file.Name));
                if (tokens.Any(token => normalizedName.Contains(token, StringComparison.OrdinalIgnoreCase) ||
                    token.Contains(normalizedName, StringComparison.OrdinalIgnoreCase)))
                {
                    yield return file;
                }
            }
        }
    }

    private static IEnumerable<string> EnumerateExactDrawingPaths(DrawingLookupRequest request)
    {
        if (!string.IsNullOrWhiteSpace(request.SourceFilePath))
        {
            var directory = Path.GetDirectoryName(request.SourceFilePath);
            var fileName = Path.GetFileNameWithoutExtension(request.SourceFilePath);
            if (!string.IsNullOrWhiteSpace(directory) && !string.IsNullOrWhiteSpace(fileName))
            {
                yield return Path.Combine(directory, $"{fileName}.pdf");
                yield return Path.Combine(directory, $"{fileName}.PDF");
            }
        }

        foreach (var directory in EnumerateDrawingDirectories(request))
        {
            foreach (var query in BuildDrawingQueries(request))
            {
                yield return Path.Combine(directory.FullName, $"{query}.pdf");
                yield return Path.Combine(directory.FullName, $"{query}.PDF");
            }
        }
    }

    private static IEnumerable<DirectoryInfo> EnumerateDrawingDirectories(DrawingLookupRequest request)
    {
        yield return new DirectoryInfo(Path.Combine(Directory.GetCurrentDirectory(), "Данные для работы", "Чертежи"));
        yield return new DirectoryInfo(Path.Combine(AppContext.BaseDirectory, "Данные для работы", "Чертежи"));
        yield return new DirectoryInfo(Path.Combine(AppContext.BaseDirectory, "drawings"));
        yield return new DirectoryInfo(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "Интеграция с сервисами",
            "factory_ai_assistant",
            "data",
            "drawings"));

        if (!string.IsNullOrWhiteSpace(request.SourceFilePath))
        {
            var sourceDirectory = Path.GetDirectoryName(request.SourceFilePath);
            if (!string.IsNullOrWhiteSpace(sourceDirectory))
            {
                yield return new DirectoryInfo(sourceDirectory);
            }
        }

        var configured = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_DIR");
        if (!string.IsNullOrWhiteSpace(configured))
        {
            foreach (var path in configured.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                yield return new DirectoryInfo(path);
            }
        }
    }

    private static IEnumerable<string> BuildDrawingQueries(DrawingLookupRequest request)
    {
        var values = new List<string?>
        {
            request.Ips,
            NormalizeNumericIps(request.Ips),
            request.Designation,
            request.Name
        };

        return values
            .Select(x => UiText.Clean(x).Trim())
            .Where(x => x.Length >= 3)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string? NormalizeNumericIps(string? value)
    {
        var text = UiText.Clean(value).Trim();
        return text.Length > 0 && text.All(char.IsDigit) && text.Length < 11
            ? text.PadLeft(11, '0')
            : null;
    }

    private static string NormalizeDrawingSearchText(string? value)
    {
        var text = UiText.Clean(value);
        return string.IsNullOrWhiteSpace(text)
            ? string.Empty
            : Regex.Replace(text.ToUpperInvariant(), @"[^0-9A-ZА-Я]+", string.Empty);
    }

    private static string MakeStableDrawingFileName(string value)
    {
        var safeValue = Regex.Replace(value, @"[^A-Za-zА-Яа-я0-9_. -]+", "_").Trim(' ', '_', '.');
        return string.IsNullOrWhiteSpace(safeValue) ? "IPS.pdf" : $"IPS_{safeValue}.pdf";
    }

    private static bool IsReadablePdf(FileInfo file)
    {
        if (!file.Exists || file.Length < 4)
        {
            return false;
        }

        using var stream = file.OpenRead();
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 &&
            header[0] == '%' &&
            header[1] == 'P' &&
            header[2] == 'D' &&
            header[3] == 'F';
    }

    private static DirectoryInfo GetDrawingCacheDirectory()
    {
        var configured = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR");
        var path = string.IsNullOrWhiteSpace(configured)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BlankDemandPlanner",
                "IpsDrawings")
            : configured;
        return Directory.CreateDirectory(path);
    }

    private static string FirstNotEmpty(params string?[] values) =>
        values.Select(x => UiText.Clean(x).Trim()).FirstOrDefault(x => x.Length > 0) ?? string.Empty;
}

public static class UiSearchText
{
    public static bool Contains(string? value, string? searchText)
    {
        var normalizedSearch = Normalize(searchText);
        if (normalizedSearch.Length == 0)
        {
            return true;
        }

        var normalizedValue = Normalize(value);
        if (normalizedValue.Contains(normalizedSearch, StringComparison.Ordinal))
        {
            return true;
        }

        return BuildEquivalentTokens(normalizedSearch)
            .Where(x => !string.Equals(x, normalizedSearch, StringComparison.Ordinal))
            .Any(x => normalizedValue.Contains(x, StringComparison.Ordinal));
    }

    public static bool ContainsAnyField(string? searchText, params string?[] values)
    {
        var normalizedSearch = Normalize(searchText);
        if (normalizedSearch.Length == 0)
        {
            return true;
        }

        if (values.Any(value => Contains(value, normalizedSearch)))
        {
            return true;
        }

        var normalizedValues = values
            .Select(Normalize)
            .Where(x => x.Length > 0)
            .ToArray();
        if (normalizedValues.Length == 0)
        {
            return false;
        }

        var combined = string.Join(' ', normalizedValues);
        var tokens = SplitSearchTokens(normalizedSearch);
        return tokens.Length > 1 && tokens.All(token =>
            BuildEquivalentTokens(token).Any(candidate => combined.Contains(candidate, StringComparison.Ordinal)));
    }

    public static bool EqualsNormalized(string? value, string? searchText)
    {
        var normalizedSearch = Normalize(searchText);
        return normalizedSearch.Length != 0 && string.Equals(Normalize(value), normalizedSearch, StringComparison.Ordinal);
    }

    public static string Normalize(string? value)
    {
        var text = UiText.Clean(value).Trim().ToUpperInvariant();
        if (text.Length == 0)
        {
            return string.Empty;
        }

        text = Regex.Replace(text, @"(?<![\p{L}\p{N}])Д\s*(?=\d)", "D", RegexOptions.CultureInvariant);

        var result = new char[text.Length];
        for (var i = 0; i < text.Length; i++)
        {
            result[i] = MapHomoglyph(text[i]);
        }

        return new string(result);
    }

    private static string[] SplitSearchTokens(string normalizedSearch) =>
        normalizedSearch
            .Split([' ', '\t', '\r', '\n', '|', ';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(x => x.Length > 0)
            .ToArray();

    private static IEnumerable<string> BuildEquivalentTokens(string token)
    {
        yield return token;
        if (token.All(char.IsDigit))
        {
            var withoutLeadingZeros = token.TrimStart('0');
            if (withoutLeadingZeros.Length == 0)
            {
                withoutLeadingZeros = "0";
            }

            yield return withoutLeadingZeros;
            if (withoutLeadingZeros.Length is > 0 and < 11)
            {
                yield return withoutLeadingZeros.PadLeft(11, '0');
            }
        }
    }

    private static char MapHomoglyph(char value) => value switch
    {
        'А' => 'A',
        'В' => 'B',
        'Е' => 'E',
        'К' => 'K',
        'М' => 'M',
        'Н' => 'H',
        'О' => 'O',
        'Р' => 'P',
        'С' => 'C',
        'Т' => 'T',
        'У' => 'Y',
        'Х' => 'X',
        'Ё' => 'Е',
        _ => value
    };
}

public static class NsiDuplicateKey
{
    public static string Build(CanonicalBlank? blank, BlankAlias? alias = null)
    {
        if (blank is null || !blank.IsActive)
        {
            return string.Empty;
        }

        var material = NormalizePart(blank.Material);
        var materialGost = NormalizePart(blank.MaterialGost);
        var profileGost = NormalizePart(blank.ProfileGost);
        var dimensions = new[]
        {
            ("D", blank.DiameterMm),
            ("W", blank.WidthMm),
            ("H", blank.HeightMm),
            ("T", blank.ThicknessMm),
            ("S", blank.WallThicknessMm),
            ("L", blank.LengthMm)
        }
            .Where(x => x.Item2 is not null)
            .Select(x => $"{x.Item1}{x.Item2!.Value.ToString("0.####", CultureInfo.InvariantCulture)}")
            .ToArray();

        if (dimensions.Length == 0)
        {
            var name = NormalizePart(alias?.SourceName ?? blank.CanonicalName);
            if (name.Length == 0 || name == material)
            {
                return string.Empty;
            }

            return string.Join("|", (int)blank.BlankType, (int)blank.BaseUnit, name, material, materialGost, profileGost);
        }

        return string.Join("|", new[]
        {
            ((int)blank.BlankType).ToString(CultureInfo.InvariantCulture),
            ((int)blank.BaseUnit).ToString(CultureInfo.InvariantCulture),
            material,
            materialGost,
            profileGost,
            string.Join(";", dimensions)
        });
    }

    private static string NormalizePart(string? value)
    {
        var normalized = UiSearchText.Normalize(value);
        return new string(normalized.Where(x => !char.IsWhiteSpace(x) && x != '-' && x != '_' && x != '/').ToArray());
    }
}

public sealed class EmptyOneCNomenclatureService : IOneCNomenclatureService
{
    public Task<IReadOnlyList<OneCNomenclatureItem>> SearchAsync(string query, int limit, CancellationToken cancellationToken) =>
        Task.FromResult<IReadOnlyList<OneCNomenclatureItem>>([]);

    public Task<IReadOnlyList<OneCNomenclatureItem>> ResolveByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken) =>
        Task.FromResult<IReadOnlyList<OneCNomenclatureItem>>([]);

    public Task<IReadOnlyList<OneCNomenclaturePrice>> ResolvePricesByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken) =>
        Task.FromResult<IReadOnlyList<OneCNomenclaturePrice>>([]);
}

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

public sealed record ProductionLaunchPreview(
    string Ips,
    string Designation,
    string PartName,
    string ComponentOneCCode,
    string ComponentName,
    decimal ConsumptionQuantity,
    MeasurementUnit ConsumptionUnit,
    decimal MaterialStockQuantity,
    decimal MaxQuantity,
    decimal DefaultQuantity,
    string MaterialLine,
    IReadOnlyList<ProductionLaunchShortageRow> ShortageRows);

public sealed record ProductionLaunchDialogResult(decimal Quantity, string Comment, bool Piecewise);

public sealed record ProductionLaunchShortageRow(decimal Quantity, string Project, string MachineNumber, DateTime? DemandDate);

public sealed partial class ExternalServiceWipRow(
    string oneCCode,
    string ips,
    string designation,
    string name,
    decimal availableQuantity,
    string project,
    string machineNumber,
    DateTime? demandDate,
    IReadOnlyList<ExternalServiceDemandAllocation>? demandAllocations = null,
    bool requiresNitriding = false,
    bool requiresHeatTreatment = false,
    bool requiresChemicalOxidation = false,
    bool requiresKeyway = false,
    bool hasDemandRequirement = false) : ObservableObject
{
    public string OneCCode { get; } = oneCCode;
    public string Ips { get; } = ips;
    public string Designation { get; } = designation;
    public string Name { get; } = name;
    public decimal AvailableQuantity { get; } = availableQuantity;
    public string AvailableText { get; } = availableQuantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
    public string Project { get; } = project;
    public string MachineNumber { get; } = machineNumber;
    public DateTime? DemandDate { get; } = demandDate;
    public string DemandDateText { get; } = demandDate?.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("ru-RU")) ?? string.Empty;
    public string Nomenclature { get; } = $"{designation} {name}".Trim();
    public IReadOnlyList<ExternalServiceDemandAllocation> DemandAllocations { get; } = demandAllocations ?? [];
    public bool RequiresNitriding { get; } = requiresNitriding;
    public bool RequiresHeatTreatment { get; } = requiresHeatTreatment;
    public bool RequiresChemicalOxidation { get; } = requiresChemicalOxidation;
    public bool RequiresKeyway { get; } = requiresKeyway;
    public bool HasDemandRequirement { get; } = hasDemandRequirement || (demandAllocations?.Any(x => x.DemandDate is not null && !string.Equals(x.MachineNumber, "на склад", StringComparison.OrdinalIgnoreCase)) ?? false);
    public bool HasExternalServiceRequirement => RequiresNitriding || RequiresHeatTreatment || RequiresChemicalOxidation || RequiresKeyway;

    [ObservableProperty] private bool isSelected;
    [ObservableProperty] private string quantityText = availableQuantity.ToString("0.####", CultureInfo.GetCultureInfo("ru-RU"));
}

public sealed record ExternalServiceDemandAllocation(string Project, string MachineNumber, decimal Quantity, DateTime? DemandDate);

public sealed record ExternalServiceRequestRow(
    string Project,
    string MachineNumber,
    string Ips,
    string Nomenclature,
    string UnitName,
    decimal Quantity,
    string Justification,
    DateTime? DemandDate,
    DateTime ServiceReadyDate,
    bool IsUrgent,
    string Note);

public sealed record MskCsvDetail(string Source, string Ips, string BlankType, string BlankName, string Material, string OneCCode, string ConsumptionQuantity, string UnitName);

public sealed record DisplayOption<T>(T Value, string DisplayName);

public sealed class SimplePageViewModel(string title, string description)
{
    public string Title { get; } = title;
    public string Description { get; } = description;
}





