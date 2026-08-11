using System.Collections.ObjectModel;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.UI.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.EntityFrameworkCore;

namespace BlankDemandPlanner.UI.ViewModels;

public sealed partial class ProductionPlanViewModel(
    BlankDemandPlannerDbContext dbContext,
    IProductionPlanningService planningService,
    IPzmcNeedService pzmcNeedService,
    IPzmcNeedExportService pzmcNeedExportService,
    IPzmcPersonnelAvailabilityService pzmcPersonnelAvailabilityService,
    IPzmcProductionApiClient pzmcProductionApiClient,
    IFileDialogService fileDialogService) : ObservableObject
{
    private readonly List<ProductionScheduleEntry> allScheduleRows = [];
    private readonly List<PzmcNeedRow> allProductionNeedRows = [];
    private CancellationTokenSource? personnelAutoRefreshCts;

    public ObservableCollection<ProductionScheduleEntry> ScheduleRows { get; } = [];
    public ObservableCollection<ProductionScheduleEntry> ShiftRows { get; } = [];
    public ObservableCollection<ProductionEquipment> EquipmentRows { get; } = [];
    public ObservableCollection<ProductionEmployee> EmployeeRows { get; } = [];
    public ObservableCollection<ProductionRouteOperation> RouteRows { get; } = [];
    public ObservableCollection<PlanningIssue> Issues { get; } = [];
    public ObservableCollection<PzmcNeedRow> ProductionNeedRows { get; } = [];
    public ObservableCollection<PzmcPersonnelAvailabilityRow> PersonnelAvailabilityRows { get; } = [];
    public IReadOnlyList<string> AvailabilityFilters { get; } = ["Все", "Есть", "Нет"];

    [ObservableProperty] private string search = string.Empty;
    [ObservableProperty] private string statusText = "План готов к настройке";
    [ObservableProperty] private string utilizationText = "Загрузка оборудования: 0%";
    [ObservableProperty] private string deadlineText = "Риски сроков: 0";
    [ObservableProperty] private string coverageText = "Охвачено деталей: 0";
    [ObservableProperty] private DateTime planningStart = DateTime.Today.AddHours(8);
    [ObservableProperty] private DateTime selectedShiftDate = DateTime.Today;
    [ObservableProperty] private string productionNeedStatusText = "Данные ПО «Производство» еще не загружены";
    [ObservableProperty] private bool isProductionNeedLoading;
    [ObservableProperty] private bool cmoOnly = true;
    [ObservableProperty] private string productionProjectFilter = string.Empty;
    [ObservableProperty] private string productionGroupFilter = string.Empty;
    [ObservableProperty] private string productionTypeFilter = string.Empty;
    [ObservableProperty] private string productionMethodFilter = string.Empty;
    [ObservableProperty] private string productionManagerFilter = string.Empty;
    [ObservableProperty] private string productionSupplierFilter = string.Empty;
    [ObservableProperty] private DateTime? productionDueFrom;
    [ObservableProperty] private DateTime? productionDueTo;
    [ObservableProperty] private string productionPdfFilter = "Все";
    [ObservableProperty] private string productionOrderFilter = "Все";
    [ObservableProperty] private string personnelAvailabilityStatusText = "Доступность персонала еще не загружена";
    [ObservableProperty] private bool isPersonnelAvailabilityLoading;
    [ObservableProperty] private string personnelAutoRefreshText = "Автообновление: не запущено";

    [ObservableProperty] private ProductionEquipment? selectedEquipment;
    [ObservableProperty] private string equipmentCode = string.Empty;
    [ObservableProperty] private string equipmentName = string.Empty;
    [ObservableProperty] private string equipmentModel = string.Empty;
    [ObservableProperty] private string equipmentGroup = string.Empty;
    [ObservableProperty] private decimal equipmentCapacity = 1m;
    [ObservableProperty] private decimal equipmentEfficiency = 1m;
    [ObservableProperty] private string equipmentStatus = "Доступно";
    [ObservableProperty] private string equipmentMaintenance = string.Empty;

    [ObservableProperty] private ProductionEmployee? selectedEmployee;
    [ObservableProperty] private string employeeNumber = string.Empty;
    [ObservableProperty] private string employeeName = string.Empty;
    [ObservableProperty] private string employeeSpecialty = string.Empty;
    [ObservableProperty] private int employeeQualification = 1;
    [ObservableProperty] private int employeeShiftStart = 8;
    [ObservableProperty] private int employeeShiftEnd = 17;
    [ObservableProperty] private decimal employeeMaxHours = 40m;
    [ObservableProperty] private bool employeeAvailable = true;

    [ObservableProperty] private ProductionRouteOperation? selectedRoute;
    [ObservableProperty] private string routeIps = string.Empty;
    [ObservableProperty] private int routeSequence = 10;
    [ObservableProperty] private string routeOperationCode = string.Empty;
    [ObservableProperty] private string routeDescription = string.Empty;
    [ObservableProperty] private string routeEquipmentGroup = string.Empty;
    [ObservableProperty] private string routeSpecialty = string.Empty;
    [ObservableProperty] private int routeQualification = 1;
    [ObservableProperty] private decimal routeSetupMinutes;
    [ObservableProperty] private decimal routePieceMinutes;
    [ObservableProperty] private decimal routeMachineMinutes;
    [ObservableProperty] private decimal routeAuxiliaryMinutes;
    [ObservableProperty] private bool routeParallel;

    [RelayCommand]
    public async Task LoadAsync()
    {
        var equipment = await dbContext.ProductionEquipment.AsNoTracking().OrderBy(x => x.ResourceGroup).ThenBy(x => x.Code).ToListAsync();
        var employees = await dbContext.ProductionEmployees.AsNoTracking().OrderBy(x => x.Specialty).ThenBy(x => x.FullName).ToListAsync();
        var routes = await dbContext.ProductionRouteOperations.AsNoTracking().OrderBy(x => x.Ips).ThenBy(x => x.Sequence).ToListAsync();
        var schedule = await dbContext.ProductionScheduleEntries.AsNoTracking().OrderBy(x => x.PlannedStart).ToListAsync();

        Replace(EquipmentRows, equipment);
        Replace(EmployeeRows, employees);
        Replace(RouteRows, routes);
        allScheduleRows.Clear();
        allScheduleRows.AddRange(schedule);
        ApplyScheduleFilter();
        RefreshKpis();
        StatusText = $"Оборудование: {equipment.Count}; сотрудники: {employees.Count}; операции маршрутов: {routes.Count}; план: {schedule.Count}";
        await LoadProductionNeedsAsync(forceRefresh: false);
        await LoadPersonnelAvailabilityAsync(forceRefresh: false);
        StartPersonnelAutoRefresh();
    }

    [RelayCommand]
    private Task RefreshProductionNeedsAsync() => LoadProductionNeedsAsync(forceRefresh: true);

    [RelayCommand]
    private Task RefreshPersonnelAvailabilityAsync() => LoadPersonnelAvailabilityAsync(forceRefresh: true);

    [RelayCommand]
    private async Task ExportProductionNeedsAsync()
    {
        var outputDirectory = fileDialogService.SelectFolder();
        if (string.IsNullOrWhiteSpace(outputDirectory))
        {
            return;
        }

        try
        {
            var path = await pzmcNeedExportService.ExportAsync(
                ProductionNeedRows.ToArray(),
                outputDirectory,
                CancellationToken.None);
            ProductionNeedStatusText = $"Выгружено строк: {ProductionNeedRows.Count}; файл: {path}";
        }
        catch (Exception ex)
        {
            ProductionNeedStatusText = $"Не удалось сформировать Excel: {ex.GetBaseException().Message}";
        }
    }

    [RelayCommand]
    private async Task BuildPlanAsync()
    {
        var latestBatchId = await dbContext.DemandBatches.AsNoTracking()
            .OrderByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync();
        if (latestBatchId is null)
        {
            StatusText = "Нет загруженной производственной потребности.";
            return;
        }

        var demandItems = await dbContext.DemandItems.AsNoTracking()
            .Where(x => x.DemandBatchId == latestBatchId.Value && x.Quantity > 0)
            .ToListAsync();
        var routes = await dbContext.ProductionRouteOperations.AsNoTracking().ToListAsync();
        var equipment = await dbContext.ProductionEquipment.AsNoTracking().ToListAsync();
        var employees = await dbContext.ProductionEmployees.AsNoTracking().ToListAsync();
        var start = PlanningStart.Date.AddHours(PlanningStart.Hour == 0 ? 8 : PlanningStart.Hour);

        var result = planningService.BuildPlan(
            demandItems.Select(x => new PlanningDemand(
                x.Id,
                JoinOrder(x.Project, x.SerialNumber),
                x.Ips,
                x.SourcePartName ?? x.Part?.Name ?? x.Ips,
                x.Quantity,
                x.DemandDate?.Date.AddHours(17) ?? start.AddDays(30),
                CalculatePriority(x.DemandDate, start))).ToArray(),
            routes.Select(x => new PlanningRoute(
                x.Id, x.Ips, x.Sequence, x.OperationCode, x.Description, x.EquipmentGroup,
                x.RequiredSpecialty, x.MinimumQualification, x.SetupMinutes, x.PieceMinutes,
                x.MachineMinutes, x.AuxiliaryMinutes, x.CanRunInParallel)).ToArray(),
            equipment.Select(x => new PlanningEquipmentResource(
                x.Id, x.Code, x.Name, x.ResourceGroup, x.CapacityPerHour, x.EfficiencyFactor,
                x.IsActive && x.Status.Equals("Доступно", StringComparison.OrdinalIgnoreCase))).ToArray(),
            employees.Select(x => new PlanningEmployeeResource(
                x.Id, x.PersonnelNumber, x.FullName, x.Specialty, x.QualificationLevel,
                x.ShiftStartHour, x.ShiftEndHour, x.MaxHoursPerWeek, x.IsAvailable)).ToArray(),
            start);

        await using var transaction = await dbContext.Database.BeginTransactionAsync();
        await dbContext.ProductionScheduleEntries.ExecuteDeleteAsync();
        dbContext.ProductionScheduleEntries.AddRange(result.Operations.Select(x => new ProductionScheduleEntry
        {
            DemandItemId = x.DemandItemId,
            RouteOperationId = x.RouteOperationId,
            EquipmentId = x.EquipmentId,
            EmployeeId = x.EmployeeId,
            OrderNumber = x.OrderNumber,
            Ips = x.Ips,
            PartName = x.PartName,
            OperationCode = x.OperationCode,
            OperationName = x.OperationName,
            EquipmentName = x.EquipmentName,
            EmployeeName = x.EmployeeName,
            Quantity = x.Quantity,
            Priority = x.Priority,
            PlannedStart = x.PlannedStart,
            PlannedEnd = x.PlannedEnd,
            DueDate = x.DueDate,
            Status = x.Status
        }));
        await dbContext.SaveChangesAsync();
        await transaction.CommitAsync();

        Replace(Issues, result.Issues);
        await LoadAsync();
        Replace(Issues, result.Issues);
        StatusText = $"План рассчитан: {result.Operations.Count} операций; замечаний: {result.Issues.Count}.";
    }

    [RelayCommand]
    private void EditEquipment()
    {
        if (SelectedEquipment is null) return;
        EquipmentCode = SelectedEquipment.Code;
        EquipmentName = SelectedEquipment.Name;
        EquipmentModel = SelectedEquipment.Model ?? string.Empty;
        EquipmentGroup = SelectedEquipment.ResourceGroup;
        EquipmentCapacity = SelectedEquipment.CapacityPerHour;
        EquipmentEfficiency = SelectedEquipment.EfficiencyFactor;
        EquipmentStatus = SelectedEquipment.Status;
        EquipmentMaintenance = SelectedEquipment.MaintenanceSchedule ?? string.Empty;
    }

    [RelayCommand]
    private async Task SaveEquipmentAsync()
    {
        if (string.IsNullOrWhiteSpace(EquipmentCode) || string.IsNullOrWhiteSpace(EquipmentName) || string.IsNullOrWhiteSpace(EquipmentGroup))
        {
            StatusText = "Для оборудования заполните ID, наименование и группу.";
            return;
        }

        var entity = SelectedEquipment is null
            ? new ProductionEquipment()
            : await dbContext.ProductionEquipment.FindAsync(SelectedEquipment.Id) ?? new ProductionEquipment();
        entity.Code = EquipmentCode.Trim();
        entity.Name = EquipmentName.Trim();
        entity.Model = NullIfEmpty(EquipmentModel);
        entity.ResourceGroup = EquipmentGroup.Trim();
        entity.CapacityPerHour = Math.Max(0, EquipmentCapacity);
        entity.EfficiencyFactor = Math.Clamp(EquipmentEfficiency, 0.1m, 2m);
        entity.Status = string.IsNullOrWhiteSpace(EquipmentStatus) ? "Доступно" : EquipmentStatus.Trim();
        entity.MaintenanceSchedule = NullIfEmpty(EquipmentMaintenance);
        entity.IsActive = true;
        entity.UpdatedAt = DateTime.UtcNow;
        if (entity.Id == 0) dbContext.ProductionEquipment.Add(entity);
        await dbContext.SaveChangesAsync();
        ClearEquipment();
        await LoadAsync();
        StatusText = "Оборудование сохранено.";
    }

    [RelayCommand]
    private void ClearEquipment()
    {
        SelectedEquipment = null;
        EquipmentCode = EquipmentName = EquipmentModel = EquipmentGroup = EquipmentMaintenance = string.Empty;
        EquipmentCapacity = EquipmentEfficiency = 1m;
        EquipmentStatus = "Доступно";
    }

    [RelayCommand]
    private void EditEmployee()
    {
        if (SelectedEmployee is null) return;
        EmployeeNumber = SelectedEmployee.PersonnelNumber;
        EmployeeName = SelectedEmployee.FullName;
        EmployeeSpecialty = SelectedEmployee.Specialty;
        EmployeeQualification = SelectedEmployee.QualificationLevel;
        EmployeeShiftStart = SelectedEmployee.ShiftStartHour;
        EmployeeShiftEnd = SelectedEmployee.ShiftEndHour;
        EmployeeMaxHours = SelectedEmployee.MaxHoursPerWeek;
        EmployeeAvailable = SelectedEmployee.IsAvailable;
    }

    [RelayCommand]
    private async Task SaveEmployeeAsync()
    {
        if (string.IsNullOrWhiteSpace(EmployeeNumber) || string.IsNullOrWhiteSpace(EmployeeName) || string.IsNullOrWhiteSpace(EmployeeSpecialty))
        {
            StatusText = "Для сотрудника заполните ID, ФИО и специальность.";
            return;
        }

        var entity = SelectedEmployee is null
            ? new ProductionEmployee()
            : await dbContext.ProductionEmployees.FindAsync(SelectedEmployee.Id) ?? new ProductionEmployee();
        entity.PersonnelNumber = EmployeeNumber.Trim();
        entity.FullName = EmployeeName.Trim();
        entity.Specialty = EmployeeSpecialty.Trim();
        entity.QualificationLevel = Math.Clamp(EmployeeQualification, 1, 5);
        entity.ShiftStartHour = Math.Clamp(EmployeeShiftStart, 0, 23);
        entity.ShiftEndHour = Math.Clamp(EmployeeShiftEnd, entity.ShiftStartHour + 1, 24);
        entity.MaxHoursPerWeek = Math.Max(1, EmployeeMaxHours);
        entity.IsAvailable = EmployeeAvailable;
        entity.UpdatedAt = DateTime.UtcNow;
        if (entity.Id == 0) dbContext.ProductionEmployees.Add(entity);
        await dbContext.SaveChangesAsync();
        ClearEmployee();
        await LoadAsync();
        StatusText = "Сотрудник сохранен.";
    }

    [RelayCommand]
    private void ClearEmployee()
    {
        SelectedEmployee = null;
        EmployeeNumber = EmployeeName = EmployeeSpecialty = string.Empty;
        EmployeeQualification = 1;
        EmployeeShiftStart = 8;
        EmployeeShiftEnd = 17;
        EmployeeMaxHours = 40m;
        EmployeeAvailable = true;
    }

    [RelayCommand]
    private void EditRoute()
    {
        if (SelectedRoute is null) return;
        RouteIps = SelectedRoute.Ips;
        RouteSequence = SelectedRoute.Sequence;
        RouteOperationCode = SelectedRoute.OperationCode;
        RouteDescription = SelectedRoute.Description;
        RouteEquipmentGroup = SelectedRoute.EquipmentGroup;
        RouteSpecialty = SelectedRoute.RequiredSpecialty;
        RouteQualification = SelectedRoute.MinimumQualification;
        RouteSetupMinutes = SelectedRoute.SetupMinutes;
        RoutePieceMinutes = SelectedRoute.PieceMinutes;
        RouteMachineMinutes = SelectedRoute.MachineMinutes;
        RouteAuxiliaryMinutes = SelectedRoute.AuxiliaryMinutes;
        RouteParallel = SelectedRoute.CanRunInParallel;
    }

    [RelayCommand]
    private async Task SaveRouteAsync()
    {
        if (string.IsNullOrWhiteSpace(RouteIps) || string.IsNullOrWhiteSpace(RouteOperationCode) ||
            string.IsNullOrWhiteSpace(RouteEquipmentGroup) || string.IsNullOrWhiteSpace(RouteSpecialty))
        {
            StatusText = "Для операции заполните IPS, код, группу оборудования и специальность.";
            return;
        }

        var entity = SelectedRoute is null
            ? new ProductionRouteOperation()
            : await dbContext.ProductionRouteOperations.FindAsync(SelectedRoute.Id) ?? new ProductionRouteOperation();
        entity.Ips = RouteIps.Trim();
        entity.Sequence = Math.Max(1, RouteSequence);
        entity.OperationCode = RouteOperationCode.Trim();
        entity.Description = RouteDescription.Trim();
        entity.EquipmentGroup = RouteEquipmentGroup.Trim();
        entity.RequiredSpecialty = RouteSpecialty.Trim();
        entity.MinimumQualification = Math.Clamp(RouteQualification, 1, 5);
        entity.SetupMinutes = Math.Max(0, RouteSetupMinutes);
        entity.PieceMinutes = Math.Max(0, RoutePieceMinutes);
        entity.MachineMinutes = Math.Max(0, RouteMachineMinutes);
        entity.AuxiliaryMinutes = Math.Max(0, RouteAuxiliaryMinutes);
        entity.CanRunInParallel = RouteParallel;
        entity.UpdatedAt = DateTime.UtcNow;
        if (entity.Id == 0) dbContext.ProductionRouteOperations.Add(entity);
        await dbContext.SaveChangesAsync();
        ClearRoute();
        await LoadAsync();
        StatusText = "Операция маршрута сохранена.";
    }

    [RelayCommand]
    private void ClearRoute()
    {
        SelectedRoute = null;
        RouteIps = RouteOperationCode = RouteDescription = RouteEquipmentGroup = RouteSpecialty = string.Empty;
        RouteSequence = 10;
        RouteQualification = 1;
        RouteSetupMinutes = RoutePieceMinutes = RouteMachineMinutes = RouteAuxiliaryMinutes = 0;
        RouteParallel = false;
    }

    partial void OnSearchChanged(string value) => ApplyScheduleFilter();
    partial void OnSelectedShiftDateChanged(DateTime value) => RefreshShiftRows();
    partial void OnProductionProjectFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionGroupFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionTypeFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionMethodFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionManagerFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionSupplierFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionDueFromChanged(DateTime? value) => ApplyProductionNeedFilters();
    partial void OnProductionDueToChanged(DateTime? value) => ApplyProductionNeedFilters();
    partial void OnProductionPdfFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnProductionOrderFilterChanged(string value) => ApplyProductionNeedFilters();
    partial void OnCmoOnlyChanged(bool value) => _ = LoadProductionNeedsAsync(forceRefresh: false);

    private async Task LoadProductionNeedsAsync(bool forceRefresh)
    {
        if (IsProductionNeedLoading)
        {
            return;
        }

        IsProductionNeedLoading = true;
        ProductionNeedStatusText = forceRefresh
            ? "Получение актуальных данных из ПО «Производство»..."
            : "Загрузка дефицита...";
        try
        {
            var result = await pzmcNeedService.GetDeficitAsync(
                new PzmcNeedFilters(CmoOnly: CmoOnly),
                forceRefresh,
                CancellationToken.None);
            allProductionNeedRows.Clear();
            allProductionNeedRows.AddRange(result.Rows);
            ApplyProductionNeedFilters();
            ProductionNeedStatusText = result.Status;
        }
        catch (Exception ex)
        {
            ProductionNeedStatusText = $"Данные «Производство» недоступны: {ex.GetBaseException().Message}";
        }
        finally
        {
            IsProductionNeedLoading = false;
        }
    }

    private async Task LoadPersonnelAvailabilityAsync(bool forceRefresh)
    {
        if (IsPersonnelAvailabilityLoading)
        {
            return;
        }

        IsPersonnelAvailabilityLoading = true;
        PersonnelAvailabilityStatusText = forceRefresh
            ? "Получение актуальной доступности персонала из ПО «Производство»..."
            : "Загрузка доступности персонала...";
        try
        {
            var result = await pzmcPersonnelAvailabilityService.GetAvailabilityAsync(forceRefresh, CancellationToken.None);
            Replace(PersonnelAvailabilityRows, result.Rows);
            PersonnelAvailabilityStatusText = result.Status;
        }
        catch (Exception ex)
        {
            PersonnelAvailabilityStatusText = $"Данные персонала недоступны: {ex.GetBaseException().Message}";
        }
        finally
        {
            IsPersonnelAvailabilityLoading = false;
        }
    }

    private void StartPersonnelAutoRefresh()
    {
        if (personnelAutoRefreshCts is not null)
        {
            return;
        }

        var refreshMinutes = Math.Clamp(pzmcProductionApiClient.PersonnelRefreshMinutes, 5, 15);
        personnelAutoRefreshCts = new CancellationTokenSource();
        PersonnelAutoRefreshText = pzmcProductionApiClient.MechanicalEntityId is null
            ? $"Автообновление: каждые {refreshMinutes} мин.; entity_id цеха не задан"
            : $"Автообновление: каждые {refreshMinutes} мин.; entity_id={pzmcProductionApiClient.MechanicalEntityId}";
        _ = RunPersonnelAutoRefreshAsync(TimeSpan.FromMinutes(refreshMinutes), personnelAutoRefreshCts.Token);
    }

    private async Task RunPersonnelAutoRefreshAsync(TimeSpan interval, CancellationToken cancellationToken)
    {
        try
        {
            using var timer = new PeriodicTimer(interval);
            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                await LoadPersonnelAvailabilityAsync(forceRefresh: true);
            }
        }
        catch (OperationCanceledException)
        {
        }
    }

    private void ApplyProductionNeedFilters()
    {
        var hasPdf = ParseAvailability(ProductionPdfFilter);
        var hasOrder = ParseAvailability(ProductionOrderFilter);
        var rows = allProductionNeedRows.Where(x =>
            UiSearchText.Contains(x.Project, ProductionProjectFilter) &&
            UiSearchText.Contains(x.ProductGroup, ProductionGroupFilter) &&
            UiSearchText.Contains(x.ManufacturingType, ProductionTypeFilter) &&
            UiSearchText.Contains(x.ManufacturingMethod, ProductionMethodFilter) &&
            UiSearchText.Contains(x.Manager, ProductionManagerFilter) &&
            UiSearchText.Contains(x.Supplier, ProductionSupplierFilter) &&
            (ProductionDueFrom is null || x.NeedDate.Date >= ProductionDueFrom.Value.Date) &&
            (ProductionDueTo is null || x.NeedDate.Date <= ProductionDueTo.Value.Date) &&
            (hasPdf is null || x.HasPdf == hasPdf.Value) &&
            (hasOrder is null || x.HasOrder == hasOrder.Value));
        Replace(ProductionNeedRows, rows);
    }

    private static bool? ParseAvailability(string? value) => value switch
    {
        "Есть" => true,
        "Нет" => false,
        _ => null
    };

    private void ApplyScheduleFilter()
    {
        var value = Search.Trim();
        var rows = string.IsNullOrWhiteSpace(value)
            ? allScheduleRows
            : allScheduleRows.Where(x =>
                UiSearchText.Contains(x.OrderNumber, value) ||
                UiSearchText.Contains(x.Ips, value) ||
                UiSearchText.Contains(x.PartName, value) ||
                UiSearchText.Contains(x.OperationCode, value) ||
                UiSearchText.Contains(x.OperationName, value) ||
                UiSearchText.Contains(x.EquipmentName, value) ||
                UiSearchText.Contains(x.EmployeeName, value) ||
                UiSearchText.Contains(x.Status, value)).ToList();
        Replace(ScheduleRows, rows);
        RefreshShiftRows();
    }

    private void RefreshShiftRows() =>
        Replace(ShiftRows, allScheduleRows.Where(x => x.PlannedStart.Date == SelectedShiftDate.Date).OrderBy(x => x.EmployeeName).ThenBy(x => x.PlannedStart));

    private void RefreshKpis()
    {
        var risks = allScheduleRows.Count(x => x.PlannedEnd > x.DueDate);
        var parts = allScheduleRows.Select(x => x.Ips).Distinct(StringComparer.OrdinalIgnoreCase).Count();
        var equipmentCount = Math.Max(1, EquipmentRows.Count(x => x.IsActive));
        var totalMinutes = allScheduleRows.Sum(x => Math.Max(0, (x.PlannedEnd - x.PlannedStart).TotalMinutes));
        var horizonDays = allScheduleRows.Count == 0
            ? 1
            : Math.Max(1, (allScheduleRows.Max(x => x.PlannedEnd).Date - allScheduleRows.Min(x => x.PlannedStart).Date).Days + 1);
        var utilization = Math.Min(100, totalMinutes / (equipmentCount * horizonDays * 9 * 60) * 100);
        UtilizationText = $"Загрузка оборудования: {utilization:0.#}%";
        DeadlineText = $"Риски сроков: {risks}";
        CoverageText = $"Охвачено деталей: {parts}";
    }

    private static int CalculatePriority(DateTime? dueDate, DateTime start)
    {
        if (dueDate is null) return 10;
        var days = (dueDate.Value.Date - start.Date).TotalDays;
        return days <= 7 ? 100 : days <= 30 ? 50 : 10;
    }

    private static string JoinOrder(string? project, string? machine) =>
        string.Join(" / ", new[] { project, machine }.Where(x => !string.IsNullOrWhiteSpace(x)));

    private static string? NullIfEmpty(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static void Replace<T>(ObservableCollection<T> target, IEnumerable<T> source)
    {
        target.Clear();
        foreach (var item in source) target.Add(item);
    }
}
