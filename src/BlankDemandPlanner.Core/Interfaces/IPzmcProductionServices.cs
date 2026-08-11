using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Core.Interfaces;

public interface IPzmcProductionApiClient
{
    bool IsConfigured { get; }
    IReadOnlyList<string> CmoGroups { get; }
    int? MechanicalEntityId { get; }
    int PersonnelRefreshMinutes { get; }
    Task<PzmcProductionSnapshot> LoadSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken);
    Task<PzmcProductionPersonnelSnapshot> LoadPersonnelSnapshotAsync(bool forceRefresh, CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcProductNeed>> GetProductNeedsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcProduct>> GetProductsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcSpecIpsObject>> GetIpsObjectsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcSklad>> GetWarehousesAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcSkladPlanBase>> GetPlannedReceiptsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcAnalog>> GetAnalogsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcProductGroup>> GetProductGroupsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcEmployee>> GetEmployeesAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcEmployeeStatistic>> GetEmployeeStatisticsAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcSkudLogUser>> GetSkudLogUsersAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<PzmcEntitie>> GetEntitiesAsync(CancellationToken cancellationToken);
}

public interface IPzmcNeedService
{
    Task<PzmcNeedResult> GetDeficitAsync(PzmcNeedFilters filters, bool forceRefresh, CancellationToken cancellationToken);
}

public interface IPzmcNeedExportService
{
    Task<string> ExportAsync(IReadOnlyCollection<PzmcNeedRow> rows, string outputDirectory, CancellationToken cancellationToken);
}

public interface IPzmcPersonnelAvailabilityService
{
    Task<PzmcPersonnelAvailabilityResult> GetAvailabilityAsync(bool forceRefresh, CancellationToken cancellationToken);
}
