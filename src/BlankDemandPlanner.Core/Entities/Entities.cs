using BlankDemandPlanner.Core.Enums;

namespace BlankDemandPlanner.Core.Entities;

public abstract class Entity
{
    public long Id { get; set; }
}

public sealed class Part : Entity
{
    public string Ips { get; set; } = string.Empty;
    public string? Designation { get; set; }
    public string Name { get; set; } = string.Empty;
    public bool HasMsk { get; set; }
    public bool RequiresNitriding { get; set; }
    public bool RequiresHeatTreatment { get; set; }
    public bool RequiresChemicalOxidation { get; set; }
    public bool RequiresKeyway { get; set; }
    public bool BlankSupplyRequiresHeatTreatment { get; set; }
    public bool BlankSupplyRequiresLaserCutting { get; set; }
    public string? Source { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public ICollection<PartBlankMap> BlankMaps { get; set; } = new List<PartBlankMap>();
}

public sealed class CanonicalBlank : Entity
{
    public string CanonicalName { get; set; } = string.Empty;
    public string CanonicalKey { get; set; } = string.Empty;
    public BlankType BlankType { get; set; } = BlankType.Unknown;
    public string? Material { get; set; }
    public string? MaterialGost { get; set; }
    public string? ProfileGost { get; set; }
    public decimal? DiameterMm { get; set; }
    public decimal? WidthMm { get; set; }
    public decimal? HeightMm { get; set; }
    public decimal? ThicknessMm { get; set; }
    public decimal? WallThicknessMm { get; set; }
    public decimal? LengthMm { get; set; }
    public MeasurementUnit BaseUnit { get; set; } = MeasurementUnit.Piece;
    public I012Status I012Status { get; set; } = I012Status.Unrecognized;
    public string? I012Section { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public bool IsActive { get; set; } = true;
    public ICollection<BlankAlias> Aliases { get; set; } = new List<BlankAlias>();
    public ICollection<PartBlankMap> PartMaps { get; set; } = new List<PartBlankMap>();
}

public sealed class BlankAlias : Entity
{
    public long CanonicalBlankId { get; set; }
    public CanonicalBlank? CanonicalBlank { get; set; }
    public string OneCCode { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string NormalizedSourceName { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public DateTime ImportedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public bool IsActive { get; set; } = true;
}

public sealed class PartBlankMap : Entity
{
    public long PartId { get; set; }
    public Part? Part { get; set; }
    public long CanonicalBlankId { get; set; }
    public CanonicalBlank? CanonicalBlank { get; set; }
    public decimal ConsumptionQuantity { get; set; } = 1m;
    public MeasurementUnit ConsumptionUnit { get; set; } = MeasurementUnit.Piece;
    public decimal LossPercent { get; set; }
    public int BlankLeadTimeDays { get; set; } = 30;
    public string Source { get; set; } = string.Empty;
    public string? SourceFile { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public bool IsActive { get; set; } = true;
    public bool IsPrimary { get; set; } = true;
}

public sealed class DemandBatch : Entity
{
    public string Name { get; set; } = string.Empty;
    public string? SourceFile { get; set; }
    public string? FileHashSha256 { get; set; }
    public DateTime ImportedAt { get; set; } = DateTime.UtcNow;
    public DateTime? PeriodFrom { get; set; }
    public DateTime? PeriodTo { get; set; }
    public ICollection<DemandItem> Items { get; set; } = new List<DemandItem>();
}

public sealed class DemandItem : Entity
{
    public long DemandBatchId { get; set; }
    public DemandBatch? DemandBatch { get; set; }
    public long? PartId { get; set; }
    public Part? Part { get; set; }
    public string Ips { get; set; } = string.Empty;
    public string? SourcePartName { get; set; }
    public string? Project { get; set; }
    public string? SerialNumber { get; set; }
    public string? Unit { get; set; }
    public decimal Quantity { get; set; }
    public DateTime? DemandDate { get; set; }
    public string? ProductionSystem { get; set; }
}

public sealed class StockSnapshot : Entity
{
    public DateTime SnapshotDate { get; set; } = DateTime.UtcNow;
    public string? SourceFile { get; set; }
    public DateTime ImportedAt { get; set; } = DateTime.UtcNow;
    public ICollection<StockItem> Items { get; set; } = new List<StockItem>();
}

public sealed class StockItem : Entity
{
    public long StockSnapshotId { get; set; }
    public StockSnapshot? StockSnapshot { get; set; }
    public long? BlankAliasId { get; set; }
    public BlankAlias? BlankAlias { get; set; }
    public string OneCCode { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public MeasurementUnit Unit { get; set; } = MeasurementUnit.Piece;
    public string? Warehouse { get; set; }
}

public sealed class OneCPriceItem : Entity
{
    public string LookupKey { get; set; } = string.Empty;
    public string Code { get; set; } = string.Empty;
    public string Article { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public string Currency { get; set; } = string.Empty;
    public string PriceType { get; set; } = string.Empty;
    public DateTime SyncedAt { get; set; } = DateTime.UtcNow;
}

public sealed class CalculationRun : Entity
{
    public DateTime StartedAt { get; set; } = DateTime.UtcNow;
    public DateTime? FinishedAt { get; set; }
    public long DemandBatchId { get; set; }
    public long? StockSnapshotId { get; set; }
    public string? Comment { get; set; }
    public ICollection<CalculationItem> Items { get; set; } = new List<CalculationItem>();
}

public sealed class CalculationItem : Entity
{
    public long CalculationRunId { get; set; }
    public CalculationRun? CalculationRun { get; set; }
    public long? CanonicalBlankId { get; set; }
    public CanonicalBlank? CanonicalBlank { get; set; }
    public string CanonicalName { get; set; } = string.Empty;
    public string? PrimaryOneCCode { get; set; }
    public string OneCCodes { get; set; } = string.Empty;
    public MeasurementUnit Unit { get; set; }
    public decimal TotalRequired { get; set; }
    public decimal TotalStock { get; set; }
    public decimal PurchaseQuantity { get; set; }
    public CalculationStatus Status { get; set; }
    public string? Comment { get; set; }
    public ICollection<CalculationItemSource> Sources { get; set; } = new List<CalculationItemSource>();
}

public sealed class CalculationItemSource : Entity
{
    public long CalculationItemId { get; set; }
    public CalculationItem? CalculationItem { get; set; }
    public string Ips { get; set; } = string.Empty;
    public string PartName { get; set; } = string.Empty;
    public decimal DemandQuantity { get; set; }
    public decimal RequiredQuantity { get; set; }
    public MeasurementUnit Unit { get; set; }
}

public sealed class I012CatalogEntry : Entity
{
    public BlankType BlankType { get; set; }
    public string SizeKey { get; set; } = string.Empty;
    public string MaterialKey { get; set; } = string.Empty;
    public string Section { get; set; } = string.Empty;
    public string? Source { get; set; }
}

public sealed class MskRecord : Entity
{
    public string Ips { get; set; } = string.Empty;
    public string? Designation { get; set; }
    public string Name { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public DateTime ImportedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ImportProfile : Entity
{
    public string ProfileName { get; set; } = string.Empty;
    public string ProfileType { get; set; } = string.Empty;
    public string? SheetNamePattern { get; set; }
    public string ColumnMappingsJson { get; set; } = "{}";
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ImportRun : Entity
{
    public string ImportType { get; set; } = string.Empty;
    public string SourceFile { get; set; } = string.Empty;
    public string? FileHashSha256 { get; set; }
    public ImportRunStatus Status { get; set; } = ImportRunStatus.Started;
    public int ReadRows { get; set; }
    public int AddedRows { get; set; }
    public int UpdatedRows { get; set; }
    public int SkippedRows { get; set; }
    public int ErrorRows { get; set; }
    public DateTime StartedAt { get; set; } = DateTime.UtcNow;
    public DateTime? FinishedAt { get; set; }
    public ICollection<ImportError> Errors { get; set; } = new List<ImportError>();
}

public sealed class AppUser : Entity
{
    public string UserName { get; set; } = string.Empty;
    public string NormalizedUserName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string PasswordHash { get; set; } = string.Empty;
    public bool IsAdmin { get; set; }
    public bool IsActive { get; set; } = true;
    public bool MustChangePassword { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    public DateTime? LastLoginAt { get; set; }
    public ICollection<AppUserPermission> Permissions { get; set; } = new List<AppUserPermission>();
}

public sealed class AppUserPermission : Entity
{
    public long AppUserId { get; set; }
    public AppUser? AppUser { get; set; }
    public string PageKey { get; set; } = string.Empty;
    public bool CanRead { get; set; }
    public bool CanEdit { get; set; }
}

public sealed class AuthLoginAttempt : Entity
{
    public string UserName { get; set; } = string.Empty;
    public string NormalizedUserName { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public string? FailureReason { get; set; }
    public string? MachineName { get; set; }
    public DateTime AttemptedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ImportError : Entity
{
    public long ImportRunId { get; set; }
    public ImportRun? ImportRun { get; set; }
    public int RowNumber { get; set; }
    public string FieldName { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
}

public sealed class ChangeHistory : Entity
{
    public string EntityType { get; set; } = string.Empty;
    public long EntityId { get; set; }
    public string FieldName { get; set; } = string.Empty;
    public string? OldValue { get; set; }
    public string? NewValue { get; set; }
    public DateTime ChangedAt { get; set; } = DateTime.UtcNow;
    public string Source { get; set; } = string.Empty;
}

public sealed class AppSetting : Entity
{
    public string Key { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

public sealed class AppLog : Entity
{
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public string Level { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? Exception { get; set; }
}

public sealed class ProductionEquipment : Entity
{
    public string Code { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string? Model { get; set; }
    public string ResourceGroup { get; set; } = string.Empty;
    public decimal CapacityPerHour { get; set; } = 1m;
    public decimal EfficiencyFactor { get; set; } = 1m;
    public string Status { get; set; } = "Доступно";
    public string? MaintenanceSchedule { get; set; }
    public bool IsActive { get; set; } = true;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ProductionEmployee : Entity
{
    public string PersonnelNumber { get; set; } = string.Empty;
    public string FullName { get; set; } = string.Empty;
    public string Specialty { get; set; } = string.Empty;
    public int QualificationLevel { get; set; } = 1;
    public int ShiftStartHour { get; set; } = 8;
    public int ShiftEndHour { get; set; } = 17;
    public decimal MaxHoursPerWeek { get; set; } = 40m;
    public bool IsAvailable { get; set; } = true;
    public string? AbsenceReason { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ProductionRouteOperation : Entity
{
    public string Ips { get; set; } = string.Empty;
    public int Sequence { get; set; }
    public string OperationCode { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string EquipmentGroup { get; set; } = string.Empty;
    public string RequiredSpecialty { get; set; } = string.Empty;
    public int MinimumQualification { get; set; } = 1;
    public decimal SetupMinutes { get; set; }
    public decimal PieceMinutes { get; set; }
    public decimal MachineMinutes { get; set; }
    public decimal AuxiliaryMinutes { get; set; }
    public bool CanRunInParallel { get; set; }
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}

public sealed class ProductionScheduleEntry : Entity
{
    public long DemandItemId { get; set; }
    public long RouteOperationId { get; set; }
    public long EquipmentId { get; set; }
    public long EmployeeId { get; set; }
    public string OrderNumber { get; set; } = string.Empty;
    public string Ips { get; set; } = string.Empty;
    public string PartName { get; set; } = string.Empty;
    public string OperationCode { get; set; } = string.Empty;
    public string OperationName { get; set; } = string.Empty;
    public string EquipmentName { get; set; } = string.Empty;
    public string EmployeeName { get; set; } = string.Empty;
    public decimal Quantity { get; set; }
    public int Priority { get; set; }
    public DateTime PlannedStart { get; set; }
    public DateTime PlannedEnd { get; set; }
    public DateTime DueDate { get; set; }
    public string Status { get; set; } = "Запланировано";
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}
