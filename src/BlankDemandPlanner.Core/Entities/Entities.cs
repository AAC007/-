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
