using BlankDemandPlanner.Core.Enums;

namespace BlankDemandPlanner.Core.Models;

public sealed record BlankNormalizationResult(
    BlankType BlankType,
    string? Material,
    string? MaterialGost,
    string? ProfileGost,
    decimal? Diameter,
    decimal? Width,
    decimal? Height,
    decimal? Thickness,
    decimal? WallThickness,
    decimal? Length,
    string NormalizedName,
    string CanonicalKey,
    double Confidence,
    IReadOnlyList<string> Warnings);

public sealed record PagedQuery(
    int PageNumber,
    int PageSize,
    string? Search = null,
    string? SortColumn = null,
    SortDirection SortDirection = SortDirection.Ascending,
    IReadOnlyDictionary<string, string>? Filters = null);

public sealed record PagedResult<T>(IReadOnlyList<T> Items, int TotalCount, int PageNumber, int PageSize);

public sealed record ImportProgress(int ReadRows, int ProcessedRows, int AddedRows, int UpdatedRows, int SkippedRows, int ErrorRows);

public sealed record ImportReport(int ReadRows, int AddedRows, int UpdatedRows, int SkippedRows, int ErrorRows, IReadOnlyList<string> Errors);

public sealed record OneCStockSyncReport(int ReadRows, int AddedRows, int WarehouseRows, int WipRows, DateTime SyncedAt, IReadOnlyList<string> Errors);

public sealed record ExcelColumnMapping(string ProgramField, string ExcelColumn);

public sealed record ExcelPreview(string FilePath, string SheetName, int HeaderRowNumber, IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string>> Rows);

public sealed record CalculationOptions(long DemandBatchId, long? StockSnapshotId = null, string? Comment = null);

public sealed record DuplicateBlankCandidate(long LeftAliasId, long RightAliasId, string LeftCode, string RightCode, string Reason, double Score);

public sealed record DashboardSummary(
    int Parts,
    int CanonicalBlanks,
    int OneCPositions,
    int PartsWithoutBlank,
    int PartsWithoutMsk,
    int DuplicateCandidates,
    int OutsideI012,
    int PurchasePositions,
    string? LastDemandImport,
    string? LastStockImport,
    string? LastCalculation);
