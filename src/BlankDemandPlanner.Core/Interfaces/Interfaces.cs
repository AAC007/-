using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Core.Interfaces;

public interface IBlankNormalizationService
{
    Task<BlankNormalizationResult> NormalizeAsync(string sourceName, CancellationToken cancellationToken);
}

public interface II012ValidationService
{
    Task<I012Status> ValidateAsync(BlankNormalizationResult blank, CancellationToken cancellationToken);
    string BuildSizeKey(BlankNormalizationResult blank);
    string BuildMaterialKey(string? material);
}

public interface IBlankDemandCalculationService
{
    Task<CalculationRun> CalculateAsync(CalculationOptions options, CancellationToken cancellationToken);
}

public interface IUnitConversionService
{
    bool CanSubtract(MeasurementUnit requiredUnit, MeasurementUnit stockUnit);
    decimal Convert(decimal quantity, MeasurementUnit fromUnit, MeasurementUnit toUnit, CanonicalBlank? blank);
}

public interface IDuplicateDetectionService
{
    Task<IReadOnlyList<DuplicateBlankCandidate>> FindBlankAliasDuplicatesAsync(CancellationToken cancellationToken);
}

public interface IDatabaseBackupService
{
    Task<string> BackupAsync(string reason, CancellationToken cancellationToken);
    Task CleanupAsync(CancellationToken cancellationToken);
}

public interface IExcelImportProfile
{
    string ProfileType { get; }
    bool CanHandle(string sheetName, IReadOnlyList<string> headers);
    IReadOnlyDictionary<string, string> AutoMapColumns(IReadOnlyList<string> headers);
}

public interface IExcelImportService
{
    Task<IReadOnlyList<string>> GetSheetNamesAsync(string filePath, CancellationToken cancellationToken);
    Task<ExcelPreview> PreviewAsync(string filePath, string sheetName, CancellationToken cancellationToken);
    Task<ImportReport> ImportOneCBlanksAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken);
    Task<ImportReport> ImportDemandAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken);
    Task<ImportReport> ImportStockAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken);
    Task<ImportReport> ImportManufacturingBlankLibraryAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken);
}

public interface IOneCStockSyncService
{
    Task<OneCStockSyncReport> SyncAsync(CancellationToken cancellationToken);
}

public interface IOneCNomenclatureService
{
    Task<IReadOnlyList<OneCNomenclatureItem>> SearchAsync(string query, int limit, CancellationToken cancellationToken);
    Task<IReadOnlyList<OneCNomenclatureItem>> ResolveByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken);
    Task<IReadOnlyList<OneCNomenclaturePrice>> ResolvePricesByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken);
}

public interface IOneCProductionLaunchService
{
    Task<OneCProductionLaunchResult> CreateAssemblyAsync(OneCProductionLaunchRequest request, CancellationToken cancellationToken);
}

public interface IOneCGoodsTransferService
{
    Task<OneCGoodsTransferResult> CreateTransferAsync(OneCGoodsTransferRequest request, CancellationToken cancellationToken);
}

public interface IReportExportService
{
    Task<string> ExportCalculationRunAsync(long calculationRunId, string outputDirectory, CancellationToken cancellationToken);
    Task<string> ExportLibraryAsync(string outputDirectory, CancellationToken cancellationToken);
}
