using System.Text.RegularExpressions;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace BlankDemandPlanner.Infrastructure.Export;

public sealed class ReportExportService(BlankDemandPlannerDbContext dbContext) : IReportExportService
{
    private const string OneTimeBlankComment = "Разовая заготовка из расчета";

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

    static ReportExportService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<string> ExportCalculationRunAsync(long calculationRunId, string outputDirectory, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(outputDirectory);
        var run = await dbContext.CalculationRuns.AsNoTracking()
            .Include(x => x.Items)
            .ThenInclude(x => x.Sources)
            .FirstAsync(x => x.Id == calculationRunId, cancellationToken);

        var path = BuildBitrixRequestPath(outputDirectory);
        var materialRows = await BuildMaterialRequestRowsAsync(run, cancellationToken);

        using var package = new ExcelPackage();
        FillMaterialRequest(package.Workbook.Worksheets.Add("Заявка"), materialRows);
        FillGroupedRequest(package.Workbook.Worksheets.Add("По группе"), materialRows);

        return await SavePackageReplacingWhenPossibleAsync(package, path, Path.GetFileNameWithoutExtension(path), cancellationToken);
    }

    public async Task<string> ExportLibraryAsync(string outputDirectory, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(outputDirectory);
        var parts = await dbContext.Parts.AsNoTracking()
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .OrderBy(x => x.Ips)
            .ToListAsync(cancellationToken);

        var path = Path.Combine(outputDirectory, "Библиотека.xlsx");
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Библиотека");
        var headers = new[]
        {
            "IPS",
            "Обозначение",
            "Наименование детали",
            "Вид заготовки",
            "Заготовка",
            "Материал",
            "УТ код",
            "Количество",
            "Ед. изм.",
            "Источник",
            "Обновлено"
        };

        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
        }

        var cleanHeaders = new[]
        {
            "IPS",
            "Обозначение",
            "Наименование детали",
            "Вид заготовки",
            "Заготовка",
            "Материал",
            "УТ код",
            "Количество",
            "Ед. изм.",
            "Срок заготовки, дней",
            "Источник",
            "Обновлено"
        };
        for (var i = 0; i < cleanHeaders.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = cleanHeaders[i];
        }

        sheet.InsertColumn(7, 1);
        sheet.Cells[1, 7].Value = "Условие поставки заготовки";

        var row = 2;
        foreach (var part in parts)
        {
            var map = part.BlankMaps.FirstOrDefault(x => x.IsPrimary) ?? part.BlankMaps.FirstOrDefault();
            var blank = map?.CanonicalBlank;
            var alias = blank?.Aliases.FirstOrDefault(x => x.IsActive);
            var unit = GetEffectiveUnit(map, blank);

            sheet.Cells[row, 1].Value = part.Ips;
            sheet.Cells[row, 2].Value = Clean(part.Designation);
            sheet.Cells[row, 3].Value = Clean(part.Name);
            sheet.Cells[row, 4].Value = blank is null ? string.Empty : DisplayBlankType(blank.BlankType);
            sheet.Cells[row, 5].Value = Clean(alias?.SourceName ?? blank?.CanonicalName);
            sheet.Cells[row, 6].Value = Clean(blank?.Material);
            sheet.Cells[row, 7].Value = BuildBlankSupplyRequirement(part);
            sheet.Cells[row, 8].Value = alias?.OneCCode;
            sheet.Cells[row, 9].Value = map?.ConsumptionQuantity;
            sheet.Cells[row, 10].Value = unit is null ? string.Empty : DisplayUnit(unit.Value);
            sheet.Cells[row, 11].Value = map?.BlankLeadTimeDays ?? 30;
            sheet.Cells[row, 12].Value = Clean(part.Source);
            sheet.Cells[row, 13].Value = part.UpdatedAt.ToLocalTime().ToString("g");
            row++;
        }

        sheet.Cells.AutoFitColumns();
        return await SavePackageReplacingWhenPossibleAsync(package, path, "Библиотека", cancellationToken);
    }

    private static async Task<string> SavePackageReplacingWhenPossibleAsync(ExcelPackage package, string preferredPath, string fileNamePrefix, CancellationToken cancellationToken)
    {
        try
        {
            if (File.Exists(preferredPath))
            {
                File.Delete(preferredPath);
            }

            await package.SaveAsAsync(new FileInfo(preferredPath), cancellationToken);
            return preferredPath;
        }
        catch (IOException)
        {
            return await SaveTimestampedCopyAsync(package, preferredPath, fileNamePrefix, cancellationToken);
        }
        catch (UnauthorizedAccessException)
        {
            return await SaveTimestampedCopyAsync(package, preferredPath, fileNamePrefix, cancellationToken);
        }
    }

    private static async Task<string> SaveTimestampedCopyAsync(ExcelPackage package, string preferredPath, string fileNamePrefix, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(preferredPath) ?? Directory.GetCurrentDirectory();
        var fallbackPath = Path.Combine(directory, $"{fileNamePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        await package.SaveAsAsync(new FileInfo(fallbackPath), cancellationToken);
        return fallbackPath;
    }

    private async Task<IReadOnlyList<MaterialRequestExportRow>> BuildMaterialRequestRowsAsync(CalculationRun run, CancellationToken cancellationToken)
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
            .ToListAsync(cancellationToken);

        var ipsValues = demands.Select(x => x.Ips).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var parts = await dbContext.Parts.AsNoTracking()
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .ThenInclude(x => x!.Aliases)
            .AsSplitQuery()
            .Where(x => ipsValues.Contains(x.Ips))
            .ToDictionaryAsync(x => x.Ips, StringComparer.OrdinalIgnoreCase, cancellationToken);
        var workInProgressByIps = await LoadWorkInProgressByIpsAsync(ipsValues, cancellationToken);
        var oneTimeAssignments = run.Items
            .Where(x => x.Comment?.StartsWith(OneTimeBlankComment, StringComparison.OrdinalIgnoreCase) == true)
            .Select(x => (Item: x, DemandItemId: TryGetOneTimeDemandItemId(x.Comment)))
            .Where(x => x.DemandItemId is not null)
            .ToDictionary(x => x.DemandItemId!.Value, x => x.Item);

        var rows = new List<MaterialRequestExportRow>();
        var number = 1;
        var meterCutSourceCounts = new Dictionary<(long BlankId, MeasurementUnit Unit), int>();
        foreach (var demand in demands)
        {
            var inProduction = ApplyWorkInProgress(demand.Ips, demand.Quantity, workInProgressByIps, out var effectiveDemandQuantity);
            parts.TryGetValue(demand.Ips, out var part);
            if (oneTimeAssignments.TryGetValue(demand.Id, out var oneTimeItem))
            {
                var source = oneTimeItem.Sources.FirstOrDefault();
                rows.Add(new MaterialRequestExportRow(
                    number++,
                    Clean(demand.Project),
                    Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    Clean(FirstNotEmpty(demand.SourcePartName, source?.PartName, demand.Ips)),
                    demand.Quantity,
                    inProduction,
                    oneTimeItem.PrimaryOneCCode ?? string.Empty,
                    Clean(oneTimeItem.CanonicalName),
                    DisplayUnit(oneTimeItem.Unit),
                    oneTimeItem.PurchaseQuantity,
                    oneTimeItem.PurchaseQuantity,
                    BlankDemandDate(demand.DemandDate, 30),
                    string.Empty,
                    null,
                    string.Empty,
                    OneTimeBlankComment));
                continue;
            }

            var maps = part?.BlankMaps
                .Where(x => x.IsActive)
                .GroupBy(x => new { x.CanonicalBlankId, x.ConsumptionUnit })
                .Select(x => x.OrderByDescending(m => m.IsPrimary).ThenByDescending(m => m.UpdatedAt).First())
                .OrderByDescending(x => x.IsPrimary)
                .ToList() ?? [];

            if (maps.Count == 0)
            {
                rows.Add(new MaterialRequestExportRow(
                    number++,
                    Clean(demand.Project),
                    Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    Clean(FirstNotEmpty(demand.SourcePartName, part?.Name, demand.Ips)),
                    demand.Quantity,
                    inProduction,
                    string.Empty,
                    string.Empty,
                    Clean(demand.Unit),
                    effectiveDemandQuantity,
                    effectiveDemandQuantity,
                    BlankDemandDate(demand.DemandDate, 30),
                    string.Empty,
                    null,
                    string.Empty,
                    "Не подобрана заготовка"));
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

                var materialQuantity = Math.Min(required, purchaseQuantity);
                rows.Add(new MaterialRequestExportRow(
                    number++,
                    Clean(demand.Project),
                    Clean(FirstNotEmpty(demand.ProductionSystem, demand.SerialNumber)),
                    demand.Ips,
                    Clean(FirstNotEmpty(demand.SourcePartName, part?.Name, demand.Ips)),
                    demand.Quantity,
                    inProduction,
                    alias?.OneCCode ?? string.Empty,
                    Clean(alias?.SourceName ?? blank?.CanonicalName),
                    DisplayUnit(map.ConsumptionUnit),
                    materialQuantity,
                    purchaseQuantity,
                    BlankDemandDate(demand.DemandDate, map.BlankLeadTimeDays),
                    string.Empty,
                    null,
                    string.Empty,
                    string.Empty));
            }
        }

        return await EnrichRowsWithPricesAsync(rows, cancellationToken);
    }

    private async Task<IReadOnlyList<MaterialRequestExportRow>> EnrichRowsWithPricesAsync(IReadOnlyList<MaterialRequestExportRow> rows, CancellationToken cancellationToken)
    {
        if (rows.Count == 0)
        {
            return rows;
        }

        var codes = rows
            .Select(x => x.OneCCode)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(StockCodeNormalizer.NormalizeForComparison)
            .Where(x => x.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (codes.Length == 0)
        {
            return rows;
        }

        var priceItems = await dbContext.OneCPriceItems.AsNoTracking()
            .Where(x => codes.Contains(x.LookupKey))
            .ToListAsync(cancellationToken);
        var pricesByCode = priceItems
            .Where(x => x.Price > 0)
            .GroupBy(x => x.LookupKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                x => x.Key,
                x => x.OrderByDescending(p => IsPreferredPriceType(p.PriceType)).ThenByDescending(p => p.SyncedAt).First(),
                StringComparer.OrdinalIgnoreCase);

        return rows.Select(row =>
            {
                var key = StockCodeNormalizer.NormalizeForComparison(row.OneCCode);
                return pricesByCode.TryGetValue(key, out var price)
                    ? row with
                    {
                        Price = FormatPrice(price),
                        UnitPrice = price.Price,
                        PriceCurrency = FormatCurrency(price.Currency)
                    }
                    : row;
            })
            .ToList();
    }

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

    private static void FillMaterialRequest(ExcelWorksheet sheet, IEnumerable<MaterialRequestExportRow> rows)
    {
        var headers = new[]
        {
            "№",
            "ПС",
            "№ станка",
            "IPS детали",
            "Наименование",
            "Количество деталей",
            "Код УТ заготовки",
            "Номенклатура",
            "Ед. измерения",
            "Кол-во к закупке",
            "Дата потребности заготовки",
            "Цена",
            "Комментарий"
        };

        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
        }

        var rowIndex = 2;
        foreach (var row in rows)
        {
            sheet.Cells[rowIndex, 1].Value = row.Number;
            sheet.Cells[rowIndex, 2].Value = row.ProductionSystem;
            sheet.Cells[rowIndex, 3].Value = row.MachineNumber;
            sheet.Cells[rowIndex, 4].Value = row.Ips;
            sheet.Cells[rowIndex, 5].Value = row.Name;
            sheet.Cells[rowIndex, 6].Value = row.PartQuantity;
            sheet.Cells[rowIndex, 7].Value = row.OneCCode;
            sheet.Cells[rowIndex, 8].Value = row.Nomenclature;
            sheet.Cells[rowIndex, 9].Value = row.UnitName;
            sheet.Cells[rowIndex, 10].Value = row.MaterialQuantity;
            sheet.Cells[rowIndex, 11].Value = row.DemandDate?.ToLocalTime().ToString("dd.MM.yyyy");
            sheet.Cells[rowIndex, 12].Value = FormatTotalPrice([row]);
            sheet.Cells[rowIndex, 13].Value = row.Comment;
            rowIndex++;
        }

        sheet.Cells.AutoFitColumns();
    }

    private static void FillGroupedRequest(ExcelWorksheet sheet, IEnumerable<MaterialRequestExportRow> rows)
    {
        var headers = new[] { "Заготовка", "Код", "Ед. изм", "Всего Купить", "Цена", "Комментарий" };
        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
        }

        var groupedRows = rows
            .Where(x => x.MaterialQuantity is > 0)
            .GroupBy(x => new
            {
                Code = string.IsNullOrWhiteSpace(x.OneCCode) ? x.Ips : x.OneCCode,
                x.UnitName
            })
            .OrderBy(x => x.Select(r => r.Nomenclature).FirstOrDefault())
            .ThenBy(x => x.Key.Code);

        sheet.OutLineSummaryBelow = false;
        var rowIndex = 2;
        foreach (var group in groupedRows)
        {
            var rowsInGroup = group.ToList();
            sheet.Cells[rowIndex, 1].Value = rowsInGroup.Select(x => x.Nomenclature).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
            sheet.Cells[rowIndex, 2].Value = group.Key.Code;
            sheet.Cells[rowIndex, 3].Value = group.Key.UnitName;
            sheet.Cells[rowIndex, 4].Value = rowsInGroup.Sum(x => x.PurchaseQuantity ?? 0m);
            sheet.Cells[rowIndex, 5].Value = FormatTotalPrice(rowsInGroup);
            sheet.Cells[rowIndex, 6].Value = string.Join("; ", rowsInGroup.Select(x => x.Comment).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
            sheet.Row(rowIndex).Collapsed = true;
            rowIndex++;

            foreach (var detail in rowsInGroup.OrderBy(x => x.Ips).ThenBy(x => x.Name))
            {
                sheet.Cells[rowIndex, 1].Value = $"{detail.Ips} {detail.Name}".Trim();
                sheet.Cells[rowIndex, 2].Value = group.Key.Code;
                sheet.Cells[rowIndex, 3].Value = detail.UnitName;
                sheet.Cells[rowIndex, 4].Value = detail.PurchaseQuantity ?? 0m;
                sheet.Cells[rowIndex, 5].Value = FormatTotalPrice([detail]);
                sheet.Cells[rowIndex, 6].Value = BuildGroupDetailComment(detail);
                sheet.Row(rowIndex).OutlineLevel = 1;
                sheet.Row(rowIndex).Hidden = true;
                rowIndex++;
            }
        }

        sheet.Cells.AutoFitColumns();
    }

    private static string BuildGroupDetailComment(MaterialRequestExportRow row)
    {
        var parts = new[]
        {
            string.IsNullOrWhiteSpace(row.ProductionSystem) ? null : $"ПС: {row.ProductionSystem}",
            string.IsNullOrWhiteSpace(row.MachineNumber) ? null : $"№ станка: {row.MachineNumber}",
            row.DemandDate is null ? null : $"Дата: {row.DemandDate.Value:dd.MM.yyyy}",
            string.IsNullOrWhiteSpace(row.Comment) ? null : row.Comment
        };
        return string.Join("; ", parts.Where(x => !string.IsNullOrWhiteSpace(x)));
    }

    private static void FillSummary(ExcelWorksheet sheet, IEnumerable<CalculationItem> items)
    {
        sheet.Cells[1, 1].Value = "Показатель";
        sheet.Cells[1, 2].Value = "Значение";
        sheet.Cells[2, 1].Value = "Позиций";
        sheet.Cells[2, 2].Value = items.Count();
        sheet.Cells[3, 1].Value = "К закупке";
        sheet.Cells[3, 2].Value = items.Count(x => x.PurchaseQuantity > 0);
        sheet.Cells[4, 1].Value = "Ошибок";
        sheet.Cells[4, 2].Value = items.Count(x => x.Status != CalculationStatus.Ok);
        sheet.Cells.AutoFitColumns();
    }

    private static void FillItems(ExcelWorksheet sheet, IEnumerable<CalculationItem> items)
    {
        var headers = new[] { "Заготовка", "Коды 1С", "Ед.", "Потребность", "Остаток", "Купить", "Статус", "Комментарий" };
        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
        }

        var row = 2;
        foreach (var item in items)
        {
            sheet.Cells[row, 1].Value = Clean(item.CanonicalName);
            sheet.Cells[row, 2].Value = item.OneCCodes;
            sheet.Cells[row, 3].Value = DisplayUnit(item.Unit);
            sheet.Cells[row, 4].Value = item.TotalRequired;
            sheet.Cells[row, 5].Value = item.TotalStock;
            sheet.Cells[row, 6].Value = item.PurchaseQuantity;
            sheet.Cells[row, 7].Value = DisplayStatus(item.Status);
            sheet.Cells[row, 8].Value = Clean(item.Comment);
            row++;
        }

        sheet.Cells.AutoFitColumns();
    }

    private static void FillDetails(ExcelWorksheet sheet, IEnumerable<CalculationItem> items)
    {
        var headers = new[] { "Заготовка", "IPS", "Деталь", "Потребность детали", "Потребность заготовки", "Ед." };
        for (var i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
        }

        var row = 2;
        foreach (var item in items)
        {
            foreach (var source in item.Sources)
            {
                sheet.Cells[row, 1].Value = Clean(item.CanonicalName);
                sheet.Cells[row, 2].Value = source.Ips;
                sheet.Cells[row, 3].Value = Clean(source.PartName);
                sheet.Cells[row, 4].Value = source.DemandQuantity;
                sheet.Cells[row, 5].Value = source.RequiredQuantity;
                sheet.Cells[row, 6].Value = DisplayUnit(source.Unit);
                row++;
            }
        }

        sheet.Cells.AutoFitColumns();
    }

    private async Task<Dictionary<string, decimal>> LoadWorkInProgressByIpsAsync(string[] ipsValues, CancellationToken cancellationToken)
    {
        if (ipsValues.Length == 0)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var latestSnapshotId = await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .ThenByDescending(x => x.ImportedAt)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync(cancellationToken);

        if (latestSnapshotId is null)
        {
            return new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        }

        var ipsKeys = ipsValues.Select(StockCodeNormalizer.NormalizeForComparison).Where(x => x.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stockItems = await dbContext.StockItems.AsNoTracking()
            .Where(x => x.StockSnapshotId == latestSnapshotId.Value && x.Unit == MeasurementUnit.Piece)
            .ToListAsync(cancellationToken);

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

    private static DateTime? BlankDemandDate(DateTime? detailDemandDate, int leadTimeDays) =>
        detailDemandDate?.AddDays(-Math.Max(0, leadTimeDays));

    private static string BuildBitrixRequestPath(string outputDirectory)
    {
        var date = DateTime.Now.ToString("dd.MM.yyyy");
        var prefix = $"Заявка на закуп заготовок ЦМО от {date}";
        var number = 1;
        string path;
        do
        {
            path = Path.Combine(outputDirectory, $"{prefix} {number}.xlsx");
            number++;
        }
        while (File.Exists(path));

        return path;
    }

    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

    private static MeasurementUnit? GetEffectiveUnit(PartBlankMap? map, CanonicalBlank? blank)
    {
        if (map is null)
        {
            return null;
        }

        return blank is not null && MeterBasedBlankTypes.Contains(blank.BlankType) && map.ConsumptionUnit == MeasurementUnit.Piece
            ? MeasurementUnit.Meter
            : map.ConsumptionUnit;
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
        BlankType.WeldingElement => "Сварное изделие",
        BlankType.Purchased => "Покупная",
        BlankType.Casting => "Литье",
        BlankType.Forging => "Поковка",
        BlankType.CustomBlank => "Прочее",
        BlankType.Unknown => "Не распознано",
        _ => "Не распознано"
    };

    private static string DisplayStatus(CalculationStatus status) => status switch
    {
        CalculationStatus.Ok => "ОК",
        CalculationStatus.MissingPart => "Нет детали",
        CalculationStatus.MissingBlankMapping => "Нет заготовки",
        CalculationStatus.MultipleActiveMappings => "Несколько связей",
        CalculationStatus.UnitMismatch => "Несовместимые единицы",
        _ => status.ToString()
    };

    private static string FormatPrice(OneCPriceItem? price)
    {
        if (price is null || price.Price <= 0)
        {
            return string.Empty;
        }

        return FormatPriceAmount(price.Price, FormatCurrency(price.Currency));
    }

    private static string FormatTotalPrice(IEnumerable<MaterialRequestExportRow> rows)
    {
        var rowsWithPrice = rows.Where(x => x.UnitPrice is > 0 && x.PurchaseQuantity is > 0).ToList();
        if (rowsWithPrice.Count == 0)
        {
            return string.Empty;
        }

        var amount = rowsWithPrice.Sum(x => x.PurchaseQuantity!.Value * x.UnitPrice!.Value);
        var currency = rowsWithPrice.Select(x => x.PriceCurrency).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;
        return FormatPriceAmount(amount, currency);
    }

    private static string FormatPriceAmount(decimal amount, string? currency)
    {
        if (amount <= 0)
        {
            return string.Empty;
        }

        var amountText = amount.ToString("0.####", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
        return string.IsNullOrWhiteSpace(currency) ? amountText : $"{amountText} {currency}";
    }

    private static string FormatCurrency(string? currency)
    {
        var value = Clean(currency);
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var upper = value.ToUpperInvariant();
        return upper is "RUB" or "RUR" or "643" || upper.Contains("РУБ", StringComparison.Ordinal)
            ? "₽"
            : value;
    }

    private static bool IsPreferredPriceType(string? value) =>
        Clean(value).Contains("Стоимость", StringComparison.OrdinalIgnoreCase);

    private static string BuildBlankSupplyRequirement(Part part)
    {
        var values = new List<string>(2);
        if (part.BlankSupplyRequiresHeatTreatment)
        {
            values.Add("ТО");
        }

        if (part.BlankSupplyRequiresLaserCutting)
        {
            values.Add("Лазерная резка");
        }

        return string.Join("; ", values);
    }

    private static string Clean(string? value)
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
                var bytes = System.Text.Encoding.GetEncoding(1251).GetBytes(text);
                var repaired = System.Text.Encoding.UTF8.GetString(bytes);
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

    private static long? TryGetOneTimeDemandItemId(string? comment)
    {
        const string marker = "DemandItemId=";
        var index = comment?.IndexOf(marker, StringComparison.OrdinalIgnoreCase) ?? -1;
        if (index < 0 || comment is null)
        {
            return null;
        }

        var value = comment[(index + marker.Length)..].Trim();
        return long.TryParse(value, out var id) ? id : null;
    }

    private sealed record MaterialRequestExportRow(
        int Number,
        string ProductionSystem,
        string MachineNumber,
        string Ips,
        string Name,
        decimal PartQuantity,
        decimal InProductionQuantity,
        string OneCCode,
        string Nomenclature,
        string UnitName,
        decimal? MaterialQuantity,
        decimal? PurchaseQuantity,
        DateTime? DemandDate,
        string Price,
        decimal? UnitPrice,
        string PriceCurrency,
        string Comment);
}
