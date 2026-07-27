using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

namespace BlankDemandPlanner.Services.Calculation;

public sealed class BlankDemandCalculationService(
    BlankDemandPlannerDbContext dbContext,
    IUnitConversionService unitConversionService,
    ILogger<BlankDemandCalculationService> logger) : IBlankDemandCalculationService
{
    public async Task<CalculationRun> CalculateAsync(CalculationOptions options, CancellationToken cancellationToken)
    {
        logger.LogInformation("Calculation start for demand batch {DemandBatchId}", options.DemandBatchId);

        var stockSnapshotId = options.StockSnapshotId ?? await dbContext.StockSnapshots.AsNoTracking()
            .OrderByDescending(x => x.SnapshotDate)
            .Select(x => (long?)x.Id)
            .FirstOrDefaultAsync(cancellationToken);

        var demandItems = await dbContext.DemandItems.AsNoTracking()
            .Where(x => x.DemandBatchId == options.DemandBatchId)
            .ToListAsync(cancellationToken);

        var ipsValues = demandItems.Select(x => x.Ips).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var parts = await dbContext.Parts.AsNoTracking()
            .Include(x => x.BlankMaps.Where(m => m.IsActive))
            .ThenInclude(x => x.CanonicalBlank)
            .Where(x => ipsValues.Contains(x.Ips))
            .ToDictionaryAsync(x => x.Ips, StringComparer.OrdinalIgnoreCase, cancellationToken);

        var blankIds = parts.Values.SelectMany(p => p.BlankMaps).Select(m => m.CanonicalBlankId).Distinct().ToArray();
        var aliases = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => blankIds.Contains(x.CanonicalBlankId) && x.IsActive)
            .ToListAsync(cancellationToken);

        var aliasIds = aliases.Select(x => x.Id).ToArray();
        var aliasKeys = aliases.Select(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var stocks = stockSnapshotId is null
            ? []
            : await dbContext.StockItems.AsNoTracking()
                .Where(x => x.StockSnapshotId == stockSnapshotId)
                .ToListAsync(cancellationToken);
        stocks = stocks
            .Where(x => StockWarehouseRules.IsProductionWarehouse(x.Warehouse))
            .Where(x => x.BlankAliasId is not null && aliasIds.Contains(x.BlankAliasId.Value) ||
                aliasKeys.Contains(StockCodeNormalizer.NormalizeForComparison(x.OneCCode)))
            .ToList();

        var ipsKeys = ipsValues.Select(StockCodeNormalizer.NormalizeForComparison).Where(x => x.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var workInProgressItems = stockSnapshotId is null
            ? new List<StockItem>()
            : await dbContext.StockItems.AsNoTracking()
                .Where(x => x.StockSnapshotId == stockSnapshotId && x.Unit == MeasurementUnit.Piece)
                .ToListAsync(cancellationToken);
        workInProgressItems = workInProgressItems.Where(x => StockWarehouseRules.IsCmoWipWarehouse(x.Warehouse)).ToList();

        var workInProgressByIps = workInProgressItems
            .Where(x => ipsKeys.Contains(StockCodeNormalizer.NormalizeForComparison(x.OneCCode)))
            .GroupBy(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.Sum(i => i.Quantity), StringComparer.OrdinalIgnoreCase);

        var run = new CalculationRun
        {
            DemandBatchId = options.DemandBatchId,
            StockSnapshotId = stockSnapshotId,
            Comment = options.Comment
        };

        var groups = new Dictionary<(long? BlankId, MeasurementUnit Unit, CalculationStatus Status, string Key), CalculationItem>();

        foreach (var demand in demandItems)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var effectiveDemandQuantity = ApplyWorkInProgress(demand.Ips, demand.Quantity, workInProgressByIps);
            if (effectiveDemandQuantity <= 0)
            {
                continue;
            }

            if (!parts.TryGetValue(demand.Ips, out var part))
            {
                AddProblem(run, groups, demand, CalculationStatus.MissingPart, "Деталь не найдена в библиотеке");
                continue;
            }

            var activeMaps = part.BlankMaps
                .Where(x => x.IsActive)
                .GroupBy(x => new { x.CanonicalBlankId, x.ConsumptionUnit })
                .Select(x => x.OrderByDescending(m => m.IsPrimary).ThenByDescending(m => m.UpdatedAt).First())
                .ToList();
            if (activeMaps.Count == 0)
            {
                AddProblem(run, groups, demand, CalculationStatus.MissingBlankMapping, "Нет активной связи IPS -> заготовка");
                continue;
            }

            if (activeMaps.Count(x => x.IsPrimary) > 1)
            {
                AddProblem(run, groups, demand, CalculationStatus.MultipleActiveMappings, "Найдено несколько активных основных связей");
                continue;
            }

            foreach (var map in activeMaps)
            {
                var required = effectiveDemandQuantity * map.ConsumptionQuantity;
                var requiredWithLoss = required * (1 + map.LossPercent / 100m);
                var blank = map.CanonicalBlank;
                var key = (BlankId: (long?)map.CanonicalBlankId, Unit: map.ConsumptionUnit, Status: CalculationStatus.Ok, Key: $"B:{map.CanonicalBlankId}:{map.ConsumptionUnit}");

                if (!groups.TryGetValue(key, out var item))
                {
                    var blankAliases = aliases.Where(x => x.CanonicalBlankId == map.CanonicalBlankId).ToList();
                    var blankAliasIds = blankAliases.Select(x => x.Id).ToHashSet();
                    var blankAliasKeys = blankAliases.Select(x => StockCodeNormalizer.NormalizeForComparison(x.OneCCode)).ToHashSet(StringComparer.OrdinalIgnoreCase);
                    var relatedStock = stocks
                        .Where(s => s.BlankAliasId is not null && blankAliasIds.Contains(s.BlankAliasId.Value) ||
                            blankAliasKeys.Contains(StockCodeNormalizer.NormalizeForComparison(s.OneCCode)))
                        .ToList();
                    var stockUnitMismatch = relatedStock.Any(s => !unitConversionService.CanSubtract(map.ConsumptionUnit, s.Unit));
                    var stockQuantity = stockUnitMismatch ? 0m : relatedStock.Sum(s => unitConversionService.Convert(s.Quantity, s.Unit, map.ConsumptionUnit, blank));

                    item = new CalculationItem
                    {
                        CanonicalBlankId = map.CanonicalBlankId,
                        CanonicalName = blank?.CanonicalName ?? $"Заготовка #{map.CanonicalBlankId}",
                        PrimaryOneCCode = blankAliases.FirstOrDefault()?.OneCCode,
                        OneCCodes = string.Join(", ", blankAliases.Select(x => x.OneCCode).Distinct()),
                        Unit = map.ConsumptionUnit,
                        TotalStock = stockQuantity,
                        Status = stockUnitMismatch ? CalculationStatus.UnitMismatch : CalculationStatus.Ok,
                        Comment = stockUnitMismatch ? "Остатки имеют несовместимые единицы измерения" : null
                    };
                    groups[key] = item;
                    run.Items.Add(item);
                }

                item.TotalRequired += requiredWithLoss;
                item.Sources.Add(new CalculationItemSource
                {
                    Ips = demand.Ips,
                    PartName = part.Name,
                    DemandQuantity = effectiveDemandQuantity,
                    RequiredQuantity = requiredWithLoss,
                    Unit = map.ConsumptionUnit
                });
            }
        }

        foreach (var item in run.Items)
        {
            item.PurchaseQuantity = item.Status == CalculationStatus.UnitMismatch ? 0 : Math.Max(0, item.TotalRequired - item.TotalStock);
        }

        run.FinishedAt = DateTime.UtcNow;
        dbContext.CalculationRuns.Add(run);
        await dbContext.SaveChangesAsync(cancellationToken);
        logger.LogInformation("Calculation finish for run {CalculationRunId}", run.Id);
        return run;
    }

    private static decimal ApplyWorkInProgress(string ips, decimal demandQuantity, IDictionary<string, decimal> workInProgressByIps)
    {
        var key = StockCodeNormalizer.NormalizeForComparison(ips);
        if (!workInProgressByIps.TryGetValue(key, out var inProduction) || inProduction <= 0)
        {
            return demandQuantity;
        }

        var used = Math.Min(demandQuantity, inProduction);
        workInProgressByIps[key] = inProduction - used;
        return demandQuantity - used;
    }

    private static void AddProblem(CalculationRun run, IDictionary<(long? BlankId, MeasurementUnit Unit, CalculationStatus Status, string Key), CalculationItem> groups, DemandItem demand, CalculationStatus status, string comment)
    {
        var key = (BlankId: (long?)null, Unit: MeasurementUnit.Piece, Status: status, Key: $"{status}:{demand.Ips}");
        if (!groups.TryGetValue(key, out var item))
        {
            item = new CalculationItem
            {
                CanonicalName = "Новая номенклатура",
                Unit = MeasurementUnit.Piece,
                Status = status,
                Comment = comment
            };
            groups[key] = item;
            run.Items.Add(item);
        }

        item.TotalRequired += demand.Quantity;
        item.PurchaseQuantity = item.TotalRequired;
        item.Sources.Add(new CalculationItemSource
        {
            Ips = demand.Ips,
            PartName = demand.SourcePartName ?? string.Empty,
            DemandQuantity = demand.Quantity,
            RequiredQuantity = demand.Quantity,
            Unit = MeasurementUnit.Piece
        });
    }
}
