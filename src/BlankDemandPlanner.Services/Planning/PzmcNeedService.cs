using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Services.Planning;

public sealed class PzmcNeedService(IPzmcProductionApiClient apiClient) : IPzmcNeedService
{
    public async Task<PzmcNeedResult> GetDeficitAsync(
        PzmcNeedFilters filters,
        bool forceRefresh,
        CancellationToken cancellationToken)
    {
        var snapshot = await apiClient.LoadSnapshotAsync(forceRefresh, cancellationToken);
        var rows = CalculateDeficit(snapshot);
        var filtered = ApplyFilters(rows, filters)
            .OrderBy(x => x.NeedDate)
            .ThenByDescending(x => x.Priority)
            .ThenBy(x => x.Ips)
            .ToArray();

        return new PzmcNeedResult(
            filtered,
            snapshot.UpdatedAt,
            snapshot.Source,
            $"Позиций дефицита: {filtered.Length}; источник: {snapshot.Source}; обновлено {snapshot.UpdatedAt:dd.MM.yyyy HH:mm}");
    }

    private IReadOnlyList<PzmcNeedRow> CalculateDeficit(PzmcProductionSnapshot snapshot)
    {
        var products = snapshot.Products.ToDictionary(x => x.Id);
        var ipsByObjectId = snapshot.IpsObjects
            .Where(x => x.ObjectId > 0)
            .GroupBy(x => x.ObjectId)
            .ToDictionary(x => x.Key, x => x.First());
        var ipsByIpsId = snapshot.IpsObjects
            .Where(x => x.ObjectIpsId > 0)
            .GroupBy(x => x.ObjectIpsId)
            .ToDictionary(x => x.Key, x => x.First());
        var groups = snapshot.ProductGroups.ToDictionary(x => x.Id);

        var productionWarehouses = snapshot.Warehouses.Where(x => x.IsUsedInProduction).ToArray();
        if (productionWarehouses.Length == 0)
        {
            productionWarehouses = snapshot.Warehouses.ToArray();
        }

        var stock = productionWarehouses
            .SelectMany(x => x.Balance)
            .Where(x => x.ObjectId > 0)
            .GroupBy(x => StockKey(x.ObjectId, x.Unit))
            .ToDictionary(
                x => x.Key,
                x => x.Sum(GetAvailableStock));
        var plannedReceipts = snapshot.PlannedReceipts
            .Where(x => x.ObjectId > 0 && x.Quantity > 0 && x.ActualDate is null)
            .GroupBy(x => StockKey(x.ObjectId, x.Unit))
            .ToDictionary(
                x => x.Key,
                x => x.OrderBy(p => p.PlannedDate)
                    .Select(p => new ReceiptBalance(p.PlannedDate, (decimal)p.Quantity))
                    .ToList());
        var analogsByParent = snapshot.Analogs
            .Where(x => x.ParentObjectId > 0 && x.ObjectId > 0)
            .GroupBy(x => x.ParentObjectId)
            .ToDictionary(
                x => x.Key,
                x => x.OrderByDescending(a => a.IsPriority).ThenBy(a => a.Id).ToArray());

        var rows = new List<PzmcNeedRow>();
        foreach (var need in snapshot.Needs
                     .Where(x => x.Quantity > 0)
                     .OrderBy(x => x.NeedDate)
                     .ThenByDescending(x => products.GetValueOrDefault(x.ProductId)?.Priority ?? 0)
                     .ThenBy(x => x.ObjectIpsId))
        {
            products.TryGetValue(need.ProductId, out var product);
            var spec = ResolveSpec(need, ipsByObjectId, ipsByIpsId);
            var objectId = need.ObjectId > 0 ? need.ObjectId : spec?.ObjectId ?? 0;
            var unit = NormalizeUnit(need.Unit);
            var required = Math.Max(0m, (decimal)(need.Quantity - need.QuantityExists));
            if (required <= 0)
            {
                continue;
            }

            var directStock = ConsumeStock(stock, StockKey(objectId, unit), required);
            var remaining = required - directStock;
            var analogStock = 0m;
            var applicableAnalogs = analogsByParent.GetValueOrDefault(objectId) ?? [];
            foreach (var analog in applicableAnalogs.Where(x => IsApplicable(x, need, product)))
            {
                var consumed = ConsumeStock(stock, StockKey(analog.ObjectId, unit), remaining);
                analogStock += consumed;
                remaining -= consumed;
                if (remaining <= 0)
                {
                    break;
                }
            }

            var planned = ConsumeReceipts(plannedReceipts, StockKey(objectId, unit), need.NeedDate, remaining);
            remaining -= planned;
            foreach (var analog in applicableAnalogs.Where(x => IsApplicable(x, need, product)))
            {
                if (remaining <= 0)
                {
                    break;
                }

                var consumed = ConsumeReceipts(plannedReceipts, StockKey(analog.ObjectId, unit), need.NeedDate, remaining);
                planned += consumed;
                remaining -= consumed;
            }

            if (remaining <= 0)
            {
                continue;
            }

            var groupId = need.ProductGroupId ?? spec?.ProductGroupId;
            groups.TryGetValue(groupId ?? 0, out var group);
            var groupText = Join(group?.Designation, group?.Name);
            var hasOrder = need.OrderId is not null ||
                           need.OrderNumber is not null ||
                           !string.IsNullOrWhiteSpace(need.OrderLink) ||
                           planned > 0;

            rows.Add(new PzmcNeedRow(
                need.Id,
                First(product?.ProjectName, "Без проекта"),
                product?.SerialId ?? string.Empty,
                product?.Priority ?? 0,
                need.ObjectIpsId > 0 ? need.ObjectIpsId : spec?.ObjectIpsId ?? 0,
                First(need.ObjectDesignation, spec?.Designation, need.Designation),
                First(need.ObjectName, spec?.Name, need.Designation),
                groupText,
                First(need.ManufacturingType, spec?.ManufacturingType),
                First(need.ManufacturingMethod, spec?.ManufacturingMethod),
                First(need.Manager, spec?.Manager),
                First(need.Supplier, spec?.Supplier),
                need.Unit,
                required,
                directStock,
                analogStock,
                planned,
                remaining,
                need.NeedDate,
                need.PdfExists || spec?.PdfExists == true,
                hasOrder,
                BuildOrderText(need),
                First(need.OperationName, need.OperationId?.ToString())));
        }

        return rows;
    }

    private IEnumerable<PzmcNeedRow> ApplyFilters(IEnumerable<PzmcNeedRow> rows, PzmcNeedFilters filters)
    {
        if (filters.CmoOnly)
        {
            rows = rows.Where(IsCmoRow);
        }

        return rows.Where(x =>
            Contains(x.Project, filters.Project) &&
            Contains(x.ProductGroup, filters.ProductGroup) &&
            Contains(x.ManufacturingType, filters.ManufacturingType) &&
            Contains(x.ManufacturingMethod, filters.ManufacturingMethod) &&
            Contains(x.Manager, filters.Manager) &&
            Contains(x.Supplier, filters.Supplier) &&
            (filters.DueFrom is null || x.NeedDate.Date >= filters.DueFrom.Value.Date) &&
            (filters.DueTo is null || x.NeedDate.Date <= filters.DueTo.Value.Date) &&
            (filters.HasPdf is null || x.HasPdf == filters.HasPdf.Value) &&
            (filters.HasOrder is null || x.HasOrder == filters.HasOrder.Value));
    }

    private bool IsCmoRow(PzmcNeedRow row)
    {
        var search = string.Join(' ', row.ProductGroup, row.ManufacturingType, row.ManufacturingMethod);
        if (apiClient.CmoGroups.Count > 0)
        {
            return apiClient.CmoGroups.Any(x => Contains(search, x));
        }

        string[] defaults = ["цмо", "заготов", "мех", "изготов", "поков", "лить", "покуп"];
        return defaults.Any(x => Contains(search, x));
    }

    private static PzmcSpecIpsObject? ResolveSpec(
        PzmcProductNeed need,
        IReadOnlyDictionary<int, PzmcSpecIpsObject> byObjectId,
        IReadOnlyDictionary<int, PzmcSpecIpsObject> byIpsId)
    {
        if (need.ObjectId > 0 && byObjectId.TryGetValue(need.ObjectId, out var byObject))
        {
            return byObject;
        }

        return need.ObjectIpsId > 0 && byIpsId.TryGetValue(need.ObjectIpsId, out var byIps) ? byIps : null;
    }

    private static decimal GetAvailableStock(PzmcSkladBalance balance)
    {
        if (balance.RemainingQuantity != 0 || balance.Quantity == 0)
        {
            return Math.Max(0m, (decimal)balance.RemainingQuantity);
        }

        return Math.Max(0m, (decimal)(balance.Quantity - balance.ReservedQuantity));
    }

    private static decimal ConsumeStock(IDictionary<string, decimal> stock, string key, decimal requested)
    {
        if (requested <= 0 || !stock.TryGetValue(key, out var available) || available <= 0)
        {
            return 0;
        }

        var consumed = Math.Min(requested, available);
        stock[key] = available - consumed;
        return consumed;
    }

    private static decimal ConsumeReceipts(
        IReadOnlyDictionary<string, List<ReceiptBalance>> receipts,
        string key,
        DateTime needDate,
        decimal requested)
    {
        if (requested <= 0 || !receipts.TryGetValue(key, out var values))
        {
            return 0;
        }

        var consumed = 0m;
        foreach (var receipt in values.Where(x => x.Date.Date <= needDate.Date && x.Quantity > 0))
        {
            var part = Math.Min(requested - consumed, receipt.Quantity);
            receipt.Quantity -= part;
            consumed += part;
            if (consumed >= requested)
            {
                break;
            }
        }

        return consumed;
    }

    private static bool IsApplicable(PzmcAnalog analog, PzmcProductNeed need, PzmcProduct? product) =>
        (analog.DateStart is null || need.NeedDate.Date >= analog.DateStart.Value.Date) &&
        (analog.DateEnd is null || need.NeedDate.Date <= analog.DateEnd.Value.Date) &&
        (analog.ProductIds.Length == 0 || analog.ProductIds.Contains(need.ProductId)) &&
        (analog.TechnologyIds.Length == 0 ||
         (need.TechnologyId is not null && analog.TechnologyIds.Contains(need.TechnologyId.Value)) ||
         (product is not null && analog.TechnologyIds.Contains(product.TechnologyId)));

    private static string StockKey(int objectId, string unit) => $"{objectId}|{NormalizeUnit(unit)}";
    private static string NormalizeUnit(string? unit) => (unit ?? string.Empty).Trim().ToUpperInvariant();
    private static bool Contains(string? value, string? filter) =>
        string.IsNullOrWhiteSpace(filter) ||
        (value ?? string.Empty).Contains(filter.Trim(), StringComparison.OrdinalIgnoreCase);
    private static string First(params string?[] values) =>
        values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;
    private static string Join(params string?[] values) =>
        string.Join(" | ", values.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x!.Trim()).Distinct(StringComparer.OrdinalIgnoreCase));
    private static string BuildOrderText(PzmcProductNeed need) =>
        Join(need.OrderType, need.OrderNumber?.ToString(), need.OrderLink);

    private sealed record ReceiptBalance(DateTime Date, decimal InitialQuantity)
    {
        public decimal Quantity { get; set; } = InitialQuantity;
    }
}
