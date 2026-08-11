namespace BlankDemandPlanner.Core.Models;

public static class StockWarehouseRules
{
    public static readonly string[] WipWarehouseFilters =
    [
        "44",
        "Детали МУ в обработке на стороне"
    ];

    public static readonly string[] ProductionWarehouseFilters =
    [
        "УТ0000037",
        "УТ0001081",
        "УТ0001062",
        "УТ0000981",
        "УТ0001018",
        "УТ0001019"
    ];

    public static readonly string[] WipWarehouseDisplayNames =
    [
        "44 секция НЗП (незавершенное производство)",
        "Детали МУ в обработке на стороне"
    ];

    public static readonly string[] ProductionWarehouseDisplayNames =
    [
        "1 секция (ПЗМЦ) только для производственных целей",
        "Покраска ТМЦ",
        "ТМЦ по проекту ФРП",
        "ТМЦ для проверки ОТК",
        "ТМЦ для временного хранения ЦСТС",
        "ТМЦ для временного хранения ЦМО"
    ];

    public static bool IsCmoWipWarehouse(string? warehouse)
    {
        var text = Normalize(warehouse);
        return text.Length == 0 ||
            WipWarehouseDisplayNames.Any(name => MatchesWarehouse(text, name)) ||
            WipWarehouseFilters.Any(filter => MatchesWarehouse(text, filter)) ||
            text.StartsWith("44 секция", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsProductionWarehouse(string? warehouse)
    {
        var text = Normalize(warehouse);
        return text.Length == 0 ||
            ProductionWarehouseDisplayNames.Any(name => MatchesWarehouse(text, name)) ||
            ProductionWarehouseFilters.Any(filter => MatchesWarehouse(text, filter)) ||
            text.StartsWith("1 секция", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Основной склад", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsProductionLaunchMaterialWarehouse(string? warehouse) =>
        IsProductionWarehouse(warehouse) || IsCmoWipWarehouse(warehouse);

    public static string WipWarehouseSummary => string.Join("; ", WipWarehouseDisplayNames);
    public static string ProductionWarehouseSummary => string.Join("; ", ProductionWarehouseDisplayNames);

    private static string Normalize(string? value) => (value ?? string.Empty).Trim();
    private static bool MatchesWarehouse(string text, string value) =>
        text.Equals(value, StringComparison.OrdinalIgnoreCase) ||
        text.Contains(value, StringComparison.OrdinalIgnoreCase);
}
