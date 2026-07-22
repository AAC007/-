using BlankDemandPlanner.Core.Interfaces;

namespace BlankDemandPlanner.Infrastructure.Excel;

public abstract class ExcelImportProfileBase : IExcelImportProfile
{
    public abstract string ProfileType { get; }
    protected abstract IReadOnlyDictionary<string, string[]> Aliases { get; }

    public virtual bool CanHandle(string sheetName, IReadOnlyList<string> headers)
    {
        var joined = string.Join(' ', headers).ToUpperInvariant();
        return Aliases.Values.SelectMany(x => x).Any(a => joined.Contains(a.ToUpperInvariant(), StringComparison.Ordinal));
    }

    public IReadOnlyDictionary<string, string> AutoMapColumns(IReadOnlyList<string> headers)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in Aliases)
        {
            var match = headers.FirstOrDefault(header => field.Value.Any(alias => Normalize(header).Equals(Normalize(alias), StringComparison.OrdinalIgnoreCase)))
                ?? headers.FirstOrDefault(header => field.Value.Any(alias => Normalize(header).Contains(Normalize(alias), StringComparison.OrdinalIgnoreCase)));
            if (!string.IsNullOrWhiteSpace(match))
            {
                result[field.Key] = match;
            }
        }

        return result;
    }

    private static string Normalize(string value) => value.ToUpperInvariant().Replace(" ", string.Empty, StringComparison.Ordinal).Replace("_", string.Empty, StringComparison.Ordinal);
}

public sealed class OneCBlankImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "OneCBlank";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = new Dictionary<string, string[]>
    {
        ["OneCCode"] = ["Код", "Артикул", "УТ", "Номенклатура"],
        ["Name"] = ["Наименование", "Наименование заготовки"]
    };
}

public sealed class DemandImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "Demand";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = new Dictionary<string, string[]>
    {
        ["Ips"] = ["Код IPS", "IPS"],
        ["Project"] = ["Проект"],
        ["SerialNumber"] = ["Серийный №", "Серийный N", "Серийный номер"],
        ["PartName"] = ["Наименование", "НАИМЕНОВАНИЕ"],
        ["Unit"] = ["Ед. изм.", "Ед", "Единица"],
        ["Quantity"] = ["Кол-во", "Количество", "ШТ"],
        ["DemandDate"] = ["Дата", "Дата потребности", "Дата потребн."],
        ["ProductionSystem"] = ["ПС", "PS", "Номер станка", "Входит в станок"]
    };
}

public sealed class StockImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "Stock";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = new Dictionary<string, string[]>
    {
        ["OneCCode"] = ["Код", "УТ", "Код заготовки", "Код позиции 1С"],
        ["Name"] = ["Наименование", "Наименование заготовки", "Наименование позиции / заготовки"],
        ["Quantity"] = ["Количество", "Кол-во", "Остаток", "Остаток/кол-во"],
        ["Unit"] = ["Ед", "Единица"],
        ["Warehouse"] = ["Склад"]
    };
}

public sealed class MatchedOrderBlankImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "MatchedOrderBlankLibrary";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = new Dictionary<string, string[]>
    {
        ["Ips"] = ["IPS"],
        ["Designation"] = ["Обозначение"],
        ["PartName"] = ["Наименование", "Наименование детали"],
        ["Dimensions"] = ["размер заготовки"],
        ["BlankTypeName"] = ["Вид заготовки"],
        ["OneCCode"] = ["УТ код"],
        ["BlankName"] = ["наименование Заготовки по 1с", "Наименование по 1С УТ", "Заготовка"],
        ["Material"] = ["Материал"],
        ["Unit"] = ["Ед. изм.", "Ед"],
        ["Quantity"] = ["потребность в материале", "Потребность", "Количество"],
        ["BlankLeadTimeDays"] = ["Срок заготовки, дней", "Срок, дней"]
    };

    public override bool CanHandle(string sheetName, IReadOnlyList<string> headers) =>
        sheetName.Equals("Итог", StringComparison.OrdinalIgnoreCase) ||
        sheetName.Equals("Библиотека", StringComparison.OrdinalIgnoreCase);
}

public class RotationalBlankImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "RotationalBlank";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = ManufacturingAliases;
    public override bool CanHandle(string sheetName, IReadOnlyList<string> headers) => sheetName.Contains("вращ", StringComparison.OrdinalIgnoreCase) || base.CanHandle(sheetName, headers);
    internal static IReadOnlyDictionary<string, string[]> ManufacturingAliases { get; } = new Dictionary<string, string[]>
    {
        ["Ips"] = ["IPS"],
        ["Designation"] = ["Обозначение"],
        ["PartName"] = ["Наименование детали", "НАИМЕНОВАНИЕ"],
        ["Material"] = ["Марка материала"],
        ["Sortament"] = ["Сортамент"],
        ["Dimensions"] = ["размер заготовки", "Размер"],
        ["BlankTypeName"] = ["Вид заготовки"],
        ["RequiredDiameter"] = ["Требуемый диаметр"],
        ["RequiredLength"] = ["Требуемая длина"],
        ["OneCCode"] = ["УТ", "УТ код"],
        ["BlankName"] = ["Наименование заготовки", "наименование Заготовки по 1с"],
        ["Unit"] = ["Ед. изм.", "Ед"],
        ["Quantity"] = ["потребность в материале", "Потребность"],
        ["BlankLeadTimeDays"] = ["Срок заготовки, дней", "Срок, дней"]
    };
}

public sealed class PipeBlankImportProfile : RotationalBlankImportProfile
{
    public override string ProfileType => "PipeBlank";
    public override bool CanHandle(string sheetName, IReadOnlyList<string> headers) => sheetName.Contains("труб", StringComparison.OrdinalIgnoreCase) || base.CanHandle(sheetName, headers);
}

public sealed class PlateBlankImportProfile : RotationalBlankImportProfile
{
    public override string ProfileType => "PlateBlank";
    public override bool CanHandle(string sheetName, IReadOnlyList<string> headers) => sheetName.Contains("квадрат", StringComparison.OrdinalIgnoreCase) || sheetName.Contains("лист", StringComparison.OrdinalIgnoreCase) || sheetName.Contains("плит", StringComparison.OrdinalIgnoreCase) || base.CanHandle(sheetName, headers);
}

public sealed class GuideBlankImportProfile : ExcelImportProfileBase
{
    public override string ProfileType => "GuideBlank";
    protected override IReadOnlyDictionary<string, string[]> Aliases { get; } = new Dictionary<string, string[]>
    {
        ["Ips"] = ["IPS"],
        ["PartName"] = ["НАИМЕНОВАНИЕ"],
        ["Dimensions"] = ["ВхШхД"],
        ["OneCCode"] = ["УТ"],
        ["BlankName"] = ["НАИМЕНОВАНИЕ ЗАГОТОВКИ"],
        ["Quantity"] = ["ШТ"]
    };

    public override bool CanHandle(string sheetName, IReadOnlyList<string> headers) => sheetName.Contains("направ", StringComparison.OrdinalIgnoreCase) || base.CanHandle(sheetName, headers);
}
