using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;

namespace BlankDemandPlanner.Services.Calculation;

public sealed class UnitConversionService : IUnitConversionService
{
    public bool CanSubtract(MeasurementUnit requiredUnit, MeasurementUnit stockUnit) => requiredUnit == stockUnit;

    public decimal Convert(decimal quantity, MeasurementUnit fromUnit, MeasurementUnit toUnit, CanonicalBlank? blank)
    {
        if (fromUnit == toUnit)
        {
            return quantity;
        }

        throw new InvalidOperationException($"Нет правила пересчета {fromUnit} -> {toUnit} для {blank?.CanonicalName ?? "заготовки"}.");
    }
}
