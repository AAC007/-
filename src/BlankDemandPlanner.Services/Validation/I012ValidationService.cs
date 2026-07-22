using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;

namespace BlankDemandPlanner.Services.Validation;

public sealed class I012ValidationService(BlankDemandPlannerDbContext dbContext) : II012ValidationService
{
    public async Task<I012Status> ValidateAsync(BlankNormalizationResult blank, CancellationToken cancellationToken)
    {
        if (blank.BlankType == BlankType.Unknown || string.IsNullOrWhiteSpace(blank.Material))
        {
            return I012Status.Unrecognized;
        }

        var sizeKey = BuildSizeKey(blank);
        var materialKey = BuildMaterialKey(blank.Material);
        if (string.IsNullOrWhiteSpace(sizeKey))
        {
            return I012Status.SizeNotAllowed;
        }

        var typeExists = await dbContext.I012CatalogEntries.AsNoTracking()
            .AnyAsync(x => x.BlankType == blank.BlankType, cancellationToken);
        if (!typeExists)
        {
            return I012Status.TypeNotFound;
        }

        var sizeExists = await dbContext.I012CatalogEntries.AsNoTracking()
            .AnyAsync(x => x.BlankType == blank.BlankType && x.SizeKey == sizeKey, cancellationToken);
        if (!sizeExists)
        {
            return I012Status.SizeNotAllowed;
        }

        var allowed = await dbContext.I012CatalogEntries.AsNoTracking()
            .AnyAsync(x => x.BlankType == blank.BlankType && x.SizeKey == sizeKey && x.MaterialKey == materialKey, cancellationToken);

        return allowed ? I012Status.Allowed : I012Status.MaterialNotAllowedForSize;
    }

    public string BuildSizeKey(BlankNormalizationResult blank)
    {
        if (blank.Diameter is not null) return $"D{Format(blank.Diameter.Value)}";
        if (blank.Width is not null && blank.Height is not null && blank.Thickness is not null) return $"{Format(blank.Width.Value)}x{Format(blank.Height.Value)}x{Format(blank.Thickness.Value)}";
        if (blank.Width is not null && blank.Height is not null) return $"{Format(blank.Width.Value)}x{Format(blank.Height.Value)}";
        if (blank.Width is not null) return Format(blank.Width.Value);
        if (blank.Thickness is not null) return $"S{Format(blank.Thickness.Value)}";
        return string.Empty;
    }

    public string BuildMaterialKey(string? material) =>
        string.IsNullOrWhiteSpace(material)
            ? string.Empty
            : material.ToUpperInvariant().Replace("СТАЛЬ", string.Empty, StringComparison.Ordinal).Replace("НЕРЖ", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);

    private static string Format(decimal value) => value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
}
