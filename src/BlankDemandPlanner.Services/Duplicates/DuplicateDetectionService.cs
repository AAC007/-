using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;
using System.Globalization;

namespace BlankDemandPlanner.Services.Duplicates;

public sealed class DuplicateDetectionService(BlankDemandPlannerDbContext dbContext) : IDuplicateDetectionService
{
    public async Task<IReadOnlyList<DuplicateBlankCandidate>> FindBlankAliasDuplicatesAsync(CancellationToken cancellationToken)
    {
        var aliases = await dbContext.BlankAliases.AsNoTracking()
            .Include(x => x.CanonicalBlank)
            .Where(x => x.IsActive && x.CanonicalBlank != null && x.CanonicalBlank.IsActive)
            .ToListAsync(cancellationToken);

        var result = new List<DuplicateBlankCandidate>();
        foreach (var group in aliases.GroupBy(x => BuildKey(x.CanonicalBlank!, x.SourceName)).Where(g => !string.IsNullOrWhiteSpace(g.Key) && g.Count() > 1))
        {
            var items = group.ToList();
            for (var i = 0; i < items.Count; i++)
            {
                for (var j = i + 1; j < items.Count; j++)
                {
                    if (items[i].CanonicalBlankId != items[j].CanonicalBlankId)
                    {
                        result.Add(new DuplicateBlankCandidate(items[i].Id, items[j].Id, items[i].OneCCode, items[j].OneCCode, "Совпадает нормализованное наименование", 0.95));
                    }
                }
            }
        }

        return result;
    }

    private static string BuildKey(Core.Entities.CanonicalBlank blank, string sourceName)
    {
        var material = Simplify(blank.Material ?? string.Empty);
        var materialGost = Simplify(blank.MaterialGost ?? string.Empty);
        var profileGost = Simplify(blank.ProfileGost ?? string.Empty);
        var dimensions = new[]
        {
            ("D", blank.DiameterMm),
            ("W", blank.WidthMm),
            ("H", blank.HeightMm),
            ("T", blank.ThicknessMm),
            ("S", blank.WallThicknessMm),
            ("L", blank.LengthMm)
        }
            .Where(x => x.Item2 is not null)
            .Select(x => $"{x.Item1}{x.Item2!.Value.ToString("0.####", CultureInfo.InvariantCulture)}")
            .ToArray();

        if (dimensions.Length == 0)
        {
            var name = Simplify(string.IsNullOrWhiteSpace(sourceName) ? blank.CanonicalName : sourceName);
            return name.Length == 0 || name == material
                ? string.Empty
                : string.Join("|", (int)blank.BlankType, (int)blank.BaseUnit, name, material, materialGost, profileGost);
        }

        return string.Join("|", (int)blank.BlankType, (int)blank.BaseUnit, material, materialGost, profileGost, string.Join(";", dimensions));
    }

    private static string Simplify(string value) =>
        new(value.ToUpperInvariant().Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_' && c != '/').ToArray());
}
