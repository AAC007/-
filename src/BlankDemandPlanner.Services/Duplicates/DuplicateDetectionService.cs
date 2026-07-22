using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using Microsoft.EntityFrameworkCore;

namespace BlankDemandPlanner.Services.Duplicates;

public sealed class DuplicateDetectionService(BlankDemandPlannerDbContext dbContext) : IDuplicateDetectionService
{
    public async Task<IReadOnlyList<DuplicateBlankCandidate>> FindBlankAliasDuplicatesAsync(CancellationToken cancellationToken)
    {
        var aliases = await dbContext.BlankAliases.AsNoTracking()
            .Where(x => x.IsActive)
            .Select(x => new { x.Id, x.OneCCode, x.NormalizedSourceName, x.CanonicalBlankId })
            .ToListAsync(cancellationToken);

        var result = new List<DuplicateBlankCandidate>();
        foreach (var group in aliases.GroupBy(x => Simplify(x.NormalizedSourceName)).Where(g => g.Count() > 1))
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

    private static string Simplify(string value) =>
        new(value.ToUpperInvariant().Where(c => !char.IsWhiteSpace(c) && c != '-' && c != '_' && c != '/').ToArray());
}
