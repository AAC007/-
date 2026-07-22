using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace BlankDemandPlanner.Data;

public static class DataServiceCollectionExtensions
{
    public static IServiceCollection AddBlankDemandPlannerData(this IServiceCollection services, string databasePath)
    {
        services.AddDbContext<BlankDemandPlannerDbContext>(options =>
            options.UseSqlite($"Data Source={databasePath}"));

        return services;
    }
}
