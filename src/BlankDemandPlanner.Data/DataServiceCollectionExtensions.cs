using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.DependencyInjection;

namespace BlankDemandPlanner.Data;

public static class DataServiceCollectionExtensions
{
    public static IServiceCollection AddBlankDemandPlannerData(this IServiceCollection services, string databasePath)
    {
        services.AddDbContext<BlankDemandPlannerDbContext>(options =>
            options.UseSqlite($"Data Source={databasePath};Cache=Shared"));
        services.AddDbContextFactory<BlankDemandPlannerDbContext>(options =>
            options.UseSqlite($"Data Source={databasePath};Cache=Shared"));

        return services;
    }
}
