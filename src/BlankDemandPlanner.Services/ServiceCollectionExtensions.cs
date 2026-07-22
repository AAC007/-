using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Services.Calculation;
using BlankDemandPlanner.Services.Duplicates;
using BlankDemandPlanner.Services.Normalization;
using BlankDemandPlanner.Services.Validation;
using Microsoft.Extensions.DependencyInjection;

namespace BlankDemandPlanner.Services;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddBlankDemandPlannerServices(this IServiceCollection services)
    {
        services.AddScoped<IBlankNormalizationService, BlankNormalizationService>();
        services.AddScoped<II012ValidationService, I012ValidationService>();
        services.AddScoped<IBlankDemandCalculationService, BlankDemandCalculationService>();
        services.AddScoped<IUnitConversionService, UnitConversionService>();
        services.AddScoped<IDuplicateDetectionService, DuplicateDetectionService>();
        return services;
    }
}
