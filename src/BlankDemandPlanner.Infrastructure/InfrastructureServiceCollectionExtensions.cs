using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Infrastructure.Backup;
using BlankDemandPlanner.Infrastructure.Excel;
using BlankDemandPlanner.Infrastructure.Export;
using BlankDemandPlanner.Infrastructure.OneC;
using Microsoft.Extensions.DependencyInjection;

namespace BlankDemandPlanner.Infrastructure;

public static class InfrastructureServiceCollectionExtensions
{
    public static IServiceCollection AddBlankDemandPlannerInfrastructure(this IServiceCollection services, string databasePath, string backupDirectory, int keepBackups)
    {
        services.AddScoped<IExcelImportProfile, OneCBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, MatchedOrderBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, RotationalBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, PipeBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, PlateBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, GuideBlankImportProfile>();
        services.AddScoped<IExcelImportProfile, DemandImportProfile>();
        services.AddScoped<IExcelImportProfile, StockImportProfile>();
        services.AddScoped<IExcelImportService, ExcelImportService>();
        services.AddScoped<IOneCStockSyncService, OneCStockSyncService>();
        services.AddScoped<IReportExportService, ReportExportService>();
        services.AddSingleton<IDatabaseBackupService>(provider =>
            new DatabaseBackupService(databasePath, backupDirectory, keepBackups, provider.GetRequiredService<Microsoft.Extensions.Logging.ILogger<DatabaseBackupService>>()));
        return services;
    }
}
