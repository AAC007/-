using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.Infrastructure.Excel;
using BlankDemandPlanner.Services.Normalization;
using BlankDemandPlanner.Services.Validation;
using BlankDemandPlanner.UI.ViewModels;
using FluentAssertions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeOpenXml;

namespace BlankDemandPlanner.Tests;

public sealed class RealFileImportTests
{
    public RealFileImportTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public async Task One_c_import_reads_rows_without_headers_by_ut_code()
    {
        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_import_no_headers_{Guid.NewGuid():N}.db");
        var filePath = Path.Combine(Path.GetTempPath(), $"bdp_onec_no_headers_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                sheet.Cells[1, 1].Value = "УТ000022800";
                sheet.Cells[1, 2].Value = "Заготовка 60х174х970мм Сталь 30ХГСА";
                await package.SaveAsAsync(new FileInfo(filePath));
            }

            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            var report = await importer.ImportOneCBlanksAsync(filePath, null, CancellationToken.None);

            report.AddedRows.Should().Be(1);
            var alias = await db.BlankAliases.SingleAsync();
            alias.OneCCode.Should().Be("УТ000022800");
            alias.SourceName.Should().Contain("30ХГСА");
        }
        finally
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }

    [Fact]
    public async Task Real_enterprise_files_import_without_entity_save_errors()
    {
        var root = FindWorkspaceRoot();
        var oneCFile = FindOptionalFile(root, "*заготовок*1с*.xlsx");
        var libraryFile = Path.Combine(root, "Исходные данные для изготовления_заказа заготовок.xlsx");
        if (oneCFile is null || !File.Exists(libraryFile))
        {
            return;
        }

        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_import_{Guid.NewGuid():N}.db");
        try
        {
            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var normalizer = new BlankNormalizationService();
            var validator = new I012ValidationService(db);
            var importer = new ExcelImportService(db, normalizer, validator, NullLogger<ExcelImportService>.Instance);

            var oneCReport = await importer.ImportOneCBlanksAsync(oneCFile, null, CancellationToken.None);
            var libraryReport = await importer.ImportManufacturingBlankLibraryAsync(libraryFile, null, CancellationToken.None);

            oneCReport.ErrorRows.Should().Be(0);
            (oneCReport.AddedRows + oneCReport.UpdatedRows).Should().BeGreaterThan(0);
            libraryReport.ErrorRows.Should().Be(0);
            (await db.BlankAliases.CountAsync()).Should().BeGreaterThan(0);
            (await db.Parts.CountAsync()).Should().BeGreaterThan(0);
        }
        finally
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }
        }
    }

    [Fact]
    public async Task Reimporting_same_demand_file_refreshes_display_fields_without_duplicate_batch()
    {
        var root = FindWorkspaceRoot();
        var demandFile = Path.Combine(root, "Потребность.xlsx");
        if (!File.Exists(demandFile))
        {
            return;
        }

        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_demand_refresh_{Guid.NewGuid():N}.db");
        try
        {
            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            var firstReport = await importer.ImportDemandAsync(demandFile, null, CancellationToken.None);
            firstReport.AddedRows.Should().BeGreaterThan(0);
            var initialItemCount = await db.DemandItems.CountAsync();

            var item = await db.DemandItems.SingleAsync(x => x.Ips == "1814870");
            item.Project = null;
            item.SerialNumber = null;
            item.ProductionSystem = string.Empty;
            item.SourcePartName = "Замена привода люнета";
            item.Unit = null;
            await db.SaveChangesAsync();

            var refreshReport = await importer.ImportDemandAsync(demandFile, null, CancellationToken.None);

            refreshReport.AddedRows.Should().Be(0);
            refreshReport.UpdatedRows.Should().Be(initialItemCount);
            (await db.DemandBatches.CountAsync()).Should().Be(1);
            (await db.DemandItems.CountAsync()).Should().Be(initialItemCount);

            var refreshed = await db.DemandItems.AsNoTracking().SingleAsync(x => x.Ips == "1814870");
            refreshed.Project.Should().Be("ПС200 ОДК-ПМ");
            refreshed.SerialNumber.Should().Be("Р24500-ВД");
            refreshed.ProductionSystem.Should().Be("Р24500");
            refreshed.SourcePartName.Should().StartWith("КУ260305.ПС200-63-P012-02.001");
            refreshed.Unit.Should().Be("шт");
            refreshed.Quantity.Should().Be(1);
            refreshed.DemandDate.Should().Be(new DateTime(2026, 7, 15));
        }
        finally
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }
        }
    }

    [Fact]
    public async Task Reimporting_same_demand_file_moves_that_batch_to_latest()
    {
        var root = FindWorkspaceRoot();
        var demandFile = Path.Combine(root, "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ.xlsx");
        if (!File.Exists(demandFile))
        {
            return;
        }

        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_demand_latest_{Guid.NewGuid():N}.db");
        try
        {
            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            await importer.ImportDemandAsync(demandFile, null, CancellationToken.None);
            var imported = await db.DemandBatches.SingleAsync();
            imported.ImportedAt = DateTime.UtcNow.AddDays(-2);
            db.DemandBatches.Add(new DemandBatch { Name = "Более новая другая потребность", ImportedAt = DateTime.UtcNow.AddDays(-1) });
            await db.SaveChangesAsync();

            await importer.ImportDemandAsync(demandFile, null, CancellationToken.None);

            var latest = await db.DemandBatches.OrderByDescending(x => x.ImportedAt).FirstAsync();
            latest.Id.Should().Be(imported.Id);
        }
        finally
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }
        }
    }

    [Fact]
    public async Task Desktop_demand_files_refresh_visible_demand_after_sequential_imports()
    {
        var desktop = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");
        var files = Directory.Exists(desktop)
            ? Directory.GetFiles(desktop, "*Потребность*.xlsx").OrderBy(x => x).Take(3).ToArray()
            : [];
        if (files.Length < 2)
        {
            return;
        }

        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_desktop_demand_sequence_{Guid.NewGuid():N}.db");
        try
        {
            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            foreach (var file in files)
            {
                await importer.ImportDemandAsync(file, null, CancellationToken.None);
                db.ChangeTracker.Clear();
                var viewModel = new DemandViewModel(db);
                await viewModel.LoadCommand.ExecuteAsync(null);
                viewModel.Rows.Should().NotBeEmpty($"после импорта {Path.GetFileName(file)} список потребности должен обновиться");
                viewModel.BatchText.Should().Contain(Path.GetFileNameWithoutExtension(file));
            }

            var repeatedFile = files[0];
            await importer.ImportDemandAsync(repeatedFile, null, CancellationToken.None);
            db.ChangeTracker.Clear();
            var refreshedViewModel = new DemandViewModel(db);
            await refreshedViewModel.LoadCommand.ExecuteAsync(null);

            refreshedViewModel.Rows.Should().NotBeEmpty();
            refreshedViewModel.BatchText.Should().Contain(Path.GetFileNameWithoutExtension(repeatedFile));
        }
        finally
        {
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }
        }
    }

    private static string? FindOptionalFile(string root, string pattern) =>
        Directory.GetFiles(root, pattern, SearchOption.TopDirectoryOnly).FirstOrDefault();

    private static string FindWorkspaceRoot()
    {
        var current = new DirectoryInfo(AppContext.BaseDirectory);
        while (current is not null)
        {
            if (File.Exists(Path.Combine(current.FullName, "BlankDemandPlanner.sln")))
            {
                return current.FullName;
            }

            current = current.Parent;
        }

        throw new DirectoryNotFoundException("Workspace root was not found.");
    }
}
