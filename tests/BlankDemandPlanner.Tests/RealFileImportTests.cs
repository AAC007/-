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
    public async Task Msk_headerless_library_file_updates_part_blank_library()
    {
        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_msk_headerless_{Guid.NewGuid():N}.db");
        var filePath = Path.Combine(Path.GetTempPath(), $"Данные из МСК_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                sheet.Cells[1, 1].Value = "1354749";
                sheet.Cells[1, 2].Value = "07.053.00.002";
                sheet.Cells[1, 3].Value = "Диск тормозной";
                sheet.Cells[1, 4].Value = "Круг";
                sheet.Cells[1, 5].Value = "D330";
                sheet.Cells[1, 6].Value = "Сталь 40Х";
                sheet.Cells[1, 7].Value = "ГОСТ 4543-2016";
                sheet.Cells[1, 8].Value = "ГОСТ 2590-2006";
                sheet.Cells[1, 9].Value = 0.025m;
                sheet.Cells[1, 10].Value = "пог. м";
                sheet.Cells[1, 11].Value = "УТ000007475";
                sheet.Cells[1, 12].Value = "Круг D330 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016";
                await package.SaveAsAsync(new FileInfo(filePath));
            }

            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            var report = await importer.ImportManufacturingBlankLibraryAsync(filePath, null, CancellationToken.None);

            report.ReadRows.Should().Be(1);
            report.SkippedRows.Should().Be(0);
            var part = await db.Parts.Include(x => x.BlankMaps).SingleAsync();
            part.Ips.Should().Be("1354749");
            part.Designation.Should().Be("07.053.00.002");
            part.Name.Should().Be("Диск тормозной");
            var map = part.BlankMaps.Single();
            map.ConsumptionQuantity.Should().Be(0.025m);
            map.ConsumptionUnit.Should().Be(BlankDemandPlanner.Core.Enums.MeasurementUnit.Meter);
            var alias = await db.BlankAliases.SingleAsync();
            alias.OneCCode.Should().Be("УТ000007475");
            alias.SourceName.Should().Contain("D330");
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
    public async Task Msk_headerless_library_file_reads_full_blank_name_format()
    {
        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_msk_full_blank_{Guid.NewGuid():N}.db");
        var filePath = Path.Combine(Path.GetTempPath(), $"Данные из МСК_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                sheet.Cells[1, 1].Value = "1320756";
                sheet.Cells[1, 2].Value = "КПТ-04.00.001А01";
                sheet.Cells[1, 3].Value = "Втулка внутренняя";
                sheet.Cells[1, 4].Value = "Круг";
                sheet.Cells[1, 5].Value = "Круг D80 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016";
                sheet.Cells[1, 6].Value = "Сталь 40Х";
                sheet.Cells[1, 7].Value = "УТ000022657";
                sheet.Cells[1, 8].Value = "ГОСТ 4543-2016";
                sheet.Cells[1, 9].Value = "ГОСТ 2590-2006";
                sheet.Cells[1, 10].Value = "пог. м";
                sheet.Cells[1, 11].Value = 0.065m;
                await package.SaveAsAsync(new FileInfo(filePath));
            }

            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            var report = await importer.ImportManufacturingBlankLibraryAsync(filePath, null, CancellationToken.None);

            report.ReadRows.Should().Be(1);
            report.SkippedRows.Should().Be(0);
            var part = await db.Parts.Include(x => x.BlankMaps).SingleAsync();
            part.Ips.Should().Be("1320756");
            part.Designation.Should().Be("КПТ-04.00.001А01");
            part.Name.Should().Be("Втулка внутренняя");
            var map = part.BlankMaps.Single();
            map.ConsumptionQuantity.Should().Be(0.065m);
            map.ConsumptionUnit.Should().Be(BlankDemandPlanner.Core.Enums.MeasurementUnit.Meter);
            var alias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync();
            alias.OneCCode.Should().Be("УТ000022657");
            alias.SourceName.Should().Be("Круг D80 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016");
            alias.CanonicalBlank!.DiameterMm.Should().Be(80);
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
    public async Task Msk_headerless_import_restores_archived_library_part()
    {
        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_msk_restore_archived_{Guid.NewGuid():N}.db");
        var filePath = Path.Combine(Path.GetTempPath(), $"Данные из МСК_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                sheet.Cells[1, 1].Value = "1354749";
                sheet.Cells[1, 2].Value = "07.053.00.002";
                sheet.Cells[1, 3].Value = "Диск тормозной";
                sheet.Cells[1, 4].Value = "Круг";
                sheet.Cells[1, 5].Value = "D330";
                sheet.Cells[1, 6].Value = "Сталь 40Х";
                sheet.Cells[1, 9].Value = 0.025m;
                sheet.Cells[1, 10].Value = "пог. м";
                sheet.Cells[1, 11].Value = "УТ000007475";
                sheet.Cells[1, 12].Value = "Круг D330 Сталь 40Х";
                await package.SaveAsAsync(new FileInfo(filePath));
            }

            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();
            db.Parts.Add(new Part { Ips = "1354749", Name = "Удаленная деталь", Source = "Потребность [ARCHIVED_LIBRARY] test" });
            await db.SaveChangesAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            await importer.ImportManufacturingBlankLibraryAsync(filePath, null, CancellationToken.None);
            var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
            await viewModel.LoadAsync();

            viewModel.Rows.Should().ContainSingle(x => x.Ips == "1354749");
            var restoredPart = await db.Parts.AsNoTracking().SingleAsync(x => x.Ips == "1354749");
            restoredPart.Source.Should().NotContain("[ARCHIVED_LIBRARY]");
            restoredPart.Name.Should().Be("Диск тормозной");
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
    public async Task Msk_headerless_import_restores_archived_nsi_blank_and_alias()
    {
        var dbPath = Path.Combine(Path.GetTempPath(), $"bdp_msk_restore_nsi_{Guid.NewGuid():N}.db");
        var filePath = Path.Combine(Path.GetTempPath(), $"Данные из МСК_{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Лист1");
                sheet.Cells[1, 1].Value = "1354749";
                sheet.Cells[1, 2].Value = "07.053.00.002";
                sheet.Cells[1, 3].Value = "Диск тормозной";
                sheet.Cells[1, 4].Value = "Круг";
                sheet.Cells[1, 5].Value = "D330";
                sheet.Cells[1, 6].Value = "Сталь 40Х";
                sheet.Cells[1, 9].Value = 0.025m;
                sheet.Cells[1, 10].Value = "пог. м";
                sheet.Cells[1, 11].Value = "УТ000007475";
                sheet.Cells[1, 12].Value = "Круг D330 Сталь 40Х";
                await package.SaveAsAsync(new FileInfo(filePath));
            }

            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={dbPath};Pooling=False")
                .Options;
            await using var db = new BlankDemandPlannerDbContext(options);
            await db.Database.MigrateAsync();
            var archivedBlank = new CanonicalBlank
            {
                CanonicalName = "Круг D330 Сталь 40Х",
                CanonicalKey = "ROUNDBAR|D330|M:СТАЛЬ 40Х",
                BlankType = BlankDemandPlanner.Core.Enums.BlankType.RoundBar,
                DiameterMm = 330,
                Material = "Сталь 40Х",
                BaseUnit = BlankDemandPlanner.Core.Enums.MeasurementUnit.Meter,
                IsActive = false
            };
            db.BlankAliases.Add(new BlankAlias
            {
                CanonicalBlank = archivedBlank,
                OneCCode = "УТ000007475",
                SourceName = "Круг D330 Сталь 40Х",
                NormalizedSourceName = "Круг D330 Сталь 40Х",
                Source = "Архив НСИ",
                IsActive = false
            });
            await db.SaveChangesAsync();

            var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
            await importer.ImportManufacturingBlankLibraryAsync(filePath, null, CancellationToken.None);

            var alias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync(x => x.OneCCode == "УТ000007475");
            alias.IsActive.Should().BeTrue();
            alias.CanonicalBlank!.IsActive.Should().BeTrue();
            var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
            await viewModel.LoadAsync();
            viewModel.Rows.Should().ContainSingle(x => x.Ips == "1354749" && x.OneCCode == "УТ000007475");
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
