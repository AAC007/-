using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.Services.Calculation;
using BlankDemandPlanner.Services.Duplicates;
using BlankDemandPlanner.Services.Normalization;
using BlankDemandPlanner.Services.Validation;
using FluentAssertions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;

namespace BlankDemandPlanner.Tests;

public sealed class DomainServiceTests
{
    [Fact]
    public async Task Database_migration_creates_sqlite_schema()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bdp_{Guid.NewGuid():N}.db");
        try
        {
            var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
                .UseSqlite($"Data Source={path};Pooling=False")
                .Options;
            await using (var db = new BlankDemandPlannerDbContext(options))
            {
                await db.Database.MigrateAsync();
                db.Parts.Add(new Part { Ips = "000001", Name = "Проверка" });
                await db.SaveChangesAsync();

                (await db.Parts.CountAsync()).Should().Be(1);
            }
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public async Task BlankNormalization_round_bar_names_build_same_key()
    {
        var service = new BlankNormalizationService();

        var first = await service.NormalizeAsync("Круг D60 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", CancellationToken.None);
        var second = await service.NormalizeAsync("Заготовка круг D60 Сталь 40Х", CancellationToken.None);

        first.BlankType.Should().Be(BlankType.RoundBar);
        second.BlankType.Should().Be(BlankType.RoundBar);
        second.CanonicalKey.Should().Be(first.CanonicalKey);
    }

    [Fact]
    public async Task BlankNormalization_detects_forging()
    {
        var service = new BlankNormalizationService();

        var result = await service.NormalizeAsync("Поковка сталь 40Х", CancellationToken.None);

        result.BlankType.Should().Be(BlankType.Forging);
    }

    [Fact]
    public async Task I012_validation_checks_type_size_and_material()
    {
        await using var db = CreateDb();
        db.I012CatalogEntries.Add(new I012CatalogEntry { BlankType = BlankType.RoundBar, SizeKey = "D60", MaterialKey = "40Х", Section = "Круг" });
        db.I012CatalogEntries.Add(new I012CatalogEntry { BlankType = BlankType.RoundBar, SizeKey = "D80", MaterialKey = "20", Section = "Круг" });
        await db.SaveChangesAsync();

        var validation = new I012ValidationService(db);
        var allowed = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, "Сталь 40Х", null, null, 60, null, null, null, null, null, "Круг D60", "ROUND|40Х|D60", 1, []), CancellationToken.None);
        var materialMismatch = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, "Сталь 30ХГСА", null, null, 60, null, null, null, null, null, "Круг D60", "ROUND|30ХГСА|D60", 1, []), CancellationToken.None);
        var sizeMismatch = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, "Сталь 40Х", null, null, 100, null, null, null, null, null, "Круг D100", "ROUND|40Х|D100", 1, []), CancellationToken.None);

        allowed.Should().Be(I012Status.Allowed);
        materialMismatch.Should().Be(I012Status.MaterialNotAllowedForSize);
        sizeMismatch.Should().Be(I012Status.SizeNotAllowed);
    }

    [Fact]
    public async Task Duplicate_detection_finds_equal_normalized_aliases()
    {
        await using var db = CreateDb();
        var first = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "A" };
        var second = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "B" };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = first, OneCCode = "УТ001", SourceName = "A", NormalizedSourceName = "Круг D60 Сталь 40Х", Source = "test" },
            new BlankAlias { CanonicalBlank = second, OneCCode = "УТ002", SourceName = "B", NormalizedSourceName = "Круг D60 Сталь 40Х", Source = "test" });
        await db.SaveChangesAsync();

        var candidates = await new DuplicateDetectionService(db).FindBlankAliasDuplicatesAsync(CancellationToken.None);

        candidates.Should().ContainSingle();
    }

    [Fact]
    public async Task Calculation_groups_many_parts_to_one_blank_and_subtracts_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "ROUND|40Х|D60", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ001", SourceName = "Круг D60", NormalizedSourceName = "Круг D60", Source = "test" };
        var p1 = AddPart(db, "100001", "Вал", blank);
        var p2 = AddPart(db, "100002", "Втулка", blank);
        var p3 = AddPart(db, "100003", "Опора", blank);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = p1, Ips = p1.Ips, Quantity = 10 },
            new DemandItem { DemandBatch = batch, Part = p2, Ips = p2.Ips, Quantity = 15 },
            new DemandItem { DemandBatch = batch, Part = p3, Ips = p3.Ips, Quantity = 8 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 8, Unit = MeasurementUnit.Piece });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(33);
        run.Items.Single().TotalStock.Should().Be(8);
        run.Items.Single().PurchaseQuantity.Should().Be(25);
    }

    [Fact]
    public async Task Calculation_subtracts_work_in_progress_parts_from_material_demand()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "ROUND|40Х|D60-WIP", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ001-WIP", SourceName = "Круг D60", NormalizedSourceName = "Круг D60", Source = "test" };
        var part = AddPart(db, "1720202", "Направляющая", blank);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 2, Unit = MeasurementUnit.Piece },
            new StockItem { StockSnapshot = snapshot, OneCCode = "000" + part.Ips, SourceName = part.Name, Quantity = 4, Unit = MeasurementUnit.Piece });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(6);
        run.Items.Single().TotalStock.Should().Be(2);
        run.Items.Single().PurchaseQuantity.Should().Be(4);
    }

    [Fact]
    public async Task Calculation_sums_fractional_meter_consumption_and_subtracts_fractional_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60 Сталь 40Х",
            CanonicalKey = "ROUND|40X|D60-FRACTION",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "UTM060", SourceName = "Круг D60", NormalizedSourceName = "Круг D60", Source = "test" };
        var p1 = AddPart(db, "200001", "Деталь 0,1", blank, MeasurementUnit.Meter, 0.1m);
        var p2 = AddPart(db, "200002", "Деталь 0,35", blank, MeasurementUnit.Meter, 0.35m);
        var p3 = AddPart(db, "200003", "Деталь 0,45", blank, MeasurementUnit.Meter, 0.45m);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = p1, Ips = p1.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p2, Ips = p2.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p3, Ips = p3.Ips, Quantity = 1 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 0.2m, Unit = MeasurementUnit.Meter });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(0.9m);
        run.Items.Single().TotalStock.Should().Be(0.2m);
        run.Items.Single().PurchaseQuantity.Should().Be(0.7m);
    }

    [Fact]
    public async Task Calculation_ignores_duplicate_active_maps_for_same_blank_and_unit()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D60", CanonicalKey = Guid.NewGuid().ToString() };
        var part = new Part { Ips = "300001", Name = "Направляющая" };
        db.PartBlankMaps.AddRange(
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = true, Source = "test" },
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = false, Source = "duplicate" },
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = false, Source = "duplicate" });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 5 });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(5);
    }


    [Fact]
    public async Task Calculation_purchase_is_zero_when_stock_exceeds_demand()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D60", CanonicalKey = Guid.NewGuid().ToString() };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ001", SourceName = "Круг", NormalizedSourceName = "Круг", Source = "test" };
        var part = AddPart(db, "100001", "Вал", blank);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 20, Unit = MeasurementUnit.Piece });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Single().PurchaseQuantity.Should().Be(0);
    }

    [Fact]
    public async Task Calculation_marks_missing_blank_mapping()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "100001", Name = "Вал" };
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id), CancellationToken.None);

        run.Items.Single().Status.Should().Be(CalculationStatus.MissingBlankMapping);
    }

    [Fact]
    public async Task Calculation_marks_unit_mismatch()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D60", CanonicalKey = Guid.NewGuid().ToString() };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ001", SourceName = "Круг", NormalizedSourceName = "Круг", Source = "test" };
        var part = AddPart(db, "100001", "Вал", blank, MeasurementUnit.Piece);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 5, Unit = MeasurementUnit.Meter });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Single().Status.Should().Be(CalculationStatus.UnitMismatch);
    }

    private static BlankDemandPlannerDbContext CreateDb()
    {
        var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
            .UseInMemoryDatabase(Guid.NewGuid().ToString())
            .Options;
        return new BlankDemandPlannerDbContext(options);
    }

    private static Part AddPart(BlankDemandPlannerDbContext db, string ips, string name, CanonicalBlank blank, MeasurementUnit unit = MeasurementUnit.Piece, decimal consumptionQuantity = 1m)
    {
        var part = new Part { Ips = ips, Name = name };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = consumptionQuantity, ConsumptionUnit = unit, Source = "test" });
        return part;
    }

    private static BlankDemandCalculationService CreateCalculation(BlankDemandPlannerDbContext db) =>
        new(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
}
