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
                db.Parts.Add(new Part { Ips = "000001", Name = "РџСЂРѕРІРµСЂРєР°" });
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

        var first = await service.NormalizeAsync("РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ Р“РћРЎРў 2590-2006 / Р“РћРЎРў 4543-2016", CancellationToken.None);
        var second = await service.NormalizeAsync("Р—Р°РіРѕС‚РѕРІРєР° РєСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ", CancellationToken.None);

        first.BlankType.Should().Be(BlankType.RoundBar);
        second.BlankType.Should().Be(BlankType.RoundBar);
        second.CanonicalKey.Should().Be(first.CanonicalKey);
    }

    [Fact]
    public async Task BlankNormalization_detects_forging()
    {
        var service = new BlankNormalizationService();

        var result = await service.NormalizeAsync("\u041F\u043E\u043A\u043E\u0432\u043A\u0430 \u0441\u0442\u0430\u043B\u044C 40\u0425", CancellationToken.None);

        result.BlankType.Should().Be(BlankType.Forging);
    }

    [Fact]
    public async Task I012_validation_checks_type_size_and_material()
    {
        await using var db = CreateDb();
        var validation = new I012ValidationService(db);
        var steel40 = "\u0421\u0442\u0430\u043B\u044C 40\u0425";
        var steel30 = "\u0421\u0442\u0430\u043B\u044C 30\u0425\u0413\u0421\u0410";
        db.I012CatalogEntries.Add(new I012CatalogEntry { BlankType = BlankType.RoundBar, SizeKey = "D60", MaterialKey = validation.BuildMaterialKey(steel40), Section = "Round" });
        db.I012CatalogEntries.Add(new I012CatalogEntry { BlankType = BlankType.RoundBar, SizeKey = "D80", MaterialKey = "20", Section = "Round" });
        await db.SaveChangesAsync();

        var allowed = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, steel40, null, null, 60, null, null, null, null, null, "Round D60", "ROUND|40X|D60", 1, []), CancellationToken.None);
        var materialMismatch = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, steel30, null, null, 60, null, null, null, null, null, "Round D60", "ROUND|30XGSA|D60", 1, []), CancellationToken.None);
        var sizeMismatch = await validation.ValidateAsync(new BlankNormalizationResult(BlankType.RoundBar, steel40, null, null, 100, null, null, null, null, null, "Round D100", "ROUND|40X|D100", 1, []), CancellationToken.None);

        allowed.Should().Be(I012Status.Allowed);
        materialMismatch.Should().Be(I012Status.MaterialNotAllowedForSize);
        sizeMismatch.Should().Be(I012Status.SizeNotAllowed);
    }

    [Fact]
    public async Task Duplicate_detection_finds_equal_physical_blank_keys()
    {
        await using var db = CreateDb();
        var first = new CanonicalBlank { CanonicalName = "Round D60 Steel 40X", CanonicalKey = "A", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40X", BaseUnit = MeasurementUnit.Meter };
        var second = new CanonicalBlank { CanonicalName = "Round D60 Steel 40X", CanonicalKey = "B", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40X", BaseUnit = MeasurementUnit.Meter };
        var sameMaterialDifferentSize = new CanonicalBlank { CanonicalName = "Round D80 Steel 40X", CanonicalKey = "C", BlankType = BlankType.RoundBar, DiameterMm = 80, Material = "40X", BaseUnit = MeasurementUnit.Meter };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = first, OneCCode = "UT001", SourceName = "A", NormalizedSourceName = "STEEL40X", Source = "test" },
            new BlankAlias { CanonicalBlank = second, OneCCode = "UT002", SourceName = "B", NormalizedSourceName = "STEEL40X", Source = "test" },
            new BlankAlias { CanonicalBlank = sameMaterialDifferentSize, OneCCode = "UT003", SourceName = "C", NormalizedSourceName = "STEEL40X", Source = "test" });
        await db.SaveChangesAsync();

        var candidates = await new DuplicateDetectionService(db).FindBlankAliasDuplicatesAsync(CancellationToken.None);

        candidates.Should().ContainSingle();
    }

    [Fact]
    public async Task Calculation_groups_many_parts_to_one_blank_and_subtracts_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ", CanonicalKey = "ROUND|40РҐ|D60", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў001", SourceName = "РљСЂСѓРі D60", NormalizedSourceName = "РљСЂСѓРі D60", Source = "test" };
        var p1 = AddPart(db, "100001", "Р’Р°Р»", blank);
        var p2 = AddPart(db, "100002", "Р’С‚СѓР»РєР°", blank);
        var p3 = AddPart(db, "100003", "РћРїРѕСЂР°", blank);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
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
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ", CanonicalKey = "ROUND|40РҐ|D60-WIP", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў001-WIP", SourceName = "РљСЂСѓРі D60", NormalizedSourceName = "РљСЂСѓРі D60", Source = "test" };
        var part = AddPart(db, "1720202", "РќР°РїСЂР°РІР»СЏСЋС‰Р°СЏ", blank);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
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
    public async Task Calculation_matches_ips_with_eleven_digit_1c_code()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D80", CanonicalKey = "ROUND|D80|IPS11", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў080", SourceName = "РљСЂСѓРі D80", NormalizedSourceName = "РљСЂСѓРі D80", Source = "test" };
        var part = AddPart(db, "1802132", "Р—Р°РіРѕС‚РѕРІРєР° РѕРїСЂР°РІРєРё", blank);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 1, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.ProductionWarehouseDisplayNames[0] },
            new StockItem { StockSnapshot = snapshot, OneCCode = "00001802132", SourceName = part.Name, Quantity = 4, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.WipWarehouseDisplayNames[0] });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(6);
        run.Items.Single().TotalStock.Should().Be(1);
        run.Items.Single().PurchaseQuantity.Should().Be(5);
    }

    [Fact]
    public async Task Calculation_uses_configured_warehouses_for_material_stock_and_work_in_progress()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D70", CanonicalKey = "ROUND|D70|SECTIONS", BaseUnit = MeasurementUnit.Piece };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў070", SourceName = "РљСЂСѓРі D70", NormalizedSourceName = "РљСЂСѓРі D70", Source = "test" };
        var part = AddPart(db, "170070", "РЎРµРєС†РёРѕРЅРЅР°СЏ РґРµС‚Р°Р»СЊ", blank);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 3, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.ProductionWarehouseDisplayNames[0] },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 2, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.ProductionWarehouseDisplayNames[1] },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 1, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.ProductionWarehouseDisplayNames[5] },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 100, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.WipWarehouseDisplayNames[0] },
            new StockItem { StockSnapshot = snapshot, OneCCode = "00000170070", SourceName = part.Name, Quantity = 4, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.WipWarehouseDisplayNames[0] },
            new StockItem { StockSnapshot = snapshot, OneCCode = "00000170070", SourceName = part.Name, Quantity = 2, Unit = MeasurementUnit.Piece, Warehouse = StockWarehouseRules.WipWarehouseDisplayNames[1] });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Single().TotalRequired.Should().Be(4);
        run.Items.Single().TotalStock.Should().Be(106);
        run.Items.Single().PurchaseQuantity.Should().Be(0);
    }

    [Fact]
    public async Task Calculation_subtracts_blank_stock_from_cmo_wip_warehouse_by_one_c_code()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Заготовка 85х1020х40мм сталь 30ХГСА",
            CanonicalKey = "SHEET|85X1020X40|30HGSA",
            BlankType = BlankType.Sheet,
            BaseUnit = MeasurementUnit.Piece
        };
        var alias = new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000022134",
            SourceName = "Заготовка 85х1020х40мм сталь 30ХГСА",
            NormalizedSourceName = "Заготовка 85х1020х40мм сталь 30ХГСА",
            Source = "test"
        };
        var part = AddPart(db, "1657364", "Направляющая задней бабки", blank, MeasurementUnit.Piece, 1m);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, ProductionSystem = "Р26666", Quantity = 2 },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, ProductionSystem = "Р26689", Quantity = 2 },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, ProductionSystem = "Р26672", Quantity = 2 },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, ProductionSystem = "Р26706", Quantity = 2 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            BlankAlias = alias,
            OneCCode = alias.OneCCode,
            SourceName = alias.SourceName,
            Quantity = 6,
            Unit = MeasurementUnit.Piece,
            Warehouse = StockWarehouseRules.WipWarehouseDisplayNames[0]
        });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(8);
        run.Items.Single().TotalStock.Should().Be(6);
        run.Items.Single().PurchaseQuantity.Should().Be(2);
    }

    [Fact]
    public async Task Calculation_sums_fractional_meter_consumption_and_subtracts_fractional_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUND|40X|D60-FRACTION",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "UTM060", SourceName = "РљСЂСѓРі D60", NormalizedSourceName = "РљСЂСѓРі D60", Source = "test" };
        var p1 = AddPart(db, "200001", "Р”РµС‚Р°Р»СЊ 0,1", blank, MeasurementUnit.Meter, 0.1m);
        var p2 = AddPart(db, "200002", "Р”РµС‚Р°Р»СЊ 0,35", blank, MeasurementUnit.Meter, 0.35m);
        var p3 = AddPart(db, "200003", "Р”РµС‚Р°Р»СЊ 0,45", blank, MeasurementUnit.Meter, 0.45m);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = p1, Ips = p1.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p2, Ips = p2.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p3, Ips = p3.Ips, Quantity = 1 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 0.2m, Unit = MeasurementUnit.Meter });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(0.91m);
        run.Items.Single().TotalStock.Should().Be(0.2m);
        run.Items.Single().PurchaseQuantity.Should().Be(0.71m);
    }

    [Fact]
    public async Task Calculation_adds_five_mm_cut_width_between_meter_blanks_and_subtracts_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D140 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUND|40X|D140-CUT",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "UTD140", SourceName = "РљСЂСѓРі D140 РЎС‚Р°Р»СЊ 40РҐ", NormalizedSourceName = "РљСЂСѓРі D140 РЎС‚Р°Р»СЊ 40РҐ", Source = "test" };
        var p1 = AddPart(db, "240001", "Р”РµС‚Р°Р»СЊ 0,1", blank, MeasurementUnit.Meter, 0.1m);
        var p2 = AddPart(db, "240002", "Р”РµС‚Р°Р»СЊ 0,1 РґСѓР±Р»СЊ", blank, MeasurementUnit.Meter, 0.1m);
        var p3 = AddPart(db, "240003", "Р”РµС‚Р°Р»СЊ 0,03", blank, MeasurementUnit.Meter, 0.03m);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = p1, Ips = p1.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p2, Ips = p2.Ips, Quantity = 1 },
            new DemandItem { DemandBatch = batch, Part = p3, Ips = p3.Ips, Quantity = 1 });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 0.1m, Unit = MeasurementUnit.Meter });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id, snapshot.Id), CancellationToken.None);

        run.Items.Should().ContainSingle();
        run.Items.Single().TotalRequired.Should().Be(0.24m);
        run.Items.Single().TotalStock.Should().Be(0.1m);
        run.Items.Single().PurchaseQuantity.Should().Be(0.14m);
        run.Items.Single().Sources.Select(x => x.RequiredQuantity).Should().Equal(0.1m, 0.1m, 0.03m);
    }

    [Fact]
    public async Task Calculation_ignores_duplicate_active_maps_for_same_blank_and_unit()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D60", CanonicalKey = Guid.NewGuid().ToString() };
        var part = new Part { Ips = "300001", Name = "РќР°РїСЂР°РІР»СЏСЋС‰Р°СЏ" };
        db.PartBlankMaps.AddRange(
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = true, Source = "test" },
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = false, Source = "duplicate" },
            new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, IsPrimary = false, Source = "duplicate" });
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
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
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D60", CanonicalKey = Guid.NewGuid().ToString() };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў001", SourceName = "РљСЂСѓРі", NormalizedSourceName = "РљСЂСѓРі", Source = "test" };
        var part = AddPart(db, "100001", "Р’Р°Р»", blank);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
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
        var part = new Part { Ips = "100001", Name = "Р’Р°Р»" };
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 10 });
        await db.SaveChangesAsync();

        var run = await CreateCalculation(db).CalculateAsync(new CalculationOptions(batch.Id), CancellationToken.None);

        run.Items.Single().Status.Should().Be(CalculationStatus.MissingBlankMapping);
    }

    [Fact]
    public async Task Calculation_marks_unit_mismatch()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "РљСЂСѓРі D60", CanonicalKey = Guid.NewGuid().ToString() };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "РЈРў001", SourceName = "РљСЂСѓРі", NormalizedSourceName = "РљСЂСѓРі", Source = "test" };
        var part = AddPart(db, "100001", "Р’Р°Р»", blank, MeasurementUnit.Piece);
        var batch = new DemandBatch { Name = "РџРѕС‚СЂРµР±РЅРѕСЃС‚СЊ" };
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
