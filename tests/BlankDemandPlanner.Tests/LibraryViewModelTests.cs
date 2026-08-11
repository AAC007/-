using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using BlankDemandPlanner.Infrastructure.Excel;
using BlankDemandPlanner.Infrastructure.Export;
using BlankDemandPlanner.Services.Calculation;
using BlankDemandPlanner.Services.Normalization;
using BlankDemandPlanner.Services.Validation;
using BlankDemandPlanner.UI.Services;
using BlankDemandPlanner.UI.ViewModels;
using FluentAssertions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.Abstractions;
using OfficeOpenXml;

namespace BlankDemandPlanner.Tests;

public sealed class LibraryViewModelTests
{
    [Fact]
    public void Ui_search_matches_case_and_cyrillic_latin_designation_letters()
    {
        UiSearchText.Contains("01-P014-01.120-T50.50.SS", "01-р014").Should().BeTrue();
        UiSearchText.Contains("Направляющая Z верхняя", "напра").Should().BeTrue();
        UiSearchText.Contains("направляющая Z верхняя", "Напра").Should().BeTrue();
        UiSearchText.Contains("1414993", "00001414993").Should().BeTrue();
        UiSearchText.ContainsAnyField("04.067.02.003 Гайка плоская", "04.067.02.003", "Гайка плоская").Should().BeTrue();
        UiSearchText.Contains("Круг D60 Сталь 20", "Круг д60 сталь 20").Should().BeTrue();
    }

    [Fact]
    public async Task Library_search_matches_combined_designation_and_name()
    {
        await using var db = CreateDb();
        db.Parts.Add(new Part { Ips = "1414993", Designation = "04.067.02.003", Name = "Гайка плоская" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            Search = "04.067.02.003 Гайка плоская"
        };
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle(x => x.Ips == "1414993");
    }

    [Fact]
    public async Task Library_editor_displays_millimeters_but_stores_meter_consumption()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60 Сталь 20",
            CanonicalKey = "ROUND|D60|STEEL20|MM",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "Сталь 20",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000012345",
            SourceName = "Круг D60 Сталь 20",
            NormalizedSourceName = "Круг D60 Сталь 20",
            Source = "test",
            IsActive = true
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            EditIps = "1371148",
            EditDesignation = "TEST.001",
            EditPartName = "Тестовая деталь",
            SelectedBlank = new LibraryBlankOption(blank.Id, blank.BlankType, blank.CanonicalName, "Круг D60 Сталь 20", blank.Material, blank.BaseUnit, "УТ000012345"),
            ConsumptionQuantityText = "62"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.RoundBar);
        viewModel.SelectedUnit = viewModel.UnitTypes.Single(x => x.Value == MeasurementUnit.Meter && x.DisplayName == "мм");

        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        var map = await db.PartBlankMaps.SingleAsync();
        map.ConsumptionUnit.Should().Be(MeasurementUnit.Meter);
        map.ConsumptionQuantity.Should().Be(0.062m);

        await viewModel.LoadAsync();
        var row = viewModel.Rows.Single(x => x.Ips == "1371148");
        row.Quantity.Should().Be("62");
        row.UnitName.Should().Be("мм");
    }

    [Fact]
    public async Task Library_demand_filter_keeps_only_latest_demand_ips()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "1414993", Designation = "04.067.02.003", Name = "Гайка плоская" },
            new Part { Ips = "1414994", Designation = "04.067.02.004", Name = "Шайба" });
        var batch = new DemandBatch { Name = "Потребность", ImportedAt = DateTime.UtcNow };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Ips = "1414993", SourcePartName = "04.067.02.003 Гайка плоская", Quantity = 1, Unit = "шт" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            DemandOnly = true
        };
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle(x => x.Ips == "1414993");
        viewModel.SummaryText.Should().Contain("в потребности");
    }

    [Fact]
    public async Task Library_load_uses_sql_sorting_before_in_memory_row_formatting()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D60 Steel 40X",
            CanonicalKey = "ROUNDBAR|40X|D60",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "Steel 40X"
        };
        var part = new Part { Ips = "100001", Name = "Shaft", UpdatedAt = DateTime.UtcNow };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, Source = "test" });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT001", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().Ips.Should().Be("100001");
        viewModel.Rows.Single().OneCCode.Should().Be("UT001");
        viewModel.Rows.Single().BlankType.Should().Be("Круг");
    }

    [Fact]
    public async Task Library_search_matches_lowercase_cyrillic_input_for_designation_and_name()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "1537692", Designation = "01-P014-01.120-T50.50.SS-01", Name = "Удленитель направляющей", UpdatedAt = DateTime.UtcNow },
            new Part { Ips = "1537737", Designation = "01-P014-01.002-T50.50.SS", Name = "Направляющая Z верхняя", UpdatedAt = DateTime.UtcNow },
            new Part { Ips = "1400000", Designation = "99-A001", Name = "Корпус", UpdatedAt = DateTime.UtcNow });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            Search = "01-р014"
        };
        await viewModel.LoadAsync();

        viewModel.Rows.Select(x => x.Ips).Should().BeEquivalentTo(["1537692", "1537737"]);

        viewModel.Search = "напра";
        await viewModel.LoadAsync();

        viewModel.Rows.Select(x => x.Ips).Should().BeEquivalentTo(["1537692", "1537737"]);
    }


    [Fact]
    public async Task Dashboard_counts_active_visible_rows_and_nsi_duplicate_filter_shows_duplicate_aliases()
    {
        await using var db = CreateDb();
        var duplicateLeft = new CanonicalBlank { CanonicalName = "Круг D50", CanonicalKey = "DUP-L", BlankType = BlankType.RoundBar, DiameterMm = 50, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Meter, I012Status = I012Status.Allowed };
        var duplicateRight = new CanonicalBlank { CanonicalName = "Круг D50", CanonicalKey = "DUP-R", BlankType = BlankType.RoundBar, DiameterMm = 50, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Meter, I012Status = I012Status.Allowed };
        var unique = new CanonicalBlank { CanonicalName = "Лист 10", CanonicalKey = "UNIQUE", BlankType = BlankType.Plate, ThicknessMm = 10, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece, I012Status = I012Status.Allowed };
        var archivedBlank = new CanonicalBlank { CanonicalName = "Архив", CanonicalKey = "ARCHIVE", IsActive = false };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = duplicateLeft, OneCCode = "UT-DUP-1", SourceName = "Круг D50", NormalizedSourceName = "ROUND D50", Source = "test" },
            new BlankAlias { CanonicalBlank = duplicateRight, OneCCode = "UT-DUP-2", SourceName = "Круг D50", NormalizedSourceName = "ROUND D50", Source = "test" },
            new BlankAlias { CanonicalBlank = unique, OneCCode = "UT-UNIQUE", SourceName = "Лист 10", NormalizedSourceName = "PLATE 10", Source = "test" },
            new BlankAlias { CanonicalBlank = archivedBlank, OneCCode = "UT-ARCHIVE", SourceName = "Архив", NormalizedSourceName = "ARCHIVE", Source = "test" });
        var activePart = new Part { Ips = "100001", Name = "Активная деталь", HasMsk = true };
        var noBlankPart = new Part { Ips = "100002", Name = "Без заготовки" };
        var archivedPart = new Part { Ips = "100003", Name = "Архивная деталь", Source = "Потребность [ARCHIVED_LIBRARY]" };
        db.PartBlankMaps.Add(new PartBlankMap { Part = activePart, CanonicalBlank = unique, Source = "test" });
        db.Parts.AddRange(noBlankPart, archivedPart);
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandBatches.Add(batch);
        await db.SaveChangesAsync();
        db.CalculationRuns.Add(new CalculationRun { DemandBatchId = batch.Id, StartedAt = new DateTime(2026, 7, 26), Items = { new CalculationItem { CanonicalName = "old", PurchaseQuantity = 2 } } });
        db.CalculationRuns.Add(new CalculationRun { DemandBatchId = batch.Id, StartedAt = new DateTime(2026, 7, 27), Items = { new CalculationItem { CanonicalName = "new", PurchaseQuantity = 1 } } });
        await db.SaveChangesAsync();

        var dashboard = new DashboardViewModel(db);
        await dashboard.LoadAsync();
        var nsi = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService()) { DuplicatesOnly = true };
        await nsi.LoadAsync();

        dashboard.Parts.Should().Be(2);
        dashboard.CanonicalBlanks.Should().Be(3);
        dashboard.PartsWithoutBlank.Should().Be(1);
        dashboard.PartsWithoutMsk.Should().Be(1);
        dashboard.DuplicateCandidates.Should().Be(2);
        dashboard.PurchasePositions.Should().Be(1);
        nsi.Rows.Select(x => x.OneCCode).Should().BeEquivalentTo("UT-DUP-1", "UT-DUP-2");
        nsi.Rows.Select(x => x.DuplicateKey).Should().AllSatisfy(x => x.Should().Contain("D50"));
    }

    [Fact]
    public async Task Nsi_duplicate_filter_groups_duplicate_rows_by_physical_key()
    {
        await using var db = CreateDb();
        AddNsiAlias(db, "UT-D90-1", "Круг ф 90 сталь 40Х", 90);
        AddNsiAlias(db, "UT-D230-1", "Круг ф 230 сталь 40Х", 230);
        AddNsiAlias(db, "UT-D150-1", "Круг ф 150 сталь 40Х", 150);
        AddNsiAlias(db, "UT-D150-2", "Круг D150 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", 150);
        AddNsiAlias(db, "UT-D230-2", "Круг D230 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", 230);
        AddNsiAlias(db, "UT-D90-2", "Круг D90 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", 90);
        await db.SaveChangesAsync();

        var nsi = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService()) { DuplicatesOnly = true };
        await nsi.LoadAsync();

        nsi.Rows.Should().HaveCount(6);
        nsi.Rows.Select(x => x.DuplicateKey).Distinct().Should().HaveCount(3);
        nsi.Rows.Select(x => x.DuplicateKey)
            .Should()
            .Equal(nsi.Rows.Select(x => x.DuplicateKey).OrderBy(x => x, StringComparer.Ordinal));
    }

    [Fact]
    public async Task Nsi_duplicate_filter_does_not_treat_same_material_different_sizes_as_duplicates()
    {
        await using var db = CreateDb();
        var round50 = new CanonicalBlank { CanonicalName = "Круг D50 Сталь 40Х", CanonicalKey = "ROUND50", BlankType = BlankType.RoundBar, DiameterMm = 50, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Meter };
        var round60 = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "ROUND60", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Meter };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round50, OneCCode = "UT050", SourceName = "Круг D50 Сталь 40Х", NormalizedSourceName = "СТАЛЬ40Х", Source = "test" },
            new BlankAlias { CanonicalBlank = round60, OneCCode = "UT060", SourceName = "Круг D60 Сталь 40Х", NormalizedSourceName = "СТАЛЬ40Х", Source = "test" });
        await db.SaveChangesAsync();

        var dashboard = new DashboardViewModel(db);
        await dashboard.LoadAsync();
        var nsi = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService()) { DuplicatesOnly = true };
        await nsi.LoadAsync();

        dashboard.DuplicateCandidates.Should().Be(0);
        nsi.Rows.Should().BeEmpty();
    }

    private static void AddNsiAlias(BlankDemandPlannerDbContext db, string code, string name, decimal diameter)
    {
        var blank = new CanonicalBlank
        {
            CanonicalName = name,
            CanonicalKey = code,
            BlankType = BlankType.RoundBar,
            DiameterMm = diameter,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = code,
            SourceName = name,
            NormalizedSourceName = name,
            Source = "test"
        });
    }

    [Fact]
    public async Task Nsi_edit_command_ignores_invalid_context_menu_parameter()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank { CanonicalName = "Круг D50", CanonicalKey = "EDIT-INVALID", BlankType = BlankType.RoundBar, DiameterMm = 50 };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT050", SourceName = "Круг D50", NormalizedSourceName = "Круг D50", Source = "test" });
        await db.SaveChangesAsync();

        var nsi = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await nsi.LoadAsync();
        nsi.SelectedRow = nsi.Rows.Single();

        var act = () => nsi.EditNsiRowCommand.Execute(new object());

        act.Should().NotThrow();
        nsi.EditOneCCode.Should().Be("UT050");
    }

    [Fact]
    public async Task Library_editor_filters_blanks_and_saves_manual_part_mapping()
    {
        await using var db = CreateDb();
        var round60 = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUNDBAR|40РҐ|D60",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "РЎС‚Р°Р»СЊ 40РҐ"
        };
        var round80 = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D80 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUNDBAR|40РҐ|D80",
            BlankType = BlankType.RoundBar,
            DiameterMm = 80,
            Material = "РЎС‚Р°Р»СЊ 40РҐ"
        };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round60, OneCCode = "UT060", SourceName = round60.CanonicalName, NormalizedSourceName = round60.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = round80, OneCCode = "UT080", SourceName = round80.CanonicalName, NormalizedSourceName = round80.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            BlankSearch = "D60"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.RoundBar);
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);

        viewModel.BlankSuggestions.Should().ContainSingle(x => x.OneCCode == "UT060");

        viewModel.EditIps = "200001";
        viewModel.EditDesignation = "A-1";
        viewModel.EditPartName = "Manual part";
        viewModel.EditRequiresNitriding = true;
        viewModel.EditRequiresHeatTreatment = true;
        viewModel.EditBlankSupplyRequiresHeatTreatment = true;
        viewModel.EditBlankSupplyRequiresLaserCutting = true;
        viewModel.SelectedBlank = viewModel.BlankSuggestions.Single();
        viewModel.ConsumptionQuantityText = "2,5";
        viewModel.SelectedUnit = viewModel.UnitTypes.Single(x => x.Value == MeasurementUnit.Meter && x.DisplayName == "пог. м");
        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        var part = await db.Parts.Include(x => x.BlankMaps).SingleAsync(x => x.Ips == "200001");
        part.BlankMaps.Should().ContainSingle();
        part.BlankMaps.Single().CanonicalBlankId.Should().Be(round60.Id);
        part.BlankMaps.Single().ConsumptionQuantity.Should().Be(2.5m);
        part.BlankMaps.Single().ConsumptionUnit.Should().Be(MeasurementUnit.Meter);
        part.RequiresNitriding.Should().BeTrue();
        part.RequiresHeatTreatment.Should().BeTrue();
        part.BlankSupplyRequiresHeatTreatment.Should().BeTrue();
        part.BlankSupplyRequiresLaserCutting.Should().BeTrue();

        await viewModel.LoadAsync();
        var row = viewModel.Rows.Single(x => x.Ips == "200001");
        row.Note.Should().Be("Азотирование; ТО");
        row.BlankSupplyRequirement.Should().Be("ТО; Лазерная резка");
    }

    [Fact]
    public async Task Library_editor_blank_search_treats_short_diameter_as_prefix()
    {
        await using var db = CreateDb();
        var round115 = new CanonicalBlank
        {
            CanonicalName = "Круг D115 Сталь 40Х",
            CanonicalKey = "ROUNDBAR|40X|D115-PREFIX",
            BlankType = BlankType.RoundBar,
            DiameterMm = 115,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        var round140 = new CanonicalBlank
        {
            CanonicalName = "Круг D140 Сталь 40Х",
            CanonicalKey = "ROUNDBAR|40X|D140-PREFIX",
            BlankType = BlankType.RoundBar,
            DiameterMm = 140,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        var round90 = new CanonicalBlank
        {
            CanonicalName = "Круг D90 Сталь 40Х",
            CanonicalKey = "ROUNDBAR|40X|D90-PREFIX",
            BlankType = BlankType.RoundBar,
            DiameterMm = 90,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };

        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round115, OneCCode = "UT115", SourceName = "Круг D115 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", NormalizedSourceName = "Круг D115 Сталь 40Х", Source = "test" },
            new BlankAlias { CanonicalBlank = round140, OneCCode = "UT140", SourceName = "Круг D140 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", NormalizedSourceName = "Круг D140 Сталь 40Х", Source = "test" },
            new BlankAlias { CanonicalBlank = round90, OneCCode = "UT090", SourceName = "Круг D90 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", NormalizedSourceName = "Круг D90 Сталь 40Х", Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            BlankSearch = "Круг D1"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.RoundBar);
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);

        viewModel.BlankSuggestions.Select(x => x.OneCCode).Should().Contain(["UT115", "UT140"]);
        viewModel.BlankSuggestions.Should().NotContain(x => x.OneCCode == "UT090");
    }

    [Fact]
    public async Task Library_load_displays_meter_for_meter_based_blank_even_when_old_map_was_piece()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D37",
            CanonicalKey = "ROUNDBAR|D37",
            BlankType = BlankType.RoundBar,
            DiameterMm = 37,
            BaseUnit = MeasurementUnit.Piece
        };
        var part = new Part { Ips = "1223368", Designation = "РћР¦-80.001", Name = "РћРїСЂР°РІРєР° С†РµРЅС‚СЂРѕРІР°СЏ", UpdatedAt = DateTime.UtcNow };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Piece,
            Source = "old import"
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        viewModel.Rows.Single().Quantity.Should().Be("1000");
        viewModel.Rows.Single().UnitName.Should().Be("мм");
        viewModel.Rows.Single().ConsumptionUnit.Should().Be(MeasurementUnit.Meter);
    }

    [Fact]
    public async Task Library_editor_selecting_blank_sets_one_c_code_and_meter_unit()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUNDBAR|40X|D60-AUTO",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "40РҐ",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT060", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            BlankSearch = "D60"
        };
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);
        viewModel.SelectedBlank = viewModel.BlankSuggestions.Single();

        viewModel.SelectedOneCCode.Should().Be("UT060");
        viewModel.SelectedUnit.Value.Should().Be(MeasurementUnit.Meter);
    }

    [Fact]
    public async Task Library_editor_finds_purchased_blank_by_name_and_code_when_import_type_is_unknown()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Заготовка оправки ВТ40-D70-L175",
            CanonicalKey = "CUSTOM|1802132",
            BlankType = BlankType.CustomBlank,
            BaseUnit = MeasurementUnit.Piece
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "1802132",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            BlankSearch = "Заготовка оправки ВТ40-D70-L175 | 1802132"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.Purchased);
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);

        viewModel.BlankSuggestions.Should().ContainSingle(x => x.OneCCode == "1802132");
    }

    [Fact]
    public async Task Library_editor_saves_blank_by_typed_name_and_code_when_selected_item_was_not_set()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Заготовка оправки ВТ40-D70-L175",
            CanonicalKey = "CUSTOM|1802132|SAVE",
            BlankType = BlankType.CustomBlank,
            BaseUnit = MeasurementUnit.Piece
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "1802132",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            EditIps = "900001",
            EditPartName = "Деталь с покупной заготовкой",
            BlankSearch = "Заготовка оправки ВТ40-D70-L175 | 1802132",
            ConsumptionQuantityText = "1"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.Purchased);

        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        viewModel.SelectedOneCCode.Should().Be("1802132");
        var part = await db.Parts.Include(x => x.BlankMaps).SingleAsync(x => x.Ips == "900001");
        part.BlankMaps.Should().ContainSingle(x => x.CanonicalBlankId == blank.Id);
    }

    [Fact]
    public async Task Library_editor_saves_blank_lead_time_days()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D30",
            CanonicalKey = "ROUND|D30|LEAD",
            BlankType = BlankType.RoundBar,
            DiameterMm = 30,
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT030", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            EditIps = "930001",
            EditPartName = "Деталь со сроком",
            BlankSearch = "Круг D30",
            ConsumptionQuantityText = "1",
            BlankLeadTimeDaysText = "45"
        };
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);
        viewModel.SelectedBlank = viewModel.BlankSuggestions.Single();

        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        var map = await db.PartBlankMaps.SingleAsync();
        map.BlankLeadTimeDays.Should().Be(45);
        await viewModel.LoadAsync();
        viewModel.Rows.Single().BlankLeadTimeDaysText.Should().Be("45");
    }

    [Fact]
    public void Library_unit_filter_does_not_offer_kilograms()
    {
        var viewModel = new LibraryViewModel(CreateDb(), new BlankNormalizationService());

        viewModel.UnitFilters.Should().NotContain(x => x.Value == MeasurementUnit.Kilogram);
        viewModel.UnitFilters.Select(x => x.DisplayName).Should().BeEquivalentTo(["Все ед.", "шт", "пог. м"]);
    }

    [Fact]
    public async Task Library_editor_replaces_previous_active_blank_for_same_ips()
    {
        await using var db = CreateDb();
        var oldBlank = new CanonicalBlank { CanonicalName = "Круг D40", CanonicalKey = "ROUND|OLD|D40", BlankType = BlankType.RoundBar, DiameterMm = 40, BaseUnit = MeasurementUnit.Meter };
        var newBlank = new CanonicalBlank { CanonicalName = "Круг D60", CanonicalKey = "ROUND|NEW|D60", BlankType = BlankType.RoundBar, DiameterMm = 60, BaseUnit = MeasurementUnit.Meter };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = oldBlank, OneCCode = "UT040", SourceName = oldBlank.CanonicalName, NormalizedSourceName = oldBlank.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = newBlank, OneCCode = "UT060", SourceName = newBlank.CanonicalName, NormalizedSourceName = newBlank.CanonicalName, Source = "test" });
        var part = new Part { Ips = "970001", Name = "Деталь с заменой материала" };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = oldBlank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Meter, IsActive = true, IsPrimary = true, Source = "old" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            EditIps = "970001",
            EditPartName = "Деталь с заменой материала",
            BlankSearch = "Круг D60",
            ConsumptionQuantityText = "2"
        };
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);
        viewModel.SelectedBlank = viewModel.BlankSuggestions.Single(x => x.OneCCode == "UT060");

        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        var maps = await db.PartBlankMaps.AsNoTracking().Where(x => x.PartId == part.Id).ToListAsync();
        maps.Single(x => x.CanonicalBlankId == oldBlank.Id).IsActive.Should().BeFalse();
        maps.Single(x => x.CanonicalBlankId == newBlank.Id).IsActive.Should().BeTrue();
        maps.Single(x => x.CanonicalBlankId == newBlank.Id).IsPrimary.Should().BeTrue();
    }

    [Fact]
    public async Task Blank_selection_suggests_bigger_round_bar_by_smallest_allowance()
    {
        await using var db = CreateDb();
        var round50 = new CanonicalBlank { CanonicalName = "Круг D50 Сталь 40Х", CanonicalKey = "PICK|ROUND|50", BlankType = BlankType.RoundBar, DiameterMm = 50, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var round60 = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "PICK|ROUND|60", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var round80 = new CanonicalBlank { CanonicalName = "Круг D80 Сталь 40Х", CanonicalKey = "PICK|ROUND|80", BlankType = BlankType.RoundBar, DiameterMm = 80, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round50, OneCCode = "UT050", SourceName = round50.CanonicalName, NormalizedSourceName = round50.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = round60, OneCCode = "UT060", SourceName = round60.CanonicalName, NormalizedSourceName = round60.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = round80, OneCCode = "UT080", SourceName = round80.CanonicalName, NormalizedSourceName = round80.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            DiameterText = "55",
            Material = "40Х"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.Rows.Select(x => x.OneCCode).Should().Equal("UT060", "UT080");
    }

    [Fact]
    public async Task Blank_selection_treats_forging_diameter_as_round_candidate()
    {
        await using var db = CreateDb();
        var forging450 = new CanonicalBlank { CanonicalName = "Поковка круг D450 L20 Сталь 40Х", CanonicalKey = "PICK|FORGING|450", BlankType = BlankType.Forging, DiameterMm = 450, LengthMm = 20, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        var forging660 = new CanonicalBlank { CanonicalName = "Поковка круг D660 L25 Сталь 40Х", CanonicalKey = "PICK|FORGING|660", BlankType = BlankType.Forging, DiameterMm = 660, LengthMm = 25, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        var forging300 = new CanonicalBlank { CanonicalName = "Поковка круг D300 L20 Сталь 40Х", CanonicalKey = "PICK|FORGING|300", BlankType = BlankType.Forging, DiameterMm = 300, LengthMm = 20, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = forging450, OneCCode = "УТ000015173", SourceName = forging450.CanonicalName, NormalizedSourceName = forging450.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = forging660, OneCCode = "УТ000022659", SourceName = forging660.CanonicalName, NormalizedSourceName = forging660.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = forging300, OneCCode = "УТ000030000", SourceName = forging300.CanonicalName, NormalizedSourceName = forging300.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            DiameterText = "400"
        };

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.Rows.Select(x => x.OneCCode).Should().Equal("УТ000015173", "УТ000022659");
    }

    [Fact]
    public async Task Blank_selection_uses_only_active_nsi_aliases()
    {
        await using var db = CreateDb();
        var nsiBlank = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "PICK|NSI|60", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var orphanBlank = new CanonicalBlank { CanonicalName = "Круг D58 Сталь 40Х", CanonicalKey = "PICK|NO-NSI|58", BlankType = BlankType.RoundBar, DiameterMm = 58, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        db.CanonicalBlanks.Add(orphanBlank);
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = nsiBlank, OneCCode = "UT060", SourceName = nsiBlank.CanonicalName, NormalizedSourceName = nsiBlank.CanonicalName, Source = "test", IsActive = true });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            DiameterText = "55",
            Material = "40Х"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().OneCCode.Should().Be("UT060");
    }

    [Fact]
    public async Task Blank_selection_can_filter_candidates_by_positive_stock()
    {
        await using var db = CreateDb();
        var round60 = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "PICK|STOCK|60", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var round80 = new CanonicalBlank { CanonicalName = "Круг D80 Сталь 40Х", CanonicalKey = "PICK|STOCK|80", BlankType = BlankType.RoundBar, DiameterMm = 80, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var snapshot = new StockSnapshot();
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round60, OneCCode = "УТ000060", SourceName = round60.CanonicalName, NormalizedSourceName = round60.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = round80, OneCCode = "УТ000080", SourceName = round80.CanonicalName, NormalizedSourceName = round80.CanonicalName, Source = "test" });
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, OneCCode = "УТ000080", SourceName = round80.CanonicalName, Quantity = 2, Unit = MeasurementUnit.Meter, Warehouse = "1 секция (ПЗМЦ) только для производственных целей" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            DiameterText = "55",
            Material = "40Х",
            PositiveStockOnly = true
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().OneCCode.Should().Be("УТ000080");
        viewModel.Rows.Single().StockValue.Should().Be(2);
    }

    [Fact]
    public async Task Blank_selection_adds_selected_part_and_blank_to_library()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1799544", Designation = "01-P002-03.002-T63.150.SS", Name = "Направляющая Z верхняя", Source = "demand" };
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60 Сталь 40Х",
            CanonicalKey = "PICK|ROUND|60|LIB",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.Parts.Add(part);
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT060", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            PartSearch = "1799544",
            DiameterText = "55",
            Material = "40Х",
            ConsumptionQuantityText = "0,25"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);
        await viewModel.LoadPartSuggestionsCommand.ExecuteAsync(null);
        viewModel.SelectedPart = viewModel.PartSuggestions.Single();

        await viewModel.FindCommand.ExecuteAsync(null);
        viewModel.SelectedBlankRow = viewModel.Rows.Single();
        await viewModel.AddToLibraryCommand.ExecuteAsync(null);

        var map = await db.PartBlankMaps.AsNoTracking().SingleAsync();
        map.PartId.Should().Be(part.Id);
        map.CanonicalBlankId.Should().Be(blank.Id);
        map.ConsumptionQuantity.Should().Be(0.25m);
        map.ConsumptionUnit.Should().Be(MeasurementUnit.Meter);
        map.IsActive.Should().BeTrue();
        map.IsPrimary.Should().BeTrue();
        viewModel.PartSuggestions.Should().BeEmpty();
    }

    [Fact]
    public async Task Blank_selection_keeps_selected_part_after_pick_and_opens_suggestions_when_typing()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "120001", Designation = "DET-120001", Name = "Первая деталь" },
            new Part { Ips = "120002", Designation = "DET-120002", Name = "Вторая деталь" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db);
        viewModel.PartSearch = "12";
        await Task.Delay(200);

        viewModel.PartSuggestions.Should().HaveCount(2);
        viewModel.IsPartSuggestionsOpen.Should().BeTrue();

        var selected = viewModel.PartSuggestions.First();
        viewModel.SelectedPart = selected;

        viewModel.SelectedPart.Should().Be(selected);
        viewModel.PartSearch.Should().Be(selected.DisplayName);
        viewModel.IsPartSuggestionsOpen.Should().BeFalse();
    }

    [Fact]
    public async Task Blank_selection_suggests_only_visible_library_parts_without_blank()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "1223371", Designation = "ОЦ-80.002", Name = "Пробка центровая", Source = "[ARCHIVED_LIBRARY] [MANUAL_LIBRARY_DELETE] Удалено из библиотеки" },
            new Part { Ips = "1223372", Designation = "ОЦ-80.003", Name = "Пробка активная", Source = "Потребность" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db) { PartSearch = "122337" };
        await viewModel.LoadPartSuggestionsCommand.ExecuteAsync(null);

        viewModel.PartSuggestions.Should().ContainSingle(x => x.Ips == "1223372");
        viewModel.PartSuggestions.Should().NotContain(x => x.Ips == "1223371");
    }

    [Fact]
    public async Task Blank_selection_find_keeps_user_selected_blank_row()
    {
        await using var db = CreateDb();
        var round60 = new CanonicalBlank { CanonicalName = "Круг D60 Сталь 40Х", CanonicalKey = "PICK|KEEP|60", BlankType = BlankType.RoundBar, DiameterMm = 60, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        var round80 = new CanonicalBlank { CanonicalName = "Круг D80 Сталь 40Х", CanonicalKey = "PICK|KEEP|80", BlankType = BlankType.RoundBar, DiameterMm = 80, Material = "40Х", BaseUnit = MeasurementUnit.Meter };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round60, OneCCode = "UT060", SourceName = round60.CanonicalName, NormalizedSourceName = round60.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = round80, OneCCode = "UT080", SourceName = round80.CanonicalName, NormalizedSourceName = round80.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db)
        {
            DiameterText = "55",
            Material = "40Х"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);
        await viewModel.FindCommand.ExecuteAsync(null);
        viewModel.SelectedBlankRow = viewModel.Rows.Single(x => x.OneCCode == "UT080");

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.SelectedBlankRow.Should().NotBeNull();
        viewModel.SelectedBlankRow!.OneCCode.Should().Be("UT080");
    }

    [Fact]
    public async Task Blank_selection_shows_selected_part_detail_text()
    {
        await using var db = CreateDb();
        db.Parts.Add(new Part { Ips = "120001", Designation = "DET-120001", Name = "Первая деталь" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db);
        await viewModel.LoadPartSuggestionsCommand.ExecuteAsync(null);

        viewModel.SelectedPart = viewModel.PartSuggestions.Single();

        viewModel.PartDetailText.Should().Contain("IPS: 120001");
        viewModel.PartDetailText.Should().Contain("Обозначение: DET-120001");
        viewModel.PartDetailText.Should().Contain("Наименование: Первая деталь");
        viewModel.PartDetailText.Should().Contain("Заготовка: Не подобрана");
    }

    [Fact]
    public async Task Blank_selection_finds_blanks_by_material_word()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D320 Сталь 40Х",
            CanonicalKey = "PICK|ROUND|320|STEEL",
            BlankType = BlankType.RoundBar,
            DiameterMm = 320,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000004539", SourceName = "Круг ф 320 сталь 40Х", NormalizedSourceName = "Круг ф 320 сталь 40Х", Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db) { Material = "Сталь" };

        await viewModel.FindCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle(x => x.OneCCode == "УТ000004539");
    }

    [Fact]
    public async Task Blank_selection_updates_results_while_typing_attributes()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60 Сталь 40Х",
            CanonicalKey = "PICK|ROUND|60|LIVE",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT060", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new BlankSelectionViewModel(db);
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);
        viewModel.Material = "40";
        viewModel.DiameterText = "55";

        await Task.Delay(800);

        viewModel.Rows.Should().ContainSingle(x => x.OneCCode == "UT060");
        viewModel.StatusText.Should().Contain("Подобрано заготовок");
    }

    [Fact]
    public async Task Library_shows_parts_without_blank_and_blank_selection_uses_same_parts()
    {
        await using var db = CreateDb();
        db.Parts.Add(new Part { Ips = "1802132", Designation = "BT40-D70-L175", Name = "Заготовка оправки", Source = "demand" });
        await db.SaveChangesAsync();

        var library = new LibraryViewModel(db, new BlankNormalizationService());
        await library.LoadAsync();

        library.Rows.Should().ContainSingle(x => x.Ips == "1802132" && x.PartBlankMapId == null);
        library.SummaryText.Should().Be("Деталей: 1; без заготовки: 1");

        var selection = new BlankSelectionViewModel(db) { PartSearch = "1802132" };
        await selection.LoadPartSuggestionsCommand.ExecuteAsync(null);

        selection.PartSuggestions.Should().ContainSingle(x => x.Ips == "1802132");
    }

    [Fact]
    public async Task Library_load_adds_new_numeric_demand_part_from_decimal_designation()
    {
        await using var db = CreateDb();
        db.DemandBatches.Add(new DemandBatch
        {
            SourceFile = "Потребность.xlsx",
            ImportedAt = DateTime.UtcNow,
            Items =
            [
                new DemandItem
                {
                    Ips = "1816772",
                    SourcePartName = "270.011.02.001 Вал задней бабки",
                    Quantity = 1,
                    Unit = "шт"
                }
            ]
        });
        await db.SaveChangesAsync();

        var library = new LibraryViewModel(db, new BlankNormalizationService());
        await library.LoadAsync();

        library.Rows.Should().ContainSingle(x =>
            x.Ips == "1816772" &&
            x.Designation == "270.011.02.001" &&
            x.PartName == "Вал задней бабки" &&
            x.PartBlankMapId == null);
    }

    [Fact]
    public async Task Library_load_restores_archived_part_when_it_returns_in_actual_demand()
    {
        await using var db = CreateDb();
        var part = new Part
        {
            Ips = "1816772",
            Designation = "270.011.02.001",
            Name = "Вал задней бабки",
            Source = "Потребность; [ARCHIVED_LIBRARY] Удалено из библиотеки 2026-07-24 11:35"
        };
        db.Parts.Add(part);
        db.DemandBatches.Add(new DemandBatch
        {
            SourceFile = "Потребность.xlsx",
            ImportedAt = DateTime.UtcNow,
            Items =
            [
                new DemandItem
                {
                    Part = part,
                    Ips = "1816772",
                    SourcePartName = "270.011.02.001 Вал задней бабки",
                    Quantity = 1,
                    Unit = "шт",
                    DemandDate = new DateTime(2026, 7, 30)
                }
            ]
        });
        await db.SaveChangesAsync();

        var library = new LibraryViewModel(db, new BlankNormalizationService());
        await library.LoadAsync();

        library.Rows.Should().ContainSingle(x =>
            x.Ips == "1816772" &&
            x.Designation == "270.011.02.001" &&
            x.PartName == "Вал задней бабки" &&
            x.PartBlankMapId == null);
        (await db.Parts.AsNoTracking().SingleAsync(x => x.Ips == "1816772")).Source.Should().NotContain("[ARCHIVED_LIBRARY]");
    }

    [Fact]
    public async Task Library_editor_loads_existing_row_for_right_click_edit()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "ROUNDBAR|40РҐ|D60-EDIT",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "РЎС‚Р°Р»СЊ 40РҐ"
        };
        var part = new Part { Ips = "300001", Designation = "B-1", Name = "Editable part", Source = "test" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 3,
            ConsumptionUnit = MeasurementUnit.Piece,
            Source = "test"
        });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UTEDIT", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();
        await viewModel.EditLibraryRowCommand.ExecuteAsync(viewModel.Rows.Single());

        viewModel.EditIps.Should().Be("300001");
        viewModel.EditDesignation.Should().Be("B-1");
        viewModel.EditPartName.Should().Be("Editable part");
        viewModel.SelectedBlank.Should().NotBeNull();
        viewModel.SelectedOneCCode.Should().Be("UTEDIT");
        viewModel.SelectedUnit.Value.Should().Be(MeasurementUnit.Meter);
    }

    [Fact]
    public async Task Library_context_commands_ignore_non_row_wpf_parameters()
    {
        await using var db = CreateDb();
        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());

        var canEdit = viewModel.EditLibraryRowCommand.CanExecute(new object());
        var canDelete = viewModel.DeleteLibraryRowCommand.CanExecute(new object());
        await viewModel.EditLibraryRowCommand.ExecuteAsync(new object());
        await viewModel.DeleteLibraryRowCommand.ExecuteAsync(new object());

        canEdit.Should().BeTrue();
        canDelete.Should().BeTrue();
        viewModel.EditorStatus.Should().Be("Выберите строку библиотеки для удаления.");
    }

    [Fact]
    public async Task Library_selecting_row_does_not_load_editor_until_edit_command()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "300002", Designation = "B-2", Name = "Do not auto edit" };
        db.Parts.Add(part);
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService())
        {
            EditIps = "MANUAL",
            EditDesignation = "KEEP",
            EditPartName = "Keep editor"
        };
        await viewModel.LoadAsync();
        var row = viewModel.Rows.Single();

        viewModel.SelectedRow = row;

        viewModel.EditIps.Should().Be("MANUAL");
        viewModel.EditDesignation.Should().Be("KEEP");
        viewModel.EditPartName.Should().Be("Keep editor");

        await viewModel.EditLibraryRowCommand.ExecuteAsync(row);

        viewModel.EditIps.Should().Be("300002");
        viewModel.EditDesignation.Should().Be("B-2");
        viewModel.EditPartName.Should().Be("Do not auto edit");
    }

    [Fact]
    public async Task Library_delete_rows_archives_selected_part_and_removes_it_from_visible_list()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D40",
            CanonicalKey = "ROUNDBAR|D40|DELETE",
            BlankType = BlankType.RoundBar
        };
        var part = new Part { Ips = "400001", Name = "Deleted from library" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        await viewModel.DeleteLibraryRowsCommand.ExecuteAsync(new[] { viewModel.Rows.Single() });

        viewModel.Rows.Should().BeEmpty();
        viewModel.SummaryText.Should().Be("Деталей: 0; без заготовки: 0");
        (await db.PartBlankMaps.IgnoreQueryFilters().SingleAsync()).IsActive.Should().BeFalse();
        (await db.Parts.CountAsync()).Should().Be(1);
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().Contain("[ARCHIVED_LIBRARY]");
    }

    [Fact]
    public async Task Library_delete_part_without_active_blank_archives_part_even_when_inactive_maps_exist()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D42",
            CanonicalKey = "ROUNDBAR|D42|DELETE-INACTIVE",
            BlankType = BlankType.RoundBar
        };
        var part = new Part { Ips = "400002", Name = "Deleted no active blank", Source = "Потребность" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test",
            IsActive = false,
            IsPrimary = false
        });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, Quantity = 1, Unit = "шт" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();
        var row = viewModel.Rows.Single();
        row.CanonicalBlankId.Should().BeNull();

        await viewModel.DeleteLibraryRowsCommand.ExecuteAsync(row);

        viewModel.Rows.Should().BeEmpty();
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().Contain("[ARCHIVED_LIBRARY]");
        (await db.DemandItems.AsNoTracking().SingleAsync()).PartId.Should().BeNull();
    }

    [Fact]
    public async Task Library_load_does_not_restore_manually_deleted_part_from_existing_demand()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D44",
            CanonicalKey = "ROUNDBAR|D44|DELETE-DEMAND",
            BlankType = BlankType.RoundBar
        };
        var part = new Part { Ips = "400003", Designation = "D-400003", Name = "Demand part", Source = "Потребность" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Part = part,
            Ips = part.Ips,
            SourcePartName = "D-400003 Demand part",
            Quantity = 1,
            Unit = "шт"
        });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();
        await viewModel.DeleteLibraryRowsCommand.ExecuteAsync(viewModel.Rows.Single());
        await viewModel.LoadAsync();

        viewModel.Rows.Should().BeEmpty();
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().Contain("[ARCHIVED_LIBRARY]");
        (await db.DemandItems.AsNoTracking().SingleAsync()).PartId.Should().BeNull();
    }

    [Fact]
    public async Task Export_library_writes_expected_file_name()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D50",
            CanonicalKey = "ROUNDBAR|D50|EXPORT",
            BlankType = BlankType.RoundBar
        };
        var part = new Part
        {
            Ips = "410001",
            Name = "Exported part",
            BlankSupplyRequiresHeatTreatment = true,
            BlankSupplyRequiresLaserCutting = true
        };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var folder = Path.Combine(Path.GetTempPath(), $"bdp_export_{Guid.NewGuid():N}");
        var path = await new ReportExportService(db).ExportLibraryAsync(folder, CancellationToken.None);
        Path.GetFileName(path).Should().Be("Библиотека.xlsx");
        File.Exists(Path.Combine(folder, "Библиотека.xlsx")).Should().BeTrue();
        using var package = new ExcelPackage(new FileInfo(path));
        var sheet = package.Workbook.Worksheets["Библиотека"];
        sheet.Cells[1, 7].Text.Should().Be("Условие поставки заготовки");
        sheet.Cells[2, 7].Text.Should().Be("ТО; Лазерная резка");
    }

    [Fact]
    public async Task Export_library_uses_timestamped_file_when_expected_file_is_locked()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Round D55",
            CanonicalKey = "ROUNDBAR|D55|EXPORT-LOCKED",
            BlankType = BlankType.RoundBar
        };
        var part = new Part { Ips = "410002", Name = "Exported locked part" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var folder = Path.Combine(Path.GetTempPath(), $"bdp_export_locked_{Guid.NewGuid():N}");
        Directory.CreateDirectory(folder);
        var lockedPath = Path.Combine(folder, "Библиотека.xlsx");
        await File.WriteAllTextAsync(lockedPath, "locked");
        await using var locked = new FileStream(lockedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var path = await new ReportExportService(db).ExportLibraryAsync(folder, CancellationToken.None);
        Path.GetFileName(path).Should().StartWith("Библиотека_");
        File.Exists(path).Should().BeTrue();
    }

    [Fact]
    public async Task Import_exported_library_file_updates_edited_rows()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D70",
            CanonicalKey = "ROUNDBAR|D70|EXPORT-IMPORT",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var part = new Part { Ips = "420001", Designation = "OLD-01", Name = "Старое наименование" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            BlankLeadTimeDays = 30,
            Source = "test"
        });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "UT070", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "1C" });
        await db.SaveChangesAsync();

        var folder = Path.Combine(Path.GetTempPath(), $"bdp_library_roundtrip_{Guid.NewGuid():N}");
        var path = await new ReportExportService(db).ExportLibraryAsync(folder, CancellationToken.None);
        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var sheet = package.Workbook.Worksheets["Библиотека"];
            sheet.Cells[2, 2].Value = "NEW-01";
            sheet.Cells[2, 3].Value = "Новое наименование";
            sheet.Cells[2, 9].Value = 2.5m;
            sheet.Cells[2, 11].Value = 45;
            await package.SaveAsync();
        }

        var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
        var report = await importer.ImportManufacturingBlankLibraryAsync(path, null, CancellationToken.None);
        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        report.UpdatedRows.Should().BeGreaterThan(0);
        var row = viewModel.Rows.Single(x => x.Ips == "420001");
        row.Designation.Should().Be("NEW-01");
        row.PartName.Should().Be("Новое наименование");
        row.Quantity.Should().Be("2500");
        row.BlankLeadTimeDays.Should().Be(45);
    }

    [Fact]
    public async Task Library_import_reads_compact_blank_name_file_and_replaces_archived_mapping()
    {
        await using var db = CreateDb();
        var oldBlank = new CanonicalBlank
        {
            CanonicalName = "Шестерня D58,5",
            CanonicalKey = "OLD-GEAR-58",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var part = new Part
        {
            Ips = "1410079",
            Designation = "01.101.00.002",
            Name = "Направляющая задней бабки",
            Source = "Исходные данные для изготовления_заказа заготовок.xlsx; Удалено из библиотеки 2026-07-24 11:35"
        };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = oldBlank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "old",
            IsActive = false,
            IsPrimary = false
        });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = oldBlank, OneCCode = "1785547", SourceName = "Шестерня D58,5мм", NormalizedSourceName = "Шестерня D58,5мм", Source = "old" });
        await db.SaveChangesAsync();

        var path = Path.Combine(Path.GetTempPath(), $"bdp_compact_library_{Guid.NewGuid():N}.xlsx");
        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var sheet = package.Workbook.Worksheets.Add("Лист Microsoft Excel");
            sheet.Cells[1, 1].Value = "1410079";
            sheet.Cells[1, 2].Value = "01.101.00.002";
            sheet.Cells[1, 3].Value = "Направляющая задней бабки";
            sheet.Cells[1, 4].Value = "Заготовка 50х105х1630мм Сталь 30ХГСА";
            sheet.Cells[1, 5].Value = "УТ000013591";
            sheet.Cells[1, 6].Value = "Сталь 30ХГС";
            sheet.Cells[1, 7].Value = "ГОСТ 4543-105";
            sheet.Cells[1, 8].Value = "ГОСТ 19903-2049";
            sheet.Cells[1, 9].Value = "Шт";
            sheet.Cells[1, 10].Value = 2;
            await package.SaveAsync();
        }

        var importer = new ExcelImportService(db, new BlankNormalizationService(), new I012ValidationService(db), NullLogger<ExcelImportService>.Instance);
        var report = await importer.ImportManufacturingBlankLibraryAsync(path, null, CancellationToken.None);
        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        report.ReadRows.Should().Be(1);
        var row = viewModel.Rows.Single(x => x.Ips == "1410079");
        row.OneCCode.Should().Be("УТ000013591");
        row.BlankName.Should().Contain("Заготовка 50х105х1630мм");
        row.Material.Should().Contain("Сталь 30ХГС");
        row.UnitName.Should().Be("шт");
        row.Quantity.Should().Be("2");
        row.Source.Should().Be(Path.GetFileName(path));
        (await db.PartBlankMaps.CountAsync(x => x.PartId == row.PartId && x.IsActive)).Should().Be(1);
    }

    [Fact]
    public async Task Export_calculation_request_omits_material_rows_fully_covered_by_stock()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D220 Сталь 40Х",
            CanonicalKey = "ROUNDBAR|D220|STOCK-COVER",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000014130", SourceName = "Круг D220 Сталь 40Х", NormalizedSourceName = "Круг D220 Сталь 40Х", Source = "test" };
        var part = new Part { Ips = "1452141", Name = "Шкив ведомый" };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 0.133m, ConsumptionUnit = MeasurementUnit.Meter, Source = "test" });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, SerialNumber = "�25604", Quantity = 1, DemandDate = new DateTime(2026, 7, 21) });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 0.133m, Unit = MeasurementUnit.Meter });
        await db.SaveChangesAsync();

        var run = new CalculationRun { DemandBatchId = batch.Id, StockSnapshotId = snapshot.Id };
        run.Items.Add(new CalculationItem { CanonicalBlank = blank, CanonicalName = blank.CanonicalName, PrimaryOneCCode = alias.OneCCode, OneCCodes = alias.OneCCode, Unit = MeasurementUnit.Meter, TotalRequired = 0.133m, TotalStock = 0.133m, PurchaseQuantity = 0, Status = CalculationStatus.Ok });
        db.CalculationRuns.Add(run);
        await db.SaveChangesAsync();

        var folder = Path.Combine(Path.GetTempPath(), $"bdp_calc_export_{Guid.NewGuid():N}");
        var path = await new ReportExportService(db).ExportCalculationRunAsync(run.Id, folder, CancellationToken.None);

        using var package = new ExcelPackage(new FileInfo(path));
        var requestSheet = package.Workbook.Worksheets["Заявка"];
        requestSheet.Should().NotBeNull();
        requestSheet!.Dimension!.Rows.Should().Be(1);
    }

    [Fact]
    public async Task Export_calculation_request_creates_bitrix_file_and_grouped_sheet()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D70 Сталь 40Х",
            CanonicalKey = "ROUND|D70|BITRIX-GROUP",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000022665", SourceName = "Круг D70 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", NormalizedSourceName = "Круг D70 Сталь 40Х", Source = "test" };
        var firstPart = new Part { Ips = "700001", Name = "Деталь 1" };
        var secondPart = new Part { Ips = "700002", Name = "Деталь 2" };
        db.PartBlankMaps.AddRange(
            new PartBlankMap { Part = firstPart, CanonicalBlank = blank, ConsumptionQuantity = 0.85m, ConsumptionUnit = MeasurementUnit.Meter, Source = "test" },
            new PartBlankMap { Part = secondPart, CanonicalBlank = blank, ConsumptionQuantity = 0.46m, ConsumptionUnit = MeasurementUnit.Meter, Source = "test" });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = firstPart, Ips = firstPart.Ips, SourcePartName = firstPart.Name, Quantity = 1, Unit = "шт", DemandDate = new DateTime(2026, 7, 21) },
            new DemandItem { DemandBatch = batch, Part = secondPart, Ips = secondPart.Ips, SourcePartName = secondPart.Name, Quantity = 1, Unit = "шт", DemandDate = new DateTime(2026, 7, 21) });
        db.BlankAliases.Add(alias);
        db.OneCPriceItems.Add(new OneCPriceItem
        {
            LookupKey = StockCodeNormalizer.NormalizeForComparison(alias.OneCCode),
            Code = alias.OneCCode,
            Price = 100m,
            Currency = "RUB",
            PriceType = "Стоимость"
        });
        await db.SaveChangesAsync();

        var run = await new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance)
            .CalculateAsync(new CalculationOptions(batch.Id), CancellationToken.None);
        var folder = Path.Combine(Path.GetTempPath(), $"bdp_bitrix_export_{Guid.NewGuid():N}");
        var path = await new ReportExportService(db).ExportCalculationRunAsync(run.Id, folder, CancellationToken.None);

        Path.GetFileName(path).Should().StartWith($"Заявка на закуп заготовок ЦМО от {DateTime.Now:dd.MM.yyyy} ");
        using var package = new ExcelPackage(new FileInfo(path));
        package.Workbook.Worksheets.Select(x => x.Name).Should().Equal("Заявка", "По группе");
        var requestSheet = package.Workbook.Worksheets["Заявка"]!;
        requestSheet.Cells[1, 7].Text.Should().Be("Код УТ заготовки");
        requestSheet.Cells[1, 7].Text.Should().NotBe("В производстве");
        requestSheet.Cells[1, 12].Text.Should().Be("Цена");
        requestSheet.Cells[2, 12].Text.Should().Be("85 ₽");
        requestSheet.Cells[3, 12].Text.Should().Be("46,5 ₽");
        var groupedSheet = package.Workbook.Worksheets["По группе"];
        groupedSheet.Should().NotBeNull();
        groupedSheet!.Cells[2, 1].Text.Should().Be("Круг D70 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016");
        groupedSheet.Cells[2, 2].Text.Should().Be("УТ000022665");
        groupedSheet.Cells[2, 3].Text.Should().Be("пог. м");
        groupedSheet.Cells[2, 4].GetValue<decimal>().Should().Be(1.315m);
        groupedSheet.Cells[1, 5].Text.Should().Be("Цена");
        groupedSheet.Cells[2, 5].Text.Should().Be("131,5 ₽");
        groupedSheet.Cells[3, 1].Text.Should().Contain("700001");
        groupedSheet.Cells[3, 5].Text.Should().Be("85 ₽");
        groupedSheet.Row(3).OutlineLevel.Should().Be(1);
        groupedSheet.Row(3).Hidden.Should().BeTrue();
        groupedSheet.Cells[4, 1].Text.Should().Contain("700002");
        groupedSheet.Cells[4, 5].Text.Should().Be("46,5 ₽");
        groupedSheet.Row(4).OutlineLevel.Should().Be(1);
        groupedSheet.Row(4).Hidden.Should().BeTrue();
    }

    [Fact]
    public async Task Calculation_view_uses_blank_lead_time_for_blank_demand_date()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D50",
            CanonicalKey = "ROUND|D50|LEAD-CALC",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "UT050", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" };
        var part = new Part { Ips = "940001", Name = "Деталь для расчета срока" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Meter,
            BlankLeadTimeDays = 45,
            Source = "test"
        });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Part = part,
            Ips = part.Ips,
            SourcePartName = part.Name,
            Quantity = 1,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        db.BlankAliases.Add(alias);
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());

        await viewModel.CalculateCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().DemandDate.Should().Be("06.06.2026");
    }

    [Fact]
    public async Task Calculation_view_omits_demand_rows_fully_covered_by_work_in_progress()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Ips = "1668899",
            SourcePartName = "01-P001-02.102-T25.Z Опора ШВП задняя",
            SerialNumber = "Р26666",
            Quantity = 1,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "0001668899",
            SourceName = "НЗП 1668899",
            Quantity = 1,
            Unit = MeasurementUnit.Piece
        });
        await db.SaveChangesAsync();
        db.CalculationRuns.Add(new CalculationRun { DemandBatchId = batch.Id, StockSnapshotId = snapshot.Id });
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());

        await viewModel.LoadLastRunAsync();

        viewModel.Rows.Should().BeEmpty();
    }

    [Fact]
    public async Task Calculation_view_allocates_blank_stock_to_earlier_rows_and_shows_only_uncovered_tail()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Заготовка 85х1020х40мм сталь 30ХГСА",
            CanonicalKey = "SHEET|85X1020X40|30HGSA|VIEW",
            BlankType = BlankType.Sheet,
            BaseUnit = MeasurementUnit.Piece
        };
        var alias = new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000022134",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "test"
        };
        var part = new Part { Ips = "1657364", Designation = "01-P001-02.002-T25.040.RS", Name = "Направляющая задней бабки", RequiresNitriding = true };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 1, ConsumptionUnit = MeasurementUnit.Piece, Source = "test" });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, ProductionSystem = "Р26666", Quantity = 2, DemandDate = new DateTime(2026, 7, 21) },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, ProductionSystem = "Р26689", Quantity = 2, DemandDate = new DateTime(2026, 7, 22) },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, ProductionSystem = "Р26672", Quantity = 2, DemandDate = new DateTime(2026, 7, 23) },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, ProductionSystem = "Р26706", Quantity = 2, DemandDate = new DateTime(2026, 7, 24) });
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

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());

        await viewModel.CalculateCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().MachineNumber.Should().Be("Р26706");
        viewModel.Rows.Single().MaterialQuantity.Should().Be("2");
        viewModel.Rows.Single().OneCCode.Should().Be("УТ000022134");
    }

    [Fact]
    public async Task Calculation_view_marks_missing_blank_as_new_nomenclature()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1668899", Name = "Опора ШВП задняя" };
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Part = part,
            Ips = part.Ips,
            SourcePartName = part.Name,
            Quantity = 1,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());

        await viewModel.CalculateCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().Nomenclature.Should().Be("Новая номенклатура");
        viewModel.Rows.Single().IsMissingBlank.Should().BeTrue();
        (await db.CalculationItems.AsNoTracking().SingleAsync()).CanonicalName.Should().Be("Новая номенклатура");
    }

    [Fact]
    public async Task Calculation_view_shows_library_quantity_without_meter_cut_allowance()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D120 Сталь 20",
            CanonicalKey = "ROUND|D120|VIEW-CUT",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000011939", SourceName = "Круг D120 Сталь 20 ГОСТ 2590-2006 / ГОСТ 1050-2013", NormalizedSourceName = "Круг D120 Сталь 20", Source = "test" };
        var part = new Part { Ips = "1320758", Designation = "КПТ-04.00.002А01", Name = "Корпус" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = blank,
            ConsumptionQuantity = 0.05m,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = $"{part.Designation} {part.Name}", Quantity = 1, Unit = "шт", DemandDate = new DateTime(2026, 7, 21) },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = $"{part.Designation} {part.Name}", Quantity = 1, Unit = "шт", DemandDate = new DateTime(2026, 7, 22) });
        db.BlankAliases.Add(alias);
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());

        await viewModel.CalculateCommand.ExecuteAsync(null);

        viewModel.Rows.Select(x => x.MaterialQuantity).Should().Equal("0,05", "0,05");
        var item = await db.CalculationItems.AsNoTracking().SingleAsync();
        item.TotalRequired.Should().Be(0.105m);
        item.PurchaseQuantity.Should().Be(0.105m);
    }

    [Fact]
    public async Task Calculation_one_time_blank_assignment_does_not_update_library()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D80",
            CanonicalKey = "ROUND|D80|ONE-TIME",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000080", SourceName = "Круг D80", NormalizedSourceName = "Круг D80", Source = "test" });
        var part = new Part { Ips = "1668899", Name = "Опора ШВП задняя" };
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Part = part,
            Ips = part.Ips,
            SourcePartName = part.Name,
            Quantity = 2,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService());
        await viewModel.CalculateCommand.ExecuteAsync(null);
        viewModel.SelectedRow = viewModel.Rows.Single();
        viewModel.BlankSearch = "D80";
        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);
        viewModel.SelectedBlank = viewModel.BlankSuggestions.Single();
        viewModel.ConsumptionQuantityText = "0,5";

        await viewModel.AssignBlankToSelectedCommand.ExecuteAsync(null);

        (await db.PartBlankMaps.CountAsync()).Should().Be(0);
        (await db.CalculationItems.CountAsync(x => x.Comment != null && x.Comment.StartsWith("Разовая заготовка из расчета"))).Should().Be(1);
        viewModel.Rows.Single().Nomenclature.Should().Be("Круг D80");
        viewModel.Rows.Single().MaterialQuantity.Should().Be("1");
    }

    [Fact]
    public async Task Calculation_blank_suggestions_support_attribute_size_search()
    {
        await using var db = CreateDb();
        var forging450 = new CanonicalBlank { CanonicalName = "Поковка круг D450 L20 Сталь 40Х", CanonicalKey = "CALC|FORGING|450", BlankType = BlankType.Forging, DiameterMm = 450, LengthMm = 20, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        var forging660 = new CanonicalBlank { CanonicalName = "Поковка круг D660 L25 Сталь 40Х", CanonicalKey = "CALC|FORGING|660", BlankType = BlankType.Forging, DiameterMm = 660, LengthMm = 25, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        var forging300 = new CanonicalBlank { CanonicalName = "Поковка круг D300 L20 Сталь 40Х", CanonicalKey = "CALC|FORGING|300", BlankType = BlankType.Forging, DiameterMm = 300, LengthMm = 20, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        var orphanForging420 = new CanonicalBlank { CanonicalName = "Поковка круг D420 L20 Сталь 40Х", CanonicalKey = "CALC|FORGING|420|NO-NSI", BlankType = BlankType.Forging, DiameterMm = 420, LengthMm = 20, Material = "Сталь 40Х", BaseUnit = MeasurementUnit.Piece };
        db.CanonicalBlanks.Add(orphanForging420);
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = forging450, OneCCode = "УТ000015173", SourceName = forging450.CanonicalName, NormalizedSourceName = forging450.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = forging660, OneCCode = "УТ000022659", SourceName = forging660.CanonicalName, NormalizedSourceName = forging660.CanonicalName, Source = "test" },
            new BlankAlias { CanonicalBlank = forging300, OneCCode = "УТ000030000", SourceName = forging300.CanonicalName, NormalizedSourceName = forging300.CanonicalName, Source = "test" });
        await db.SaveChangesAsync();

        var calculation = new BlankDemandCalculationService(db, new UnitConversionService(), NullLogger<BlankDemandCalculationService>.Instance);
        var viewModel = new CalculationViewModel(db, calculation, new ReportExportService(db), new StubFileDialogService())
        {
            DiameterText = "400",
            Material = "сталь"
        };
        viewModel.SelectedBlankType = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.Forging);

        await viewModel.LoadBlankSuggestionsCommand.ExecuteAsync(null);

        viewModel.BlankSuggestions.Select(x => x.OneCCode).Should().Equal("УТ000015173", "УТ000022659");
    }

    [Fact]
    public async Task Demand_load_cleans_machine_number_replacement_character()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Ips = "910001",
            SourcePartName = "Деталь",
            SerialNumber = "�25633",
            Quantity = 1,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Single().MachineNumber.Should().Be("Р25633");
    }

    [Fact]
    public async Task Demand_load_distributes_work_in_progress_by_visible_row_order()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "920001",
                SourcePartName = "Первая строка",
                SerialNumber = "Р1",
                Quantity = 3,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 21)
            },
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "920001",
                SourcePartName = "Вторая строка",
                SerialNumber = "Р2",
                Quantity = 4,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 22)
            });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "0000920001",
            SourceName = "НЗП 920001",
            Quantity = 5,
            Unit = MeasurementUnit.Piece
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Select(x => x.InProductionQuantity).Should().Equal("5", "2");
    }

    [Fact]
    public async Task Demand_details_show_dates_quantities_projects_machines_and_remaining_work_in_progress()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "950001",
                Project = "Проект 1",
                SerialNumber = "Р1",
                SourcePartName = "Повторная деталь",
                Quantity = 2,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 21)
            },
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "950001",
                Project = "Проект 2",
                SerialNumber = "Р2",
                SourcePartName = "Повторная деталь",
                Quantity = 3,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 22)
            });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "0000950001",
            SourceName = "НЗП 950001",
            Quantity = 4,
            Unit = MeasurementUnit.Piece
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.ShowDetailsCommand.Execute(viewModel.Rows.First());

        viewModel.DetailTitle.Should().Be("Детализация IPS 950001; всего деталей: 5; в производстве: 4; дефицит: 1");
        viewModel.DetailRows.Select(x => x.DemandDate).Should().Equal("21.07.2026", "22.07.2026");
        viewModel.DetailRows.Select(x => x.Quantity).Should().Equal("2", "3");
        viewModel.DetailRows.Select(x => x.InProductionQuantity).Should().Equal("4", "2");
        viewModel.DetailRows.Select(x => x.WorkInProgressAfter).Should().Equal("2", "0");
        viewModel.DetailRows.Select(x => x.Project).Should().Equal("Проект 1", "Проект 2");
        viewModel.DetailRows.Select(x => x.MachineNumber).Should().Equal("Р1", "Р2");
    }

    [Fact]
    public async Task Demand_details_deficit_uses_initial_work_in_progress_for_repeated_ips()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "1668899",
                SourcePartName = "01-P001-02.102-T25.Z Опора ШВП задняя",
                Quantity = 10,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 21)
            },
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "1668899",
                SourcePartName = "01-P001-02.102-T25.Z Опора ШВП задняя",
                Quantity = 13,
                Unit = "шт",
                DemandDate = new DateTime(2026, 7, 22)
            });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "00001668899",
            SourceName = "НЗП 1668899",
            Quantity = 19,
            Unit = MeasurementUnit.Piece
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.ShowDetailsCommand.Execute(viewModel.Rows.First());

        viewModel.DetailTitle.Should().Be("Детализация IPS 1668899; всего деталей: 23; в производстве: 19; дефицит: 4");
    }

    [Fact]
    public void Demand_detail_panel_can_be_hidden_and_shown()
    {
        using var db = CreateDb();
        var viewModel = new DemandViewModel(db);

        viewModel.IsDetailPanelVisible.Should().BeTrue();
        viewModel.DetailPanelButtonText.Should().Be("Скрыть детализацию");

        viewModel.ToggleDetailPanelCommand.Execute(null);

        viewModel.IsDetailPanelVisible.Should().BeFalse();
        viewModel.DetailPanelButtonText.Should().Be("Отобразить детализацию");

        viewModel.ToggleDetailPanelCommand.Execute(null);

        viewModel.IsDetailPanelVisible.Should().BeTrue();
        viewModel.DetailPanelButtonText.Should().Be("Скрыть детализацию");
    }

    [Fact]
    public async Task Demand_delete_rows_removes_selected_row_and_undo_restores_it()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Ips = "960001",
            SourcePartName = "Удаляемая деталь",
            Quantity = 2,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);

        await viewModel.DeleteDemandRowsCommand.ExecuteAsync(new[] { viewModel.Rows.Single() });

        viewModel.Rows.Should().BeEmpty();
        (await db.DemandItems.CountAsync()).Should().Be(0);

        await UndoCenter.UndoAsync();

        (await db.DemandItems.CountAsync()).Should().Be(1);
    }

    [Fact]
    public async Task Library_load_adds_only_detail_parts_from_demand_ips()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "970777",
                SourcePartName = "01-P002-03.002-T63.150.SS Новая деталь из потребности",
                Quantity = 1,
                Unit = "шт"
            },
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "970778",
                SourcePartName = "Операция фрезерная",
                Quantity = 1,
                Unit = "шт"
            },
            new DemandItem
            {
                DemandBatch = batch,
                Ips = "970779",
                SourcePartName = "01-P002 Сборка узла",
                Quantity = 1,
                Unit = "шт"
            });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle(x => x.Ips == "970777" && x.PartBlankMapId == null);
        viewModel.Rows.Single(x => x.Ips == "970777").IsIncomplete.Should().BeTrue();
        viewModel.Rows.Should().NotContain(x => x.Ips == "970778" || x.Ips == "970779");
        var saved = await db.Parts.AsNoTracking().SingleAsync();
        saved.Designation.Should().Be("01-P002-03.002-T63.150.SS");
        saved.Name.Should().Be("Новая деталь из потребности");
    }

    [Fact]
    public async Task Library_saves_part_name_and_designation_without_selected_blank()
    {
        await using var db = CreateDb();
        db.Parts.Add(new Part { Ips = "970888", Designation = "OLD", Name = "Old name" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadCommand.ExecuteAsync(null);
        await viewModel.EditLibraryRowCommand.ExecuteAsync(viewModel.Rows.Single());
        viewModel.EditDesignation = "NEW-001";
        viewModel.EditPartName = "New name";
        viewModel.BlankSearch = string.Empty;
        viewModel.SelectedBlank = null;

        await viewModel.SaveLibraryEntryCommand.ExecuteAsync(null);

        var saved = await db.Parts.AsNoTracking().SingleAsync();
        saved.Designation.Should().Be("NEW-001");
        saved.Name.Should().Be("New name");
        (await db.PartBlankMaps.CountAsync()).Should().Be(0);
    }

    [Fact]
    public async Task Library_delete_part_without_blank_archives_row_and_undo_restores_it()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1807581", Name = "Доработка бака, сверление, сварка (КУ260508-Т34-02.010)", Source = "Потребность" };
        var batch = new DemandBatch { Name = "Потребность" };
        db.Parts.Add(part);
        db.DemandItems.Add(new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, SourcePartName = part.Name, Quantity = 1, Unit = "шт" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadCommand.ExecuteAsync(null);

        await viewModel.DeleteLibraryRowsCommand.ExecuteAsync(new[] { viewModel.Rows.Single() });

        viewModel.Rows.Should().BeEmpty();
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().Contain("[ARCHIVED_LIBRARY]");
        (await db.DemandItems.AsNoTracking().SingleAsync()).PartId.Should().BeNull();

        await UndoCenter.UndoAsync();
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle(x => x.Ips == "1807581");
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().Be("Потребность");
    }

    [Fact]
    public async Task Library_load_restores_archived_part_when_import_added_active_blank_map()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1354749", Designation = "07.053.00.002", Name = "Диск тормозной", Source = "Потребность [ARCHIVED_LIBRARY] test" };
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D330 Сталь 40Х",
            CanonicalKey = "ROUND|D330|40H",
            BlankType = BlankType.RoundBar,
            DiameterMm = 330,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 0.025m, ConsumptionUnit = MeasurementUnit.Meter, Source = "MskHeaderlessBlankLibrary" });
        await db.SaveChangesAsync();

        var viewModel = new LibraryViewModel(db, new BlankNormalizationService());
        await viewModel.LoadAsync();

        viewModel.Rows.Should().ContainSingle(x => x.Ips == "1354749");
        (await db.Parts.AsNoTracking().SingleAsync()).Source.Should().NotContain("[ARCHIVED_LIBRARY]");
    }

    [Fact]
    public async Task Msk_load_joins_msk_records_with_library_by_ips()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1335520", Designation = "10.301.05.011", Name = "Адаптер G1/4 - М20х1,5" };
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг ф 320 сталь 40Х",
            CanonicalKey = "MSK|ROUND|320|40H",
            BlankType = BlankType.RoundBar,
            DiameterMm = 320,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.Parts.Add(part);
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 0.052m, ConsumptionUnit = MeasurementUnit.Meter, BlankLeadTimeDays = 30, Source = "test" });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000004539", SourceName = "Круг ф 320 сталь 40Х", NormalizedSourceName = "Круг ф 320 сталь 40Х", Source = "test" });
        db.MskRecords.Add(new MskRecord
        {
            Ips = "1335520",
            Designation = "МСК-001",
            Name = "МСК Адаптер",
            FileName = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК\test.xlsx",
            ImportedAt = new DateTime(2026, 7, 23, 8, 0, 0, DateTimeKind.Utc)
        });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        await viewModel.LoadCommand.ExecuteAsync(null);

        var row = viewModel.Rows.Should().ContainSingle().Subject;
        row.Ips.Should().Be("1335520");
        row.HasMsk.Should().Be("Да");
        row.HasLibraryPart.Should().Be("Да");
        row.Name.Should().Be("Адаптер G1/4 - М20х1,5");
        row.BlankType.Should().Be("Круг");
        row.BlankName.Should().Be("Круг ф 320 сталь 40Х");
        row.Material.Should().Be("Сталь 40Х");
        row.OneCCode.Should().Be("УТ000004539");
        row.ConsumptionQuantity.Should().Be("0,052");
        row.UnitName.Should().Be("пог. м");
        row.BlankLeadTimeDays.Should().Be("30");
        viewModel.SelectedRow = row;
        viewModel.DetailText.Should().StartWith("IPS: 1335520");
        viewModel.DetailText.Should().NotContain("\t");
        viewModel.DetailText.Should().NotContain("Файл:");

        viewModel.Search = "1335520";

        viewModel.SelectedRow.Should().Be(row);
        viewModel.DetailText.Should().Contain("Адаптер G1/4 - М20х1,5");
    }

    [Fact]
    public async Task Msk_detail_metrics_sums_decimal_demand_on_client_for_sqlite()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1335520", Designation = "10.301.05.011", Name = "Адаптер G1/4 - М20х1,5" };
        var batch = new DemandBatch { ImportedAt = DateTime.UtcNow };
        db.Parts.Add(part);
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 1.1m },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Quantity = 2.2m });
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = new StockSnapshot(),
            OneCCode = "00001335520",
            SourceName = part.Name,
            Quantity = 1m,
            Unit = MeasurementUnit.Piece,
            Warehouse = "44 секция НЗП (незавершенное производство)"
        });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var row = new MskLibraryRow(part.Ips, part.Designation!, part.Name, "Да", "Да", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "шт", "30");

        viewModel.SelectedRow = row;
        for (var attempt = 0; attempt < 20 && viewModel.DetailMetricsText.Contains("SQLite", StringComparison.OrdinalIgnoreCase) == false && viewModel.DetailMetricsText.Contains("3,3", StringComparison.OrdinalIgnoreCase) == false; attempt++)
        {
            await Task.Delay(50);
        }

        viewModel.DetailMetricsText.Should().NotContain("SQLite cannot apply aggregate operator");
        viewModel.DetailMetricsText.Should().Contain("3,3");
    }

    [Fact]
    public async Task Msk_production_launch_preview_limits_quantity_by_material_stock_and_shows_deficit_targets()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1335520", Designation = "10.301.05.011", Name = "Адаптер G1/4 - М20х1,5" };
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг ф 320 сталь 40Х",
            CanonicalKey = "MSK|ROUND|320|40H",
            BlankType = BlankType.RoundBar,
            DiameterMm = 320,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.Parts.Add(part);
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 0.5m, ConsumptionUnit = MeasurementUnit.Meter, BlankLeadTimeDays = 30, Source = "test" });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000004539", SourceName = "Круг ф 320 сталь 40Х", NormalizedSourceName = "Круг ф 320 сталь 40Х", Source = "test" });
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Project = "ПВ001", ProductionSystem = "Р26666", Quantity = 3, DemandDate = new DateTime(2026, 8, 1) },
            new DemandItem { DemandBatch = batch, Part = part, Ips = part.Ips, Project = "ПВ002", ProductionSystem = "Р26667", Quantity = 2, DemandDate = new DateTime(2026, 8, 2) });
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, OneCCode = "УТ000004539", SourceName = "Круг ф 320 сталь 40Х", Quantity = 2.6m, Unit = MeasurementUnit.Meter },
            new StockItem { StockSnapshot = snapshot, OneCCode = "00001335520", SourceName = "Адаптер", Quantity = 1m, Unit = MeasurementUnit.Piece });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var row = new MskLibraryRow(part.Ips, part.Designation!, part.Name, "Да", "Да", string.Empty, string.Empty, string.Empty, "Круг", "Круг ф 320 сталь 40Х", "Сталь 40Х", "УТ000004539", "0,5", "пог. м", "30");

        var preview = await viewModel.BuildProductionLaunchPreviewAsync(row);

        preview.MaxQuantity.Should().Be(5);
        preview.DefaultQuantity.Should().Be(4);
        preview.MaterialLine.Should().Contain("УТ000004539");
        preview.ShortageRows.Select(x => x.Quantity).Should().Equal(2, 2);
        preview.ShortageRows.Select(x => x.MachineNumber).Should().Equal("Р26666", "Р26667");
    }

    [Fact]
    public async Task Msk_production_launch_preview_uses_cmo_wip_material_stock()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1347282", Designation = "03.062.02.003", Name = "Винт поджимной" };
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D40 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016",
            CanonicalKey = "ROUND|D40|40H",
            BlankType = BlankType.RoundBar,
            DiameterMm = 40,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.Parts.Add(part);
        db.PartBlankMaps.Add(new PartBlankMap { Part = part, CanonicalBlank = blank, ConsumptionQuantity = 0.06m, ConsumptionUnit = MeasurementUnit.Meter, Source = "test" });
        db.BlankAliases.Add(new BlankAlias { CanonicalBlank = blank, OneCCode = "УТ000022660", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "test" });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "УТ000022660",
            SourceName = blank.CanonicalName,
            Quantity = 9.965m,
            Unit = MeasurementUnit.Meter,
            Warehouse = "44 секция НЗП (незавершенное производство)"
        });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var row = new MskLibraryRow(part.Ips, part.Designation!, part.Name, "Да", "Да", string.Empty, string.Empty, string.Empty, "Круг", blank.CanonicalName, "Сталь 40Х", "УТ000022660", "0,06", "пог. м", "30");

        var preview = await viewModel.BuildProductionLaunchPreviewAsync(row);

        preview.MaterialStockQuantity.Should().Be(9.965m);
        preview.MaxQuantity.Should().Be(153);
    }

    [Fact]
    public void Msk_piecewise_production_launch_builds_one_document_per_detail_and_distributes_machines()
    {
        var preview = new ProductionLaunchPreview(
            "1657364",
            "01-P001-02.002-T25.040.RS",
            "Направляющая задней бабки",
            "УТ000022134",
            "Заготовка 85х1020х40мм сталь 30ХГСА",
            1m,
            MeasurementUnit.Piece,
            6m,
            6m,
            6m,
            "Материал: Заготовка 85х1020х40мм сталь 30ХГСА; код УТ000022134; остаток склада: 6 шт; норма: 1 шт на деталь.",
            [
                new ProductionLaunchShortageRow(2m, "ПВ/ПС438", "Р26666", new DateTime(2026, 8, 1)),
                new ProductionLaunchShortageRow(2m, "ПВ/ПС473", "Р26689", new DateTime(2026, 8, 2)),
                new ProductionLaunchShortageRow(2m, "ПВ1004", "Р26672", new DateTime(2026, 8, 3)),
                new ProductionLaunchShortageRow(2m, "ПС515", "Р26706", new DateTime(2026, 8, 4))
            ]);

        var requests = MskViewModel.BuildProductionLaunchRequests(preview, 6m, "Создано ПО Планирование", piecewise: true);

        requests.Should().HaveCount(6);
        requests.Select(x => x.Quantity).Should().AllBeEquivalentTo(1m);
        requests.Select(x => x.Components.Single().Quantity).Should().AllBeEquivalentTo(1m);
        requests.Select(x => x.Comment).Should().Equal(
            "Создано ПО Планирование Р26666; 1 шт",
            "Создано ПО Планирование Р26666; 1 шт",
            "Создано ПО Планирование Р26689; 1 шт",
            "Создано ПО Планирование Р26689; 1 шт",
            "Создано ПО Планирование Р26672; 1 шт",
            "Создано ПО Планирование Р26672; 1 шт");
    }

    [Fact]
    public void Msk_piecewise_production_launch_uses_edited_comment_lines_as_document_comments()
    {
        var preview = new ProductionLaunchPreview(
            "1657364",
            "01-P001-02.002-T25.040.RS",
            "Направляющая задней бабки",
            "УТ000022134",
            "Заготовка 85х1020х40мм сталь 30ХГСА",
            1m,
            MeasurementUnit.Piece,
            2m,
            2m,
            2m,
            "Материал: Заготовка 85х1020х40мм сталь 30ХГСА; код УТ000022134; остаток склада: 2 шт; норма: 1 шт на деталь.",
            [
                new ProductionLaunchShortageRow(1m, "ПВ/ПС438", "Р26666", new DateTime(2026, 8, 1)),
                new ProductionLaunchShortageRow(1m, "ПВ/ПС473", "Р26689", new DateTime(2026, 8, 2))
            ]);
        var editedComments = "Создано ПО Планирование Р26666; 1 шт\r\nСоздано ПО Планирование Р26689; 1 шт";

        var requests = MskViewModel.BuildProductionLaunchRequests(preview, 2m, editedComments, piecewise: true);

        requests.Select(x => x.Comment).Should().Equal(
            "Создано ПО Планирование Р26666; 1 шт",
            "Создано ПО Планирование Р26689; 1 шт");
    }

    [Fact]
    public async Task Msk_external_services_loads_piece_details_only_from_44_wip_section()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1657364", Designation = "01-P001-02.002-T25.040.RS", Name = "Направляющая задней бабки", RequiresNitriding = true };
        var snapshot = new StockSnapshot();
        db.Parts.Add(part);
        db.StockItems.AddRange(
            new StockItem
            {
                StockSnapshot = snapshot,
                OneCCode = "00001657364",
                SourceName = part.Name,
                Quantity = 6,
                Unit = MeasurementUnit.Piece,
                Warehouse = "44 секция НЗП (незавершенное производство)"
            },
            new StockItem
            {
                StockSnapshot = snapshot,
                OneCCode = "00001657364",
                SourceName = part.Name,
                Quantity = 2,
                Unit = MeasurementUnit.Piece,
                Warehouse = "Детали МУ в обработке на стороне"
            });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);

        var rows = await viewModel.LoadExternalServiceWipRowsAsync();

        rows.Should().ContainSingle();
        rows.Single().Ips.Should().Be("1657364");
        rows.Single().Designation.Should().Be("01-P001-02.002-T25.040.RS");
        rows.Single().AvailableQuantity.Should().Be(6);
        rows.Single().RequiresNitriding.Should().BeTrue();
    }

    [Fact]
    public async Task Msk_external_services_skips_wip_details_without_library_service_flags()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "1657364", Designation = "01-P001-02.002-T25.040.RS", Name = "Направляющая задней бабки", RequiresKeyway = true },
            new Part { Ips = "1816772", Designation = "270.011.02.001", Name = "Вал задней бабки" });
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, OneCCode = "00001657364", SourceName = "Направляющая задней бабки", Quantity = 2, Unit = MeasurementUnit.Piece, Warehouse = "44 секция НЗП (незавершенное производство)" },
            new StockItem { StockSnapshot = snapshot, OneCCode = "00001816772", SourceName = "Вал задней бабки", Quantity = 2, Unit = MeasurementUnit.Piece, Warehouse = "44 секция НЗП (незавершенное производство)" });
        await db.SaveChangesAsync();

        var rows = await new MskViewModel(db, NullLogger.Instance).LoadExternalServiceWipRowsAsync();

        rows.Should().ContainSingle(x => x.Ips == "1657364");
        rows.Should().NotContain(x => x.Ips == "1816772");
        MskViewModel.MatchesExternalServiceType(rows.Single(), "Шпоночный паз").Should().BeTrue();
        MskViewModel.MatchesExternalServiceType(rows.Single(), "Азотирование").Should().BeFalse();
    }

    [Fact]
    public void Msk_external_services_builds_goods_transfer_request_from_selected_wip_rows()
    {
        var row = new ExternalServiceWipRow("00001657364", "1657364", "01-P001-02.002-T25.040.RS", "Направляющая задней бабки", 6, "ПВ/ПС438", "Р26666", new DateTime(2026, 8, 1))
        {
            IsSelected = true,
            QuantityText = "2"
        };

        var ok = MskViewModel.TryBuildExternalServiceTransferRequest(
            "Азотирование",
            "Услуги на стороне: Азотирование",
            [row],
            out var request,
            out var error);

        ok.Should().BeTrue(error);
        request.Should().NotBeNull();
        request!.ServiceType.Should().Be("Азотирование");
        request.Comment.Should().Be("Услуги на стороне: Азотирование");
        request.Items.Should().ContainSingle();
        request.Items.Single().OneCCode.Should().Be("00001657364");
        request.Items.Single().Quantity.Should().Be(2);
    }

    [Fact]
    public async Task Msk_external_services_groups_wip_rows_and_keeps_demand_allocations()
    {
        await using var db = CreateDb();
        var part = new Part { Ips = "1657364", Designation = "01-P001-02.002-T25.040.RS", Name = "Направляющая задней бабки", RequiresNitriding = true };
        var snapshot = new StockSnapshot();
        var batch = new DemandBatch { ImportedAt = DateTime.UtcNow };
        db.Parts.Add(part);
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "00001657364",
            SourceName = part.Name,
            Quantity = 6,
            Unit = MeasurementUnit.Piece,
            Warehouse = "44 секция НЗП (незавершенное производство)"
        });
        db.DemandItems.AddRange(
            new DemandItem { DemandBatch = batch, Ips = "1657364", Part = part, Project = "ПВ/ПС438", ProductionSystem = "Р26666", Quantity = 2, DemandDate = new DateTime(2026, 8, 1) },
            new DemandItem { DemandBatch = batch, Ips = "1657364", Part = part, Project = "ПВ/ПС473", ProductionSystem = "Р26689", Quantity = 4, DemandDate = new DateTime(2026, 8, 2) });
        await db.SaveChangesAsync();

        var viewModel = new MskViewModel(db, NullLogger.Instance);

        var rows = await viewModel.LoadExternalServiceWipRowsAsync();

        rows.Should().ContainSingle();
        rows[0].Ips.Should().Be("1657364");
        rows[0].AvailableQuantity.Should().Be(6);
        rows[0].DemandAllocations.Should().HaveCount(2);
        rows[0].DemandAllocations[0].Project.Should().Be("ПВ/ПС438");
        rows[0].DemandAllocations[0].MachineNumber.Should().Be("Р26666");
        rows[0].DemandAllocations[0].Quantity.Should().Be(2);
        rows[0].DemandAllocations[1].Project.Should().Be("ПВ/ПС473");
        rows[0].DemandAllocations[1].MachineNumber.Should().Be("Р26689");
        rows[0].DemandAllocations[1].Quantity.Should().Be(4);
        rows[0].HasDemandRequirement.Should().BeTrue();
    }

    [Fact]
    public void Msk_external_services_uses_service_lead_time_for_request_date()
    {
        MskViewModel.CalculateExternalServiceReadyDate("Азотирование", new DateTime(2026, 8, 4))
            .Should().Be(new DateTime(2026, 8, 14));
        MskViewModel.CalculateExternalServiceReadyDate("Хим окс", new DateTime(2026, 8, 4))
            .Should().Be(new DateTime(2026, 8, 11));
    }

    [Fact]
    public void Msk_external_services_demand_quantity_ignores_stock_remainder_and_caps_by_available()
    {
        var withRemainder = new ExternalServiceWipRow(
            "00001657364",
            "1657364",
            "01-P001-02.002-T25.040.RS",
            "Направляющая задней бабки",
            6,
            string.Empty,
            string.Empty,
            null,
            [
                new ExternalServiceDemandAllocation("ПВ/ПС438", "Р26666", 2, new DateTime(2026, 8, 1)),
                new ExternalServiceDemandAllocation("ПВ/ПС473", "Р26689", 1, new DateTime(2026, 8, 2)),
                new ExternalServiceDemandAllocation("на склад", "на склад", 3, null)
            ]);
        var capped = new ExternalServiceWipRow(
            "00001657364",
            "1657364",
            "01-P001-02.002-T25.040.RS",
            "Направляющая задней бабки",
            4,
            string.Empty,
            string.Empty,
            null,
            [
                new ExternalServiceDemandAllocation("ПВ/ПС438", "Р26666", 3, new DateTime(2026, 8, 1)),
                new ExternalServiceDemandAllocation("ПВ/ПС473", "Р26689", 3, new DateTime(2026, 8, 2))
            ]);

        MskViewModel.CalculateExternalServiceDemandQuantity(withRemainder).Should().Be(3);
        MskViewModel.CalculateExternalServiceDemandQuantity(capped).Should().Be(4);
    }

    [Fact]
    public void Msk_external_services_comment_uses_selected_machine_numbers_and_quantities()
    {
        var rows = new[]
        {
            new ExternalServiceWipRow(
                "00001657364",
                "1657364",
                "01-P001-02.002-T25.040.RS",
                "Направляющая задней бабки",
                6,
                string.Empty,
                string.Empty,
                null,
                [
                    new ExternalServiceDemandAllocation("ПВ/ПС438", "Р26666", 2, new DateTime(2026, 8, 1)),
                    new ExternalServiceDemandAllocation("ПВ/ПС473", "Р26689", 4, new DateTime(2026, 8, 2))
                ])
            {
                IsSelected = true,
                QuantityText = "3"
            }
        };

        var comment = MskViewModel.BuildExternalServiceComment("Азотирование", rows);

        comment.Should().Be("Услуги на стороне: Азотирование IPS 1657364 Р26666 - 2 шт; Р26689 - 1 шт");
    }

    [Fact]
    public void Msk_external_services_exports_request_form()
    {
        var path = Path.Combine(Path.GetTempPath(), $"external-service-request-{Guid.NewGuid():N}.xlsx");
        var rows = new[]
        {
            new ExternalServiceWipRow(
                "00001657364",
                "1657364",
                "01-P001-02.002-T25.040.RS",
                "Направляющая задней бабки",
                6,
                string.Empty,
                string.Empty,
                null,
                [
                    new ExternalServiceDemandAllocation("ПВ/ПС438", "Р26666", 2, new DateTime(2026, 8, 1)),
                    new ExternalServiceDemandAllocation("ПВ/ПС473", "Р26689", 3, new DateTime(2026, 8, 2))
                ])
            {
                IsSelected = true,
                QuantityText = "6"
            }
        };

        MskViewModel.ExportExternalServiceRequestForm(path, "Азотирование", rows, new DateTime(2026, 7, 25));

        using var package = new ExcelPackage(new FileInfo(path));
        var sheet = package.Workbook.Worksheets["Заявка"];
        sheet.Should().NotBeNull();
        sheet!.Cells[1, 1].Text.Should().Be("№");
        sheet.Cells[1, 10].Text.Should().Be("Примечание");
        sheet.Cells[1, 12].Text.Should().Be("Стоимость, руб");
        sheet.Cells[2, 2].Text.Should().Be("ПВ/ПС438");
        sheet.Cells[2, 3].Text.Should().Be("Р26666");
        sheet.Cells[2, 4].Text.Should().Be("1657364");
        sheet.Cells[2, 6].Text.Should().Be("шт");
        sheet.Cells[2, 7].Value.Should().Be(2d);
        sheet.Cells[2, 8].Text.Should().Be("Азотирование");
        sheet.Cells[2, 9].GetValue<DateTime>().Should().Be(new DateTime(2026, 8, 4));
        sheet.Cells[2, 9].Style.Font.Bold.Should().BeTrue();
        sheet.Cells[2, 10].Text.Should().Be("СРОЧНО!");
        sheet.Cells[2, 10].Style.Font.Bold.Should().BeTrue();
        sheet.Cells[3, 2].Text.Should().Be("ПВ/ПС473");
        sheet.Cells[3, 3].Text.Should().Be("Р26689");
        sheet.Cells[3, 7].Value.Should().Be(3d);
        sheet.Cells[3, 10].Text.Should().Be("СРОЧНО!");
        sheet.Cells[4, 2].Text.Should().Be("на склад");
        sheet.Cells[4, 3].Text.Should().Be("на склад");
        sheet.Cells[4, 7].Value.Should().Be(1d);
        sheet.Cells[4, 9].GetValue<DateTime>().Should().Be(new DateTime(2026, 8, 4));
        sheet.Cells[4, 10].Text.Should().BeEmpty();
        File.Delete(path);
    }

    [Fact]
    public async Task Msk_open_drawing_command_uses_local_pdf_without_bridge_lookup()
    {
        await using var db = CreateDb();
        var drawingsDirectory = Path.Combine(Path.GetTempPath(), $"msk-drawings-{Guid.NewGuid():N}");
        var cacheDirectory = Path.Combine(Path.GetTempPath(), $"msk-cache-{Guid.NewGuid():N}");
        Directory.CreateDirectory(drawingsDirectory);
        Directory.CreateDirectory(cacheDirectory);
        var oldDrawingsDirectory = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_DIR");
        var oldCacheDirectory = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR");
        Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_DIR", drawingsDirectory);
        Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", cacheDirectory);
        var localPdf = Path.Combine(drawingsDirectory, "КПТ-04.00.002А01 Корпус.pdf");
        var cachedPdf = Path.Combine(cacheDirectory, "IPS_1320758.pdf");
        await File.WriteAllBytesAsync(localPdf, "%PDF-1.4\n%%EOF"u8.ToArray());

        db.MskRecords.Add(new MskRecord
        {
            Ips = "1320758",
            Designation = "КПТ-04.00.002А01",
            Name = "Корпус",
            FileName = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК\1320758.xlsx",
            ImportedAt = new DateTime(2026, 7, 27, 8, 0, 0, DateTimeKind.Utc)
        });
        await db.SaveChangesAsync();
        var drawingService = new CountingDrawingService();
        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance, drawingService);
        await viewModel.LoadCommand.ExecuteAsync(null);
        var row = viewModel.Rows.Single();

        try
        {
            await viewModel.OpenDrawingCommand.ExecuteAsync(row);

            drawingService.Calls.Should().Be(0);
            viewModel.DrawingViewerSource.Should().Be(new Uri(cachedPdf));
            viewModel.DrawingStatusText.Should().Contain("загружен и открыт");
        }
        finally
        {
            Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_DIR", oldDrawingsDirectory);
            Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", oldCacheDirectory);
            Directory.Delete(drawingsDirectory, recursive: true);
            Directory.Delete(cacheDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task Msk_update_adds_new_ips_to_library_and_attaches_nsi_blank()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D80 Сталь 40Х",
            CanonicalKey = "MSK|NSI|D80",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter,
            IsActive = true
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000022657",
            SourceName = "Круг D80 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016",
            NormalizedSourceName = "КРУГD80СТАЛЬ40Х",
            Source = "НСИ",
            IsActive = true
        });
        await db.SaveChangesAsync();

        const string fileName = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК\1320756.xlsx";
        var record = new MskRecord
        {
            Ips = "1320756",
            Designation = "КПТ-04.00.001А01",
            Name = "Втулка внутренняя",
            FileName = fileName,
            ImportedAt = new DateTime(2026, 7, 28, 8, 0, 0, DateTimeKind.Utc)
        };
        var details = new Dictionary<string, MskCsvDetail>(StringComparer.OrdinalIgnoreCase)
        {
            [fileName] = new("test", "1320756", "Круг", "Круг D80 Сталь 40Х", "Сталь 40Х", "УТ000022657", "0,065", "пог. м")
        };

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var added = await viewModel.AddMissingLibraryPartsFromMskAsync([record], details);
        await db.SaveChangesAsync();

        added.Should().Be(1);
        var part = await db.Parts.Include(x => x.BlankMaps).ThenInclude(x => x.CanonicalBlank).SingleAsync();
        part.Ips.Should().Be("1320756");
        part.Designation.Should().Be("КПТ-04.00.001А01");
        part.Name.Should().Be("Втулка внутренняя");
        part.HasMsk.Should().BeTrue();
        part.Source.Should().Be("МСК");
        var map = part.BlankMaps.Single();
        map.CanonicalBlank.Should().Be(blank);
        map.ConsumptionQuantity.Should().Be(0.065m);
        map.ConsumptionUnit.Should().Be(MeasurementUnit.Meter);
    }

    [Fact]
    public async Task Planning_update_adds_only_valid_numeric_ips_and_removes_invalid_msk_parts()
    {
        await using var db = CreateDb();
        db.Parts.AddRange(
            new Part { Ips = "-", Designation = "СТП-04-12-8М.30.32.001-20", Name = "Шкив", Source = "МСК", HasMsk = true },
            new Part { Ips = "463", Designation = "ТКИ028.001", Name = "Планшайба 250", Source = "МСК", HasMsk = true },
            new Part { Ips = "ZAGOTOVKA.PS.001.T630", Designation = "ZAGOTOVKA.PS.001.T630", Name = "ZAGOTOVKA.PS.001.T630", Source = "МСК", HasMsk = true },
            new Part { Ips = "1320756", Designation = "КПТ-04.00.001А01", Name = "Втулка внутренняя", Source = "Ручной ввод", HasMsk = true });
        await db.SaveChangesAsync();

        var records = new[]
        {
            new MskRecord { Ips = "1320758", Designation = "КПТ-04.00.002А01", Name = "Корпус", FileName = "1320758.xlsx", ImportedAt = DateTime.UtcNow },
            new MskRecord { Ips = "-", Designation = "СТП-04-12-8М.30.32.001-20", Name = "Шкив", FileName = "dash.xlsx", ImportedAt = DateTime.UtcNow },
            new MskRecord { Ips = "463", Designation = "ТКИ028.001", Name = "Планшайба 250", FileName = "463.xlsx", ImportedAt = DateTime.UtcNow },
            new MskRecord { Ips = "ZAGOTOVKA.PS.001.T630", Designation = "ZAGOTOVKA.PS.001.T630", Name = "ZAGOTOVKA.PS.001.T630", FileName = "zagotovka.xlsx", ImportedAt = DateTime.UtcNow }
        };

        var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance);
        var removed = await viewModel.RemoveInvalidMskLibraryPartsAsync();
        var added = await viewModel.AddMissingLibraryPartsFromMskAsync(records);
        await db.SaveChangesAsync();

        removed.Should().Be(3);
        added.Should().Be(1);
        var parts = await db.Parts.AsNoTracking().OrderBy(x => x.Ips).ToListAsync();
        parts.Select(x => x.Ips).Should().BeEquivalentTo(["1320756", "1320758"]);
        MskViewModel.IsValidMskLibraryIps("1320758").Should().BeTrue();
        MskViewModel.IsValidMskLibraryIps("00001320758").Should().BeTrue();
        MskViewModel.IsValidMskLibraryIps("463").Should().BeFalse();
        MskViewModel.IsValidMskLibraryIps("-").Should().BeFalse();
        MskViewModel.IsValidMskLibraryIps("ZAGOTOVKA.PS.001.T630").Should().BeFalse();
    }

    [Fact]
    public async Task Msk_selecting_row_does_not_open_pdf_until_command()
    {
        await using var db = CreateDb();
        const string ips = "9907771";
        db.MskRecords.Add(new MskRecord
        {
            Ips = ips,
            Designation = "TEST-9907771",
            Name = "Тестовая деталь с кэшем PDF",
            FileName = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК\9907771.xlsx",
            ImportedAt = new DateTime(2026, 7, 27, 8, 0, 0, DateTimeKind.Utc)
        });
        await db.SaveChangesAsync();

        var cacheDirectory = Path.Combine(Path.GetTempPath(), $"msk-cache-{Guid.NewGuid():N}");
        var oldCacheDirectory = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR");
        Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", cacheDirectory);
        Directory.CreateDirectory(cacheDirectory);
        var cachedPdf = Path.Combine(cacheDirectory, $"IPS_{ips}.pdf");
        await File.WriteAllBytesAsync(cachedPdf, "%PDF-1.4\n%%EOF"u8.ToArray());

        try
        {
            var drawingService = new CountingDrawingService();
            var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance, drawingService);
            await viewModel.LoadCommand.ExecuteAsync(null);
            var row = viewModel.Rows.Single();

            viewModel.SelectedRow = row;
            await Task.Delay(50);

            drawingService.Calls.Should().Be(0);
            viewModel.DrawingViewerSource.Should().BeNull();
            viewModel.DrawingStatusText.Should().NotContain("локально не найден");

            await viewModel.OpenDrawingCommand.ExecuteAsync(row);

            viewModel.DrawingViewerSource.Should().Be(new Uri(cachedPdf));
            viewModel.DrawingStatusText.Should().Contain("открыт из кэша");
        }
        finally
        {
            Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", oldCacheDirectory);
            Directory.Delete(cacheDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task Msk_open_drawing_command_uses_bridge_when_local_pdf_is_missing()
    {
        await using var db = CreateDb();
        const string ips = "9908881";
        db.MskRecords.Add(new MskRecord
        {
            Ips = ips,
            Designation = "TEST-9908881",
            Name = "Тестовая деталь без локального PDF",
            FileName = @"X:\19_МЕХ УЧАСТОК\База МСК\СПИСОК МСК\9908881.xlsx",
            ImportedAt = new DateTime(2026, 7, 28, 8, 0, 0, DateTimeKind.Utc)
        });
        await db.SaveChangesAsync();

        var cacheDirectory = Path.Combine(Path.GetTempPath(), $"msk-cache-{Guid.NewGuid():N}");
        var oldCacheDirectory = Environment.GetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR");
        Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", cacheDirectory);
        Directory.CreateDirectory(cacheDirectory);
        var cachedPdf = Path.Combine(cacheDirectory, $"IPS_{ips}.pdf");

        try
        {
            var drawingService = new CountingDrawingService { CreatePdf = true };
            var viewModel = new MskViewModel(db, Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance, drawingService);
            await viewModel.LoadCommand.ExecuteAsync(null);

            await viewModel.OpenDrawingCommand.ExecuteAsync(viewModel.Rows.Single());

            drawingService.Calls.Should().Be(1);
            drawingService.Queries.Should().Contain(ips);
            File.Exists(cachedPdf).Should().BeTrue();
            viewModel.DrawingViewerSource.Should().Be(new Uri(cachedPdf));
            viewModel.DrawingStatusText.Should().Contain("загружен и открыт");
        }
        finally
        {
            Environment.SetEnvironmentVariable("BLANK_DEMAND_DRAWINGS_CACHE_DIR", oldCacheDirectory);
            Directory.Delete(cacheDirectory, recursive: true);
        }
    }

    [Fact]
    public async Task Demand_load_matches_work_in_progress_when_one_c_code_has_leading_zeros()
    {
        await using var db = CreateDb();
        var batch = new DemandBatch { Name = "Потребность" };
        db.DemandItems.Add(new DemandItem
        {
            DemandBatch = batch,
            Ips = "1799544",
            SourcePartName = "01-P002-03.002-T63.150.SS Направляющая Z верхняя стыкуемая",
            SerialNumber = "Р25632",
            Quantity = 1,
            Unit = "шт",
            DemandDate = new DateTime(2026, 7, 21)
        });
        var snapshot = new StockSnapshot();
        db.StockItems.Add(new StockItem
        {
            StockSnapshot = snapshot,
            OneCCode = "00001799544",
            SourceName = "01-P002-03.002-T63.150.SS Направляющая Z верхняя стыкуемая",
            Quantity = 2,
            Unit = MeasurementUnit.Piece
        });
        await db.SaveChangesAsync();

        var viewModel = new DemandViewModel(db);
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Single().InProductionQuantity.Should().Be("2");
    }

    [Fact]
    public async Task Nsi_loads_one_c_blanks_and_filters_by_size_type_and_material()
    {
        await using var db = CreateDb();
        var round60 = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D60 РЎС‚Р°Р»СЊ 40РҐ",
            CanonicalKey = "NSI|ROUND|60|40X",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            Material = "40РҐ",
            BaseUnit = MeasurementUnit.Meter
        };
        var square100 = new CanonicalBlank
        {
            CanonicalName = "РљРІР°РґСЂР°С‚ 100 РЎС‚Р°Р»СЊ 20",
            CanonicalKey = "NSI|SQUARE|100|20",
            BlankType = BlankType.SquareBar,
            WidthMm = 100,
            Material = "20",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round60, OneCCode = "UT060", SourceName = round60.CanonicalName, NormalizedSourceName = round60.CanonicalName, Source = "1C" },
            new BlankAlias { CanonicalBlank = square100, OneCCode = "UT100", SourceName = square100.CanonicalName, NormalizedSourceName = square100.CanonicalName, Source = "1C" });
        await db.SaveChangesAsync();

        var part = new Part { Ips = "500001", Designation = "D-1", Name = "Р”РµС‚Р°Р»СЊ РёР· РєСЂСѓРіР°" };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = round60,
            ConsumptionQuantity = 0.35m,
            ConsumptionUnit = MeasurementUnit.Meter,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService())
        {
            SizeFilter = "60",
            MaterialFilter = "40РҐ"
        };
        viewModel.SelectedBlankTypeFilter = viewModel.BlankTypeFilters.Single(x => x.Value == BlankType.RoundBar);
        await viewModel.LoadCommand.ExecuteAsync(null);

        viewModel.Rows.Should().ContainSingle();
        viewModel.Rows.Single().OneCCode.Should().Be("UT060");
        viewModel.Rows.Single().UnitName.Should().Be("пог. м");
        viewModel.Rows.Single().BlankType.Should().Be("Круг");
        viewModel.Rows.Single().Size.Should().Contain("D60");
        viewModel.Rows.Single().StockQuantity.Should().Be("0");

        await viewModel.LoadUsageCommand.ExecuteAsync(viewModel.Rows.Single());
        viewModel.UsageRows.Should().ContainSingle();
        viewModel.UsageRows.Single().Ips.Should().Be("500001");
        viewModel.UsageRows.Single().Quantity.Should().Be("0,35");
        viewModel.UsageRows.Single().UnitName.Should().Be("пог. м");
    }

    [Fact]
    public async Task Nsi_load_shows_cmo_and_warehouse_stock_separately()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60",
            CanonicalKey = "NSI|STOCK|SPLIT",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            BaseUnit = MeasurementUnit.Meter
        };
        var alias = new BlankAlias { CanonicalBlank = blank, OneCCode = "UT-SPLIT", SourceName = blank.CanonicalName, NormalizedSourceName = blank.CanonicalName, Source = "1C" };
        var snapshot = new StockSnapshot();
        db.StockItems.AddRange(
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 2, Unit = MeasurementUnit.Meter, Warehouse = "44 секция НЗП (незавершенное производство)" },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 3, Unit = MeasurementUnit.Meter, Warehouse = "Детали МУ в обработке на стороне" },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 5, Unit = MeasurementUnit.Meter, Warehouse = "Основной склад" },
            new StockItem { StockSnapshot = snapshot, BlankAlias = alias, OneCCode = alias.OneCCode, SourceName = alias.SourceName, Quantity = 4, Unit = MeasurementUnit.Meter, Warehouse = "ТМЦ по проекту ФРП" });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await viewModel.LoadCommand.ExecuteAsync(null);

        var row = viewModel.Rows.Single(x => x.OneCCode == "UT-SPLIT");
        row.CmoStockQuantity.Should().Be("5000");
        row.WarehouseStockQuantity.Should().Be("9000");
        row.StockUnitName.Should().Be("мм");
        row.Price.Should().BeEmpty();
    }

    [Fact]
    public async Task Nsi_load_shows_one_c_price_with_currency_in_single_column()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D60",
            CanonicalKey = "NSI|PRICE|D60",
            BlankType = BlankType.RoundBar,
            DiameterMm = 60,
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000011939",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "1C"
        });
        db.OneCPriceItems.Add(new OneCPriceItem
        {
            LookupKey = StockCodeNormalizer.NormalizeForComparison("УТ000011939"),
            Code = "УТ000011939",
            Price = 100m,
            Currency = "RUB",
            PriceType = "Закупочная"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService(), new StubOneCNomenclatureService([]));

        await viewModel.LoadAsync();

        viewModel.Rows.Single().Price.Should().Be("100 ₽");
    }


    [Fact]
    public async Task Nsi_load_refreshes_source_name_from_one_c_by_code()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг ф 115 сталь 40Х",
            CanonicalKey = "ROUND|40X|D115",
            BlankType = BlankType.RoundBar,
            Material = "Сталь 40Х",
            DiameterMm = 115,
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000022886",
            SourceName = "Круг ф 115 сталь 40Х",
            NormalizedSourceName = "Круг ф 115 сталь 40Х",
            Source = "Библиотека.xlsx",
            IsActive = true
        });
        await db.SaveChangesAsync();

        var oneC = new StubOneCNomenclatureService([
            new OneCNomenclatureItem("УТ000022886", "", "Круг D115 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016", "пог. м")
        ]);
        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService(), oneC);

        await viewModel.LoadAsync();

        viewModel.Rows.Single().SourceName.Should().Be("Круг D115 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016");
        var updatedAlias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync();
        updatedAlias.SourceName.Should().Be("Круг D115 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016");
        updatedAlias.CanonicalBlank!.CanonicalName.Should().Be("Круг D115 Сталь 40Х ГОСТ 2590-2006 / ГОСТ 4543-2016");
    }

    [Fact]
    public async Task Nsi_explicit_one_c_refresh_updates_name_after_initial_open()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D115 старое имя",
            CanonicalKey = "ROUND|40X|D115",
            BlankType = BlankType.RoundBar,
            Material = "Сталь 40Х",
            DiameterMm = 115,
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "УТ000022886",
            SourceName = "Круг D115 старое имя",
            NormalizedSourceName = "Круг D115 старое имя",
            Source = "Библиотека.xlsx",
            IsActive = true
        });
        await db.SaveChangesAsync();

        var oneCItems = new List<OneCNomenclatureItem>
        {
            new("УТ000022886", "", "Круг D115 первое имя 1С", "пог. м")
        };
        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService(), new StubOneCNomenclatureService(oneCItems));

        await viewModel.LoadAsync();
        oneCItems.Clear();
        oneCItems.Add(new OneCNomenclatureItem("УТ000022886", "", "Круг D115 обновленное имя 1С", "пог. м"));

        var updated = await viewModel.RefreshNamesFromOneCAsync(CancellationToken.None);
        await viewModel.LoadAsync();

        updated.Should().Be(1);
        viewModel.Rows.Single().SourceName.Should().Be("Круг D115 обновленное имя 1С");
        var updatedAlias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync();
        updatedAlias.SourceName.Should().Be("Круг D115 обновленное имя 1С");
        updatedAlias.CanonicalBlank!.CanonicalName.Should().Be("Круг D115 обновленное имя 1С");
    }

    [Fact]
    public async Task Nsi_usage_does_not_show_parts_from_similar_but_different_blank()
    {
        await using var db = CreateDb();
        var oneCBlank = new CanonicalBlank
        {
            CanonicalName = "Р—Р°РіРѕС‚РѕРІРєР° 60С…174С…970РјРј РЎС‚Р°Р»СЊ 30РҐР“РЎРђ",
            CanonicalKey = "NSI|UT|22800",
            BlankType = BlankType.Plate,
            WidthMm = 60,
            HeightMm = 174,
            LengthMm = 970,
            Material = "30РҐР“РЎРђ"
        };
        var libraryBlank = new CanonicalBlank
        {
            CanonicalName = "Р—Р°РіРѕС‚РѕРІРєР° 60С…174С…970РјРј РЎС‚Р°Р»СЊ 30РҐР“РЎРђ",
            CanonicalKey = "LIB|UT|22800",
            BlankType = BlankType.Plate,
            WidthMm = 60,
            HeightMm = 174,
            LengthMm = 970,
            Material = "30РҐР“РЎРђ"
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = oneCBlank,
            OneCCode = "РЈРў000022800",
            SourceName = "Р—Р°РіРѕС‚РѕРІРєР° 60С…174С…970РјРј РЎС‚Р°Р»СЊ 30РҐР“РЎРђ",
            NormalizedSourceName = "Р—Р°РіРѕС‚РѕРІРєР° 60С…174С…970РјРј РЎС‚Р°Р»СЊ 30РҐР“РЎРђ",
            Source = "1C"
        });
        var part = new Part
        {
            Ips = "1720202",
            Name = "01-P006-01.007-T63.200.SS (РќР°РїСЂР°РІР»СЏСЋС‰Р°СЏ Р—Р‘ РІРµСЂС…РЅСЏСЏ СЃС‚С‹РєСѓРµРјР°СЏ)"
        };
        db.PartBlankMaps.Add(new PartBlankMap
        {
            Part = part,
            CanonicalBlank = libraryBlank,
            ConsumptionQuantity = 1,
            ConsumptionUnit = MeasurementUnit.Piece,
            Source = "library"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService())
        {
            Search = "РЈРў000022800"
        };
        await viewModel.LoadCommand.ExecuteAsync(null);
        await viewModel.LoadUsageCommand.ExecuteAsync(viewModel.Rows.Single());

        viewModel.UsageRows.Should().BeEmpty();
        viewModel.UsageStatusText.Should().Be("Применяемость не найдена");
    }

    [Fact]
    public async Task Nsi_delete_selected_removes_alias_from_database()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "РљСЂСѓРі D20",
            CanonicalKey = "NSI|ARCHIVE|D20",
            BlankType = BlankType.RoundBar,
            DiameterMm = 20
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "UTARCHIVE",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.SelectedRow = viewModel.Rows.Single();

        await viewModel.DeleteSelectedCommand.ExecuteAsync(null);

        viewModel.Rows.Should().BeEmpty();
        (await db.BlankAliases.CountAsync()).Should().Be(0);
        (await db.CanonicalBlanks.CountAsync()).Should().Be(0);
    }

    [Fact]
    public async Task Nsi_delete_selected_can_be_restored_by_undo()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Круг D25",
            CanonicalKey = "NSI|UNDO|D25",
            BlankType = BlankType.RoundBar,
            DiameterMm = 25,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "UTUNDO",
            SourceName = blank.CanonicalName,
            NormalizedSourceName = blank.CanonicalName,
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.SelectedRow = viewModel.Rows.Single();

        await viewModel.DeleteSelectedCommand.ExecuteAsync(null);
        await UndoCenter.UndoAsync();

        var alias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync();
        alias.OneCCode.Should().Be("UTUNDO");
        alias.CanonicalBlank.Should().NotBeNull();
        alias.CanonicalBlank!.CanonicalName.Should().Be("Круг D25");
        alias.CanonicalBlank.Material.Should().Be("Сталь 40Х");
        viewModel.Rows.Should().ContainSingle(x => x.OneCCode == "UTUNDO");
        viewModel.StatusText.Should().Be("Восстановлено позиций НСИ: 1.");
    }

    [Fact]
    public async Task Nsi_editor_updates_code_name_type_and_material()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Старая заготовка",
            CanonicalKey = "NSI|EDIT|MANUAL",
            BlankType = BlankType.Unknown,
            BaseUnit = MeasurementUnit.Piece
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "OLD001",
            SourceName = "Старая заготовка",
            NormalizedSourceName = "Старая заготовка",
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.SelectedRow = viewModel.Rows.Single();
        viewModel.EditOneCCode = "NEW001";
        viewModel.EditSourceName = "Новая заготовка";
        viewModel.EditBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.Forging);
        viewModel.EditMaterial = "40Х";

        await viewModel.SaveSelectedEditCommand.ExecuteAsync(null);

        var alias = await db.BlankAliases.Include(x => x.CanonicalBlank).SingleAsync();
        alias.OneCCode.Should().Be("NEW001");
        alias.SourceName.Should().Be("Новая заготовка");
        alias.CanonicalBlank!.CanonicalName.Should().Be("Новая заготовка");
        alias.CanonicalBlank.BlankType.Should().Be(BlankType.Forging);
        alias.CanonicalBlank.Material.Should().Be("40Х");
        viewModel.StatusText.Should().Be("Деталь добавлена в список: NEW001.");
    }

    [Fact]
    public async Task Nsi_editor_suggestions_match_cyrillic_case_insensitive()
    {
        await using var db = CreateDb();
        db.CanonicalBlanks.Add(new CanonicalBlank
        {
            CanonicalName = "Круг D80 Сталь 40Х",
            CanonicalKey = "NSI|SUGGESTIONS|1",
            BlankType = BlankType.RoundBar,
            BaseUnit = MeasurementUnit.Meter,
            Material = "Сталь 40Х",
            MaterialGost = "ГОСТ 4543-2016",
            ProfileGost = "ГОСТ 2590-2006"
        });
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = db.CanonicalBlanks.Local.Single(),
            OneCCode = "UT-SUGGEST",
            SourceName = "Круг D80 Сталь 40Х",
            NormalizedSourceName = "Круг D80 Сталь 40Х",
            Source = "1C"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());

        await viewModel.LoadMaterialSuggestionsCommand.ExecuteAsync("сталь");
        viewModel.EditMaterialGost = "4543";
        await Task.Delay(100);
        viewModel.EditProfileGost = "2590";
        await Task.Delay(100);

        viewModel.MaterialSuggestions.Should().Contain("Сталь 40Х");
        viewModel.MaterialGostSuggestions.Should().Contain("ГОСТ 4543-2016");
        viewModel.ProfileGostSuggestions.Should().Contain("ГОСТ 2590-2006");
    }

    [Fact]
    public async Task Nsi_editor_material_suggestions_use_only_active_nsi_aliases()
    {
        await using var db = CreateDb();
        db.CanonicalBlanks.Add(new CanonicalBlank
        {
            CanonicalName = "Библиотечная заготовка без НСИ",
            CanonicalKey = "NO-NSI-MATERIAL",
            BlankType = BlankType.RoundBar,
            Material = "Материал не из НСИ",
            BaseUnit = MeasurementUnit.Meter
        });
        var nsiBlank = new CanonicalBlank
        {
            CanonicalName = "Круг D80 Сталь 40Х",
            CanonicalKey = "NSI|MATERIAL|ACTIVE",
            BlankType = BlankType.RoundBar,
            Material = "Сталь 40Х",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = nsiBlank,
            OneCCode = "UT-MATERIAL",
            SourceName = nsiBlank.CanonicalName,
            NormalizedSourceName = nsiBlank.CanonicalName,
            Source = "1C"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());

        await viewModel.LoadMaterialSuggestionsCommand.ExecuteAsync(string.Empty);

        viewModel.MaterialSuggestions.Should().Contain("Сталь 40Х");
        viewModel.MaterialSuggestions.Should().NotContain("Материал не из НСИ");
    }

    [Fact]
    public async Task Nsi_editor_auto_fills_gosts_from_active_nsi_without_overwriting_manual_values()
    {
        await using var db = CreateDb();
        var round40 = new CanonicalBlank
        {
            CanonicalName = "Круг D80 Сталь 40Х",
            CanonicalKey = "NSI|AUTO-GOST|40H",
            BlankType = BlankType.RoundBar,
            Material = "Сталь 40Х",
            MaterialGost = "ГОСТ 4543-2016",
            ProfileGost = "ГОСТ 2590-2006",
            BaseUnit = MeasurementUnit.Meter
        };
        var round20 = new CanonicalBlank
        {
            CanonicalName = "Круг D80 Сталь 20",
            CanonicalKey = "NSI|AUTO-GOST|20",
            BlankType = BlankType.RoundBar,
            Material = "Сталь 20",
            MaterialGost = "ГОСТ 1050-2013",
            ProfileGost = "ГОСТ 2590-2006",
            BaseUnit = MeasurementUnit.Meter
        };
        db.BlankAliases.AddRange(
            new BlankAlias { CanonicalBlank = round40, OneCCode = "UT-GOST-40", SourceName = round40.CanonicalName, NormalizedSourceName = round40.CanonicalName, Source = "1C" },
            new BlankAlias { CanonicalBlank = round20, OneCCode = "UT-GOST-20", SourceName = round20.CanonicalName, NormalizedSourceName = round20.CanonicalName, Source = "1C" });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());

        viewModel.EditMaterial = "сталь 40х";
        viewModel.EditBlankType = viewModel.BlankTypes.Single(x => x.Value == BlankType.RoundBar);
        for (var attempt = 0; attempt < 20 && (viewModel.EditMaterialGost.Length == 0 || viewModel.EditProfileGost.Length == 0); attempt++)
        {
            await Task.Delay(50);
        }

        viewModel.EditMaterialGost.Should().Be("ГОСТ 4543-2016");
        viewModel.EditProfileGost.Should().Be("ГОСТ 2590-2006");

        viewModel.EditMaterialGost = "Ручной ГОСТ";
        viewModel.EditMaterial = "Сталь 20";
        await Task.Delay(150);

        viewModel.EditMaterialGost.Should().Be("Ручной ГОСТ");
    }

    [Fact]
    public async Task Nsi_editor_accepts_display_size_format_with_named_dimensions()
    {
        await using var db = CreateDb();
        var blank = new CanonicalBlank
        {
            CanonicalName = "Плита 60x174x970",
            CanonicalKey = "NSI|EDIT|SIZE",
            BlankType = BlankType.Plate,
            WidthMm = 60,
            HeightMm = 174,
            LengthMm = 970,
            BaseUnit = MeasurementUnit.Piece
        };
        db.BlankAliases.Add(new BlankAlias
        {
            CanonicalBlank = blank,
            OneCCode = "SIZE001",
            SourceName = "Плита 60x174x970",
            NormalizedSourceName = "Плита 60x174x970",
            Source = "test"
        });
        await db.SaveChangesAsync();

        var viewModel = new NormalizationViewModel(db, new StubExcelImportService(), new StubFileDialogService());
        await viewModel.LoadCommand.ExecuteAsync(null);
        viewModel.SelectedRow = viewModel.Rows.Single();
        viewModel.EditSize.Should().BeEmpty();
        viewModel.EditNsiRowCommand.Execute(viewModel.SelectedRow);
        viewModel.EditSize.Should().Be("W60 H174 L970");
        viewModel.EditMaterial = "30ХГСА";

        await viewModel.SaveSelectedEditCommand.ExecuteAsync(null);

        var saved = await db.CanonicalBlanks.SingleAsync();
        saved.WidthMm.Should().Be(60);
        saved.HeightMm.Should().Be(174);
        saved.LengthMm.Should().Be(970);
        saved.Material.Should().Be("30ХГСА");
    }

    private static BlankDemandPlannerDbContext CreateDb()
    {
        var options = new DbContextOptionsBuilder<BlankDemandPlannerDbContext>()
            .UseSqlite($"Data Source={Path.Combine(Path.GetTempPath(), $"bdp_library_{Guid.NewGuid():N}.db")};Pooling=False")
            .Options;
        var db = new BlankDemandPlannerDbContext(options);
        db.Database.Migrate();
        return db;
    }

    private sealed class StubFileDialogService : IFileDialogService
    {
        public string? OpenExcelFile() => null;
        public string? SelectFolder() => null;
    }

    private sealed class StubExcelImportService : IExcelImportService
    {
        public Task<IReadOnlyList<string>> GetSheetNamesAsync(string filePath, CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<string>>([]);
        public Task<ExcelPreview> PreviewAsync(string filePath, string sheetName, CancellationToken cancellationToken) => Task.FromResult(new ExcelPreview(filePath, sheetName, 1, [], []));
        public Task<ImportReport> ImportOneCBlanksAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken) => Task.FromResult(new ImportReport(0, 0, 0, 0, 0, []));
        public Task<ImportReport> ImportDemandAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken) => Task.FromResult(new ImportReport(0, 0, 0, 0, 0, []));
        public Task<ImportReport> ImportStockAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken) => Task.FromResult(new ImportReport(0, 0, 0, 0, 0, []));
        public Task<ImportReport> ImportManufacturingBlankLibraryAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken) => Task.FromResult(new ImportReport(0, 0, 0, 0, 0, []));
    }

    private sealed class StubOneCNomenclatureService(
        IReadOnlyList<OneCNomenclatureItem> items,
        IReadOnlyList<OneCNomenclaturePrice>? prices = null) : IOneCNomenclatureService
    {
        public Task<IReadOnlyList<OneCNomenclatureItem>> SearchAsync(string query, int limit, CancellationToken cancellationToken) =>
            Task.FromResult<IReadOnlyList<OneCNomenclatureItem>>(items
                .Where(x =>
                    x.Code.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                    x.Article.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                    x.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
                .Take(limit)
                .ToList());

        public Task<IReadOnlyList<OneCNomenclatureItem>> ResolveByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken)
        {
            var keys = codes.Select(StockCodeNormalizer.NormalizeForComparison).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return Task.FromResult<IReadOnlyList<OneCNomenclatureItem>>(items
                .Where(x =>
                    keys.Contains(StockCodeNormalizer.NormalizeForComparison(x.Code)) ||
                    keys.Contains(StockCodeNormalizer.NormalizeForComparison(x.Article)))
                .ToList());
        }

        public Task<IReadOnlyList<OneCNomenclaturePrice>> ResolvePricesByCodesAsync(IEnumerable<string> codes, CancellationToken cancellationToken)
        {
            var keys = codes.Select(StockCodeNormalizer.NormalizeForComparison).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return Task.FromResult<IReadOnlyList<OneCNomenclaturePrice>>((prices ?? [])
                .Where(x =>
                    keys.Contains(StockCodeNormalizer.NormalizeForComparison(x.Code)) ||
                    keys.Contains(StockCodeNormalizer.NormalizeForComparison(x.Article)))
                .ToList());
        }
    }

    private sealed class CountingDrawingService : IIpsDrawingService
    {
        public int Calls { get; private set; }
        public bool CreatePdf { get; init; }
        public List<string> Queries { get; } = [];

        public async Task<FileInfo?> FindDrawingPdfAsync(string query, DirectoryInfo outputDirectory, CancellationToken cancellationToken)
        {
            Calls++;
            Queries.Add(query);
            if (!CreatePdf)
            {
                return null;
            }

            outputDirectory.Create();
            var path = Path.Combine(outputDirectory.FullName, $"{query}.pdf");
            await File.WriteAllBytesAsync(path, "%PDF-1.4\n%%EOF"u8.ToArray(), cancellationToken);
            return new FileInfo(path);
        }
    }
}
