using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using BlankDemandPlanner.Core.Entities;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using BlankDemandPlanner.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace BlankDemandPlanner.Infrastructure.Excel;

public sealed class ExcelImportService(
    BlankDemandPlannerDbContext dbContext,
    IBlankNormalizationService normalizationService,
    II012ValidationService validationService,
    ILogger<ExcelImportService> logger) : IExcelImportService
{
    static ExcelImportService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<IReadOnlyList<string>> GetSheetNamesAsync(string filePath, CancellationToken cancellationToken)
    {
        await using var stream = File.OpenRead(filePath);
        using var package = new ExcelPackage(stream);
        cancellationToken.ThrowIfCancellationRequested();
        return package.Workbook.Worksheets.Select(x => x.Name).ToList();
    }

    public async Task<ExcelPreview> PreviewAsync(string filePath, string sheetName, CancellationToken cancellationToken)
    {
        await using var stream = File.OpenRead(filePath);
        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[sheetName] ?? throw new InvalidOperationException($"Лист '{sheetName}' не найден.");
        var headerRow = FindHeaderRow(sheet);
        var headers = ReadRow(sheet, headerRow);
        var rows = new List<IReadOnlyList<string>>();
        for (var row = headerRow + 1; row <= Math.Min(sheet.Dimension.End.Row, headerRow + 100); row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rows.Add(ReadRow(sheet, row, headers.Count));
        }

        return new ExcelPreview(filePath, sheetName, headerRow, headers, rows);
    }

    public async Task<ImportReport> ImportOneCBlanksAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken)
    {
        logger.LogInformation("Import OneC blanks start from {File}", filePath);
        var report = new MutableImportReport();
        var hash = await ComputeSha256Async(filePath, cancellationToken);
        await using var transaction = await dbContext.Database.BeginTransactionAsync(cancellationToken);
        var run = CreateRun("OneCBlank", filePath, hash);
        dbContext.ImportRuns.Add(run);
        await dbContext.SaveChangesAsync(cancellationToken);

        await foreach (var row in EnumerateOneCBlankRowsAsync(filePath, cancellationToken))
        {
            report.ReadRows++;
            var code = Get(row, "OneCCode");
            var name = Get(row, "Name");
            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name))
            {
                report.SkippedRows++;
                continue;
            }

            var normalized = await normalizationService.NormalizeAsync(name, cancellationToken);
            var status = await validationService.ValidateAsync(normalized, cancellationToken);
            var canonical = await GetOrCreateCanonicalBlankAsync(normalized, status, cancellationToken);
            var alias = await FindAliasAsync(code, cancellationToken);
            if (alias is null)
            {
                dbContext.BlankAliases.Add(new BlankAlias
                {
                    CanonicalBlank = canonical,
                    OneCCode = code,
                    SourceName = name,
                    NormalizedSourceName = normalized.NormalizedName,
                    Source = Path.GetFileName(filePath)
                });
                report.AddedRows++;
            }
            else
            {
                alias.CanonicalBlank = canonical;
                alias.SourceName = name;
                alias.NormalizedSourceName = normalized.NormalizedName;
                alias.UpdatedAt = DateTime.UtcNow;
                report.UpdatedRows++;
            }

            await SaveBatchAsync(report, progress, cancellationToken);
        }

        await FinishRunAsync(run, report, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return report.ToImmutable();
    }

    public async Task<ImportReport> ImportManufacturingBlankLibraryAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken)
    {
        logger.LogInformation("Import manufacturing blank library start from {File}", filePath);
        var report = new MutableImportReport();
        await using var transaction = await dbContext.Database.BeginTransactionAsync(cancellationToken);
        var run = CreateRun("ManufacturingBlankLibrary", filePath, await ComputeSha256Async(filePath, cancellationToken));
        dbContext.ImportRuns.Add(run);

        var manufacturingProfiles = new IExcelImportProfile[] { new MatchedOrderBlankImportProfile(), new GuideBlankImportProfile(), new RotationalBlankImportProfile(), new PipeBlankImportProfile(), new PlateBlankImportProfile() };
        foreach (var profile in manufacturingProfiles)
        {
            var readRowsBeforeProfile = report.ReadRows;
            await foreach (var row in EnumerateRowsAsync(filePath, profile, cancellationToken))
            {
                report.ReadRows++;
                var ips = Get(row, "Ips");
                if (string.IsNullOrWhiteSpace(ips))
                {
                    report.SkippedRows++;
                    continue;
                }

                var partName = FirstNotEmpty(Get(row, "PartName"), Get(row, "НАИМЕНОВАНИЕ"), "Без наименования");
                var part = await FindPartAsync(ips, cancellationToken);
                if (part is null)
                {
                    part = new Part { Ips = ips, Designation = Get(row, "Designation"), Name = partName, Source = Path.GetFileName(filePath) };
                    dbContext.Parts.Add(part);
                    report.AddedRows++;
                }
                else
                {
                    part.Name = partName;
                    part.Designation = FirstNotEmpty(Get(row, "Designation"), part.Designation);
                    part.UpdatedAt = DateTime.UtcNow;
                    report.UpdatedRows++;
                }

                var blankName = FirstNotEmpty(Get(row, "BlankName"), BuildBlankName(row));
                if (!string.IsNullOrWhiteSpace(blankName))
                {
                    var code = Get(row, "OneCCode");
                    var existingAlias = string.IsNullOrWhiteSpace(code) ? null : await FindAliasAsync(code, cancellationToken);
                    var normalized = await normalizationService.NormalizeAsync(blankName, cancellationToken);
                    var status = await validationService.ValidateAsync(normalized, cancellationToken);
                    var blank = existingAlias?.CanonicalBlank ?? await GetOrCreateCanonicalBlankAsync(normalized, status, cancellationToken);
                    var activeMap = await FindPartBlankMapAsync(part, blank, cancellationToken);
                    if (activeMap is null)
                    {
                        await DeactivateOtherPartBlankMapsAsync(part, null, cancellationToken);
                        dbContext.PartBlankMaps.Add(new PartBlankMap
                        {
                            Part = part,
                            CanonicalBlank = blank,
                            ConsumptionQuantity = ParseDecimal(Get(row, "Quantity")) ?? 1m,
                            ConsumptionUnit = string.IsNullOrWhiteSpace(Get(row, "Unit")) ? blank.BaseUnit : ParseUnit(Get(row, "Unit")),
                            BlankLeadTimeDays = ParseLeadTimeDays(Get(row, "BlankLeadTimeDays")),
                            Source = profile.ProfileType,
                            SourceFile = Path.GetFileName(filePath),
                            IsPrimary = true
                        });
                    }
                    else
                    {
                        await DeactivateOtherPartBlankMapsAsync(part, activeMap.Id, cancellationToken);
                        activeMap.ConsumptionQuantity = ParseDecimal(Get(row, "Quantity")) ?? activeMap.ConsumptionQuantity;
                        activeMap.ConsumptionUnit = string.IsNullOrWhiteSpace(Get(row, "Unit")) ? activeMap.ConsumptionUnit : ParseUnit(Get(row, "Unit"));
                        activeMap.BlankLeadTimeDays = ParseLeadTimeDays(Get(row, "BlankLeadTimeDays"), activeMap.BlankLeadTimeDays);
                        activeMap.Source = profile.ProfileType;
                        activeMap.SourceFile = Path.GetFileName(filePath);
                        activeMap.IsPrimary = true;
                        activeMap.UpdatedAt = DateTime.UtcNow;
                    }

                    if (existingAlias is not null && !string.IsNullOrWhiteSpace(blankName))
                    {
                        existingAlias.SourceName = blankName;
                        existingAlias.NormalizedSourceName = normalized.NormalizedName;
                        existingAlias.Source = Path.GetFileName(filePath);
                        existingAlias.UpdatedAt = DateTime.UtcNow;
                    }

                    if (!string.IsNullOrWhiteSpace(code) && existingAlias is null)
                    {
                        dbContext.BlankAliases.Add(new BlankAlias
                        {
                            CanonicalBlank = blank,
                            OneCCode = code,
                            SourceName = blankName,
                            NormalizedSourceName = normalized.NormalizedName,
                            Source = Path.GetFileName(filePath)
                        });
                    }
                }

                await SaveBatchAsync(report, progress, cancellationToken);
            }

            if (profile is MatchedOrderBlankImportProfile && report.ReadRows > readRowsBeforeProfile)
            {
                break;
            }
        }

        await FinishRunAsync(run, report, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return report.ToImmutable();
    }

    public async Task<ImportReport> ImportDemandAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken)
    {
        var report = new MutableImportReport();
        var hash = await ComputeSha256Async(filePath, cancellationToken);
        var existingBatch = await dbContext.DemandBatches.FirstOrDefaultAsync(x => x.FileHashSha256 == hash, cancellationToken);
        if (existingBatch is not null)
        {
            return await RefreshExistingDemandBatchAsync(filePath, existingBatch, hash, progress, cancellationToken);
        }

        if (await dbContext.DemandBatches.AnyAsync(x => x.FileHashSha256 == hash, cancellationToken))
        {
            return new ImportReport(0, 0, 0, 0, 1, ["Файл потребности уже импортирован. Повторный незаметный импорт заблокирован."]);
        }

        await using var transaction = await dbContext.Database.BeginTransactionAsync(cancellationToken);
        var batch = new DemandBatch { Name = Path.GetFileNameWithoutExtension(filePath), SourceFile = filePath, FileHashSha256 = hash };
        dbContext.DemandBatches.Add(batch);
        var run = CreateRun("Demand", filePath, hash);
        dbContext.ImportRuns.Add(run);

        await foreach (var row in EnumerateRowsAsync(filePath, new DemandImportProfile(), cancellationToken))
        {
            report.ReadRows++;
            var ips = Get(row, "Ips");
            var quantity = ParseDecimal(Get(row, "Quantity"));
            if (string.IsNullOrWhiteSpace(ips) || quantity is null)
            {
                report.SkippedRows++;
                continue;
            }

            var part = IsDemandLibraryPartCandidate(row, ips)
                ? await FindPartAsync(ips, cancellationToken)
                : null;
            if (part is null && IsDemandLibraryPartCandidate(row, ips))
            {
                var (designation, name) = SplitDesignationAndName(Get(row, "PartName"));
                part = new Part { Ips = ips, Designation = designation, Name = FirstNotEmpty(name, Get(row, "PartName"), "Из потребности"), Source = Path.GetFileName(filePath) };
                dbContext.Parts.Add(part);
            }

            dbContext.DemandItems.Add(new DemandItem
            {
                DemandBatch = batch,
                Part = part,
                Ips = ips,
                SourcePartName = Get(row, "PartName"),
                Project = Get(row, "Project"),
                SerialNumber = Get(row, "SerialNumber"),
                Unit = Get(row, "Unit"),
                Quantity = quantity.Value,
                DemandDate = ParseDate(Get(row, "DemandDate")),
                ProductionSystem = Get(row, "ProductionSystem")
            });
            report.AddedRows++;
            await SaveBatchAsync(report, progress, cancellationToken);
        }

        batch.PeriodFrom = await dbContext.DemandItems.Where(x => x.DemandBatch == batch && x.DemandDate != null).MinAsync(x => x.DemandDate, cancellationToken);
        batch.PeriodTo = await dbContext.DemandItems.Where(x => x.DemandBatch == batch && x.DemandDate != null).MaxAsync(x => x.DemandDate, cancellationToken);
        await FinishRunAsync(run, report, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return report.ToImmutable();
    }

    private async Task<ImportReport> RefreshExistingDemandBatchAsync(
        string filePath,
        DemandBatch batch,
        string hash,
        IProgress<ImportProgress>? progress,
        CancellationToken cancellationToken)
    {
        var report = new MutableImportReport();
        await using var transaction = await dbContext.Database.BeginTransactionAsync(cancellationToken);
        var run = CreateRun("DemandRefresh", filePath, hash);
        dbContext.ImportRuns.Add(run);

        var existingItems = await dbContext.DemandItems
            .Where(x => x.DemandBatchId == batch.Id)
            .OrderBy(x => x.Id)
            .ToListAsync(cancellationToken);

        var updatedItems = new HashSet<long>();
        var sequentialIndex = 0;

        await foreach (var row in EnumerateRowsAsync(filePath, new DemandImportProfile(), cancellationToken))
        {
            report.ReadRows++;
            var ips = Get(row, "Ips");
            var quantity = ParseDecimal(Get(row, "Quantity"));
            if (string.IsNullOrWhiteSpace(ips) || quantity is null)
            {
                report.SkippedRows++;
                continue;
            }

            var demandDate = ParseDate(Get(row, "DemandDate"));
            var item = FindExistingDemandItem(existingItems, updatedItems, sequentialIndex, ips, quantity.Value, demandDate);
            if (item is null)
            {
                report.SkippedRows++;
                continue;
            }

            updatedItems.Add(item.Id);
            sequentialIndex = Math.Max(sequentialIndex + 1, existingItems.IndexOf(item) + 1);

            var part = IsDemandLibraryPartCandidate(row, ips)
                ? await FindPartAsync(ips, cancellationToken)
                : null;
            if (part is null && IsDemandLibraryPartCandidate(row, ips))
            {
                var (designation, name) = SplitDesignationAndName(Get(row, "PartName"));
                part = new Part { Ips = ips, Designation = designation, Name = FirstNotEmpty(name, Get(row, "PartName"), "Из потребности"), Source = Path.GetFileName(filePath) };
                dbContext.Parts.Add(part);
            }

            item.Ips = ips;
            item.Part = part;
            item.SourcePartName = Get(row, "PartName");
            item.Project = Get(row, "Project");
            item.SerialNumber = Get(row, "SerialNumber");
            item.Unit = Get(row, "Unit");
            item.Quantity = quantity.Value;
            item.DemandDate = demandDate;
            item.ProductionSystem = Get(row, "ProductionSystem");
            report.UpdatedRows++;

            await SaveBatchAsync(report, progress, cancellationToken);
        }

        var dates = existingItems
            .Where(x => updatedItems.Contains(x.Id) && x.DemandDate is not null)
            .Select(x => x.DemandDate!.Value)
            .ToList();
        if (dates.Count > 0)
        {
            batch.PeriodFrom = dates.Min();
            batch.PeriodTo = dates.Max();
        }

        batch.SourceFile = filePath;
        batch.ImportedAt = DateTime.UtcNow;
        await FinishRunAsync(run, report, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return report.ToImmutable();
    }

    private static DemandItem? FindExistingDemandItem(
        IReadOnlyList<DemandItem> existingItems,
        ISet<long> updatedItems,
        int sequentialIndex,
        string ips,
        decimal quantity,
        DateTime? demandDate)
    {
        if (sequentialIndex < existingItems.Count &&
            !updatedItems.Contains(existingItems[sequentialIndex].Id) &&
            string.Equals(existingItems[sequentialIndex].Ips, ips, StringComparison.OrdinalIgnoreCase))
        {
            return existingItems[sequentialIndex];
        }

        return existingItems.FirstOrDefault(x =>
            !updatedItems.Contains(x.Id) &&
            string.Equals(x.Ips, ips, StringComparison.OrdinalIgnoreCase) &&
            x.Quantity == quantity &&
            x.DemandDate?.Date == demandDate?.Date);
    }

    public async Task<ImportReport> ImportStockAsync(string filePath, IProgress<ImportProgress>? progress, CancellationToken cancellationToken)
    {
        var report = new MutableImportReport();
        await using var transaction = await dbContext.Database.BeginTransactionAsync(cancellationToken);
        var snapshot = new StockSnapshot { SourceFile = filePath, SnapshotDate = DateTime.Today };
        dbContext.StockSnapshots.Add(snapshot);
        var run = CreateRun("Stock", filePath, await ComputeSha256Async(filePath, cancellationToken));
        dbContext.ImportRuns.Add(run);

        await foreach (var row in EnumerateRowsAsync(filePath, new StockImportProfile(), cancellationToken))
        {
            report.ReadRows++;
            var code = Get(row, "OneCCode");
            var quantity = ParseDecimal(Get(row, "Quantity"));
            if (string.IsNullOrWhiteSpace(code) || quantity is null)
            {
                report.SkippedRows++;
                continue;
            }

            var alias = await FindAliasAsync(code, cancellationToken);
            dbContext.StockItems.Add(new StockItem
            {
                StockSnapshot = snapshot,
                BlankAlias = alias,
                OneCCode = code,
                SourceName = Get(row, "Name"),
                Quantity = quantity.Value,
                Unit = ParseUnit(Get(row, "Unit")),
                Warehouse = Get(row, "Warehouse")
            });
            report.AddedRows++;
            await SaveBatchAsync(report, progress, cancellationToken);
        }

        await FinishRunAsync(run, report, cancellationToken);
        await transaction.CommitAsync(cancellationToken);
        return report.ToImmutable();
    }

    private async IAsyncEnumerable<IReadOnlyDictionary<string, string>> EnumerateRowsAsync(string filePath, IExcelImportProfile profile, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
    {
        if (!CanEnumerateWithEpplus(filePath))
        {
            logger.LogWarning("EPPlus could not open {File}; using OpenXML fallback reader.", filePath);
            await foreach (var row in EnumerateRowsOpenXmlAsync(filePath, profile, cancellationToken))
            {
                yield return row;
            }

            yield break;
        }

        await using var stream = File.OpenRead(filePath);
        using var package = new ExcelPackage(stream);
        foreach (var sheet in package.Workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var headerRow = FindHeaderRow(sheet);
            var headers = ReadRow(sheet, headerRow);
            if (!profile.CanHandle(sheet.Name, headers))
            {
                continue;
            }

            var map = profile.AutoMapColumns(headers);
            for (var row = headerRow + 1; row <= sheet.Dimension.End.Row; row++)
            {
                var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var mapping in map)
                {
                    var column = headers.IndexOf(mapping.Value) + 1;
                    if (column > 0)
                    {
                        values[mapping.Key] = sheet.Cells[row, column].Text.Trim();
                    }
                }

                if (values.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                {
                    yield return values;
                }
            }
        }
    }

    private static bool CanEnumerateWithEpplus(string filePath)
    {
        try
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            _ = package.Workbook.Worksheets.Count;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private async IAsyncEnumerable<IReadOnlyDictionary<string, string>> EnumerateRowsOpenXmlAsync(string filePath, IExcelImportProfile profile, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
    {
        await Task.Yield();
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            yield break;
        }

        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable
            .Elements<SharedStringItem>()
            .Select(x => x.InnerText)
            .ToArray() ?? [];

        foreach (var sheet in workbookPart.Workbook.Sheets?.Elements<Sheet>() ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (sheet.Id?.Value is null)
            {
                continue;
            }

            var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
            var rows = worksheetPart?.Worksheet.Descendants<Row>().ToList();
            if (rows is null || rows.Count == 0)
            {
                continue;
            }

            var headerRowIndex = FindHeaderRow(rows, sharedStrings);
            var headerRow = rows.FirstOrDefault(x => (int?)x.RowIndex?.Value == headerRowIndex);
            if (headerRow is null)
            {
                continue;
            }

            var headers = ReadRow(headerRow, sharedStrings);
            if (!profile.CanHandle(sheet.Name?.Value ?? string.Empty, headers))
            {
                continue;
            }

            var map = profile.AutoMapColumns(headers);
            foreach (var row in rows.Where(x => x.RowIndex?.Value > headerRowIndex))
            {
                var rowValues = ReadRow(row, sharedStrings, headers.Count);
                var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var mapping in map)
                {
                    var column = headers.IndexOf(mapping.Value);
                    if (column >= 0 && column < rowValues.Count)
                    {
                        values[mapping.Key] = rowValues[column].Trim();
                    }
                }

                if (values.Values.Any(v => !string.IsNullOrWhiteSpace(v)))
                {
                    yield return values;
                }
            }
        }
    }

    private async IAsyncEnumerable<IReadOnlyDictionary<string, string>> EnumerateOneCBlankRowsAsync(string filePath, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
    {
        var yielded = false;
        await foreach (var row in EnumerateRowsAsync(filePath, new OneCBlankImportProfile(), cancellationToken))
        {
            if (!string.IsNullOrWhiteSpace(Get(row, "OneCCode")) && !string.IsNullOrWhiteSpace(Get(row, "Name")))
            {
                yielded = true;
                yield return row;
            }
        }

        if (yielded)
        {
            yield break;
        }

        await using var stream = File.OpenRead(filePath);
        using var package = new ExcelPackage(stream);
        foreach (var sheet in package.Workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var rows = sheet.Dimension?.End.Row ?? 0;
            var columns = sheet.Dimension?.End.Column ?? 0;
            for (var row = 1; row <= rows; row++)
            {
                string? code = null;
                string? name = null;
                for (var column = 1; column <= columns; column++)
                {
                    var text = sheet.Cells[row, column].Text.Trim();
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    var match = Regex.Match(text, @"УТ\d{6,}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (!match.Success)
                    {
                        match = Regex.Match(text, @"UT\d{3,}", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    }

                    if (match.Success)
                    {
                        code = match.Value.ToUpperInvariant();
                        name = ExtractNameFromOneCRow(sheet, row, column, text, columns);
                        break;
                    }

                    if (Regex.IsMatch(text, @"^\d{6,8}$", RegexOptions.CultureInvariant))
                    {
                        var ipsName = ExtractNameFromOneCRow(sheet, row, column, text, columns);
                        if (ipsName.Length > 8 && !Regex.IsMatch(ipsName, @"^\d+(?:[\.,]\d+)?$", RegexOptions.CultureInvariant))
                        {
                            code = text;
                            name = ipsName;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrWhiteSpace(code) && !string.IsNullOrWhiteSpace(name))
                {
                    yield return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["OneCCode"] = code,
                        ["Name"] = name
                    };
                }
            }
        }
    }

    private static string ExtractNameFromOneCRow(ExcelWorksheet sheet, int row, int codeColumn, string codeCellText, int columns)
    {
        var inline = Regex.Replace(codeCellText, @"((УТ|UT)\d{3,}|\b\d{6,8}\b)", string.Empty, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Trim();
        if (inline.Length > 8)
        {
            return inline;
        }

        for (var column = codeColumn + 1; column <= Math.Min(columns, codeColumn + 4); column++)
        {
            var candidate = sheet.Cells[row, column].Text.Trim();
            if (candidate.Length > 8 && !Regex.IsMatch(candidate, @"^((УТ|UT)\d{3,}|\d{6,8})$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return candidate;
            }
        }

        for (var column = Math.Max(1, codeColumn - 4); column < codeColumn; column++)
        {
            var candidate = sheet.Cells[row, column].Text.Trim();
            if (candidate.Length > 8 && !Regex.IsMatch(candidate, @"^((УТ|UT)\d{3,}|\d{6,8})$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return candidate;
            }
        }

        return inline;
    }

    private async Task<CanonicalBlank> GetOrCreateCanonicalBlankAsync(BlankNormalizationResult normalized, I012Status status, CancellationToken cancellationToken)
    {
        var blank = dbContext.ChangeTracker.Entries<CanonicalBlank>()
            .Select(x => x.Entity)
            .FirstOrDefault(x => x.CanonicalKey == normalized.CanonicalKey);
        if (blank is not null)
        {
            return blank;
        }

        blank = await dbContext.CanonicalBlanks.FirstOrDefaultAsync(x => x.CanonicalKey == normalized.CanonicalKey, cancellationToken);
        if (blank is not null)
        {
            return blank;
        }

        blank = new CanonicalBlank
        {
            CanonicalName = normalized.NormalizedName,
            CanonicalKey = normalized.CanonicalKey,
            BlankType = normalized.BlankType,
            Material = normalized.Material,
            MaterialGost = normalized.MaterialGost,
            ProfileGost = normalized.ProfileGost,
            DiameterMm = normalized.Diameter,
            WidthMm = normalized.Width,
            HeightMm = normalized.Height,
            ThicknessMm = normalized.Thickness,
            WallThicknessMm = normalized.WallThickness,
            LengthMm = normalized.Length,
            BaseUnit = GetDefaultBaseUnit(normalized.BlankType),
            I012Status = status
        };
        dbContext.CanonicalBlanks.Add(blank);
        return blank;
    }

    private static MeasurementUnit GetDefaultBaseUnit(BlankType blankType) => blankType switch
    {
        BlankType.RoundBar or
        BlankType.SquareBar or
        BlankType.HexBar or
        BlankType.PipeRound or
        BlankType.PipeRectangular or
        BlankType.Angle or
        BlankType.Channel or
        BlankType.IBeam or
        BlankType.BronzeBar => MeasurementUnit.Meter,
        _ => MeasurementUnit.Piece
    };

    private async Task<Part?> FindPartAsync(string ips, CancellationToken cancellationToken)
    {
        var local = dbContext.ChangeTracker.Entries<Part>()
            .Select(x => x.Entity)
            .FirstOrDefault(x => string.Equals(x.Ips, ips, StringComparison.OrdinalIgnoreCase));
        return local ?? await dbContext.Parts.FirstOrDefaultAsync(x => x.Ips == ips, cancellationToken);
    }

    private async Task<BlankAlias?> FindAliasAsync(string oneCCode, CancellationToken cancellationToken)
    {
        var local = dbContext.ChangeTracker.Entries<BlankAlias>()
            .Select(x => x.Entity)
            .FirstOrDefault(x => string.Equals(x.OneCCode, oneCCode, StringComparison.OrdinalIgnoreCase));
        return local ?? await dbContext.BlankAliases
            .Include(x => x.CanonicalBlank)
            .FirstOrDefaultAsync(x => x.OneCCode == oneCCode, cancellationToken);
    }

    private async Task<PartBlankMap?> FindPartBlankMapAsync(Part part, CanonicalBlank blank, CancellationToken cancellationToken)
    {
        var local = dbContext.ChangeTracker.Entries<PartBlankMap>()
            .Select(x => x.Entity)
            .FirstOrDefault(x => x.IsActive && ReferenceEquals(x.Part, part) && ReferenceEquals(x.CanonicalBlank, blank));
        if (local is not null)
        {
            return local;
        }

        if (part.Id == 0 || blank.Id == 0)
        {
            return null;
        }

        return await dbContext.PartBlankMaps
            .FirstOrDefaultAsync(x => x.PartId == part.Id && x.CanonicalBlankId == blank.Id && x.IsActive, cancellationToken);
    }

    private async Task DeactivateOtherPartBlankMapsAsync(Part part, long? keepMapId, CancellationToken cancellationToken)
    {
        var localMaps = dbContext.ChangeTracker.Entries<PartBlankMap>()
            .Select(x => x.Entity)
            .Where(x => x.IsActive && ReferenceEquals(x.Part, part) && (keepMapId is null || x.Id != keepMapId.Value))
            .ToList();
        foreach (var map in localMaps)
        {
            map.IsActive = false;
            map.IsPrimary = false;
            map.UpdatedAt = DateTime.UtcNow;
        }

        if (part.Id == 0)
        {
            return;
        }

        var maps = await dbContext.PartBlankMaps
            .Where(x => x.PartId == part.Id && x.IsActive && (keepMapId == null || x.Id != keepMapId.Value))
            .ToListAsync(cancellationToken);
        foreach (var map in maps)
        {
            map.IsActive = false;
            map.IsPrimary = false;
            map.UpdatedAt = DateTime.UtcNow;
        }
    }

    private static int FindHeaderRow(IReadOnlyList<Row> rows, IReadOnlyList<string> sharedStrings)
    {
        var bestRow = (int)(rows.FirstOrDefault()?.RowIndex?.Value ?? 1);
        var bestScore = -1;
        foreach (var row in rows.Take(30))
        {
            var values = ReadRow(row, sharedStrings);
            var score = values.Count(v => !string.IsNullOrWhiteSpace(v) && v.Any(char.IsLetter));
            if (score > bestScore)
            {
                bestScore = score;
                bestRow = (int)(row.RowIndex?.Value ?? 1);
            }
        }

        return bestRow;
    }

    private static List<string> ReadRow(Row row, IReadOnlyList<string> sharedStrings, int? columns = null)
    {
        var valuesByColumn = new Dictionary<int, string>();
        foreach (var cell in row.Elements<Cell>())
        {
            var column = GetColumnIndex(cell.CellReference?.Value);
            if (column > 0)
            {
                valuesByColumn[column] = ReadCell(cell, sharedStrings);
            }
        }

        var maxColumn = columns ?? (valuesByColumn.Count == 0 ? 1 : valuesByColumn.Keys.Max());
        var values = new List<string>();
        for (var column = 1; column <= maxColumn; column++)
        {
            values.Add(valuesByColumn.TryGetValue(column, out var value) ? value : string.Empty);
        }

        return values;
    }

    private static string ReadCell(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        var value = cell.CellValue?.InnerText ?? cell.InnerText ?? string.Empty;
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sharedStringIndex) &&
            sharedStringIndex >= 0 &&
            sharedStringIndex < sharedStrings.Count)
        {
            return sharedStrings[sharedStringIndex].Trim();
        }

        return value.Trim();
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return 0;
        }

        var column = 0;
        foreach (var character in cellReference.Where(char.IsLetter))
        {
            column = (column * 26) + char.ToUpperInvariant(character) - 'A' + 1;
        }

        return column;
    }

    private static int FindHeaderRow(ExcelWorksheet sheet)
    {
        var maxRows = Math.Min(sheet.Dimension?.End.Row ?? 1, 30);
        var bestRow = 1;
        var bestScore = -1;
        for (var row = 1; row <= maxRows; row++)
        {
            var values = ReadRow(sheet, row);
            var score = values.Count(v => !string.IsNullOrWhiteSpace(v) && v.Any(char.IsLetter));
            if (score > bestScore)
            {
                bestScore = score;
                bestRow = row;
            }
        }

        return bestRow;
    }

    private static List<string> ReadRow(ExcelWorksheet sheet, int row, int? columns = null)
    {
        var maxColumn = columns ?? sheet.Dimension?.End.Column ?? 1;
        var values = new List<string>();
        for (var column = 1; column <= maxColumn; column++)
        {
            values.Add(sheet.Cells[row, column].Text.Trim());
        }

        return values;
    }

    private static async Task<string> ComputeSha256Async(string filePath, CancellationToken cancellationToken)
    {
        await using var stream = File.OpenRead(filePath);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken);
        return Convert.ToHexString(hash);
    }

    private async Task SaveBatchAsync(MutableImportReport report, IProgress<ImportProgress>? progress, CancellationToken cancellationToken)
    {
        if ((report.ReadRows % 1000) == 0)
        {
            await dbContext.SaveChangesAsync(cancellationToken);
            progress?.Report(report.ToProgress());
        }
    }

    private async Task FinishRunAsync(ImportRun run, MutableImportReport report, CancellationToken cancellationToken)
    {
        run.ReadRows = report.ReadRows;
        run.AddedRows = report.AddedRows;
        run.UpdatedRows = report.UpdatedRows;
        run.SkippedRows = report.SkippedRows;
        run.ErrorRows = report.ErrorRows;
        run.Status = report.ErrorRows == 0 ? ImportRunStatus.Completed : ImportRunStatus.CompletedWithErrors;
        run.FinishedAt = DateTime.UtcNow;
        await dbContext.SaveChangesAsync(cancellationToken);
    }

    private static ImportRun CreateRun(string type, string filePath, string hash) => new()
    {
        ImportType = type,
        SourceFile = filePath,
        FileHashSha256 = hash
    };

    private static string Get(IReadOnlyDictionary<string, string> row, string key) => row.TryGetValue(key, out var value) ? value.Trim() : string.Empty;
    private static string FirstNotEmpty(params string?[] values) => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;
    private static bool IsDemandLibraryPartCandidate(IReadOnlyDictionary<string, string> row, string ips)
    {
        var unit = Get(row, "Unit");
        if (!string.IsNullOrWhiteSpace(unit) && !unit.Contains("шт", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var name = Get(row, "PartName");
        if (string.IsNullOrWhiteSpace(ips) || string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        var lower = name.ToLowerInvariant();
        if (lower.Contains("операц", StringComparison.Ordinal) ||
            lower.Contains("сбор", StringComparison.Ordinal) ||
            lower.Contains("узел", StringComparison.Ordinal) ||
            lower.Contains("комплект", StringComparison.Ordinal) ||
            lower.Contains("монтаж", StringComparison.Ordinal))
        {
            return false;
        }

        var (designation, partName) = SplitDesignationAndName(name);
        return !string.IsNullOrWhiteSpace(designation) && !string.IsNullOrWhiteSpace(partName);
    }

    private static (string? Designation, string Name) SplitDesignationAndName(string? value)
    {
        var text = value?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return (null, string.Empty);
        }

        var index = text.IndexOf(' ');
        if (index <= 0)
        {
            return (null, text);
        }

        var first = text[..index].Trim();
        var rest = text[(index + 1)..].Trim();
        var looksLikeDesignation = first.Any(char.IsDigit) &&
            first.Length >= 5 &&
            first.Any(ch => ch is '.' or '-' or '_' or '/');
        return looksLikeDesignation && !string.IsNullOrWhiteSpace(rest)
            ? (first, rest)
            : (null, text);
    }

    private static string BuildBlankName(IReadOnlyDictionary<string, string> row)
    {
        var blankType = Get(row, "BlankTypeName");
        var dimensions = Get(row, "Dimensions");
        var sortament = Get(row, "Sortament");
        var material = Get(row, "Material");
        var diameter = Get(row, "RequiredDiameter");
        var length = Get(row, "RequiredLength");
        return FirstNotEmpty(
            $"{blankType} {dimensions} {material}".Trim(),
            $"{sortament} D{diameter} L{length} {material}".Trim(),
            material);
    }

    private static decimal? ParseDecimal(string value) =>
        decimal.TryParse(value.Replace(",", ".", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture, out var result) ? result : null;

    private static DateTime? ParseDate(string value) =>
        DateTime.TryParse(value, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.AssumeLocal, out var result) ? result : null;

    private static int ParseLeadTimeDays(string value, int fallback = 30)
    {
        var parsed = ParseDecimal(value);
        if (parsed is null)
        {
            return fallback;
        }

        var days = (int)Math.Round(parsed.Value, MidpointRounding.AwayFromZero);
        return Math.Clamp(days, 0, 3650);
    }

    private static MeasurementUnit ParseUnit(string value)
    {
        var normalized = value.ToUpperInvariant();
        if (normalized.Contains("М", StringComparison.Ordinal) && !normalized.Contains("ММ", StringComparison.Ordinal)) return MeasurementUnit.Meter;
        if (normalized.Contains("КГ", StringComparison.Ordinal)) return MeasurementUnit.Kilogram;
        return MeasurementUnit.Piece;
    }

    private sealed class MutableImportReport
    {
        public int ReadRows { get; set; }
        public int AddedRows { get; set; }
        public int UpdatedRows { get; set; }
        public int SkippedRows { get; set; }
        public int ErrorRows { get; set; }
        public List<string> Errors { get; } = [];
        public ImportProgress ToProgress() => new(ReadRows, ReadRows, AddedRows, UpdatedRows, SkippedRows, ErrorRows);
        public ImportReport ToImmutable() => new(ReadRows, AddedRows, UpdatedRows, SkippedRows, ErrorRows, Errors);
    }
}
