using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BlankDemandPlanner.Infrastructure.Export;

public sealed class PzmcNeedExportService : IPzmcNeedExportService
{
    public async Task<string> ExportAsync(
        IReadOnlyCollection<PzmcNeedRow> rows,
        string outputDirectory,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(outputDirectory);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var path = Path.Combine(outputDirectory, $"Дефицит ЦМО от {DateTime.Now:yyyy-MM-dd HH-mm}.xlsx");
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Дефицит ЦМО");
        string[] headers =
        [
            "Проект / ПС", "№ станка", "Приоритет", "IPS", "Обозначение", "Наименование",
            "Группа изготовления", "Тип", "Метод изготовления", "Менеджер", "Поставщик",
            "Ед. изм.", "Потребность", "Остаток", "Остаток аналогов", "Плановые поставки",
            "Дефицит", "Срок", "PDF", "Заказ", "Основание", "Операция"
        ];
        for (var column = 0; column < headers.Length; column++)
        {
            sheet.Cells[1, column + 1].Value = headers[column];
        }

        var rowIndex = 2;
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var values = new object?[]
            {
                row.Project, row.SerialNumber, row.Priority, row.Ips, row.Designation, row.Name,
                row.ProductGroup, row.ManufacturingType, row.ManufacturingMethod, row.Manager, row.Supplier,
                row.Unit, row.RequiredQuantity, row.StockQuantity, row.AnalogStockQuantity,
                row.PlannedReceiptQuantity, row.ShortageQuantity, row.NeedDate, row.HasPdf ? "Да" : "Нет",
                row.HasOrder ? "Да" : "Нет", row.Order, row.Operation
            };
            for (var column = 0; column < values.Length; column++)
            {
                sheet.Cells[rowIndex, column + 1].Value = values[column];
            }

            rowIndex++;
        }

        using (var header = sheet.Cells[1, 1, 1, headers.Length])
        {
            header.Style.Font.Bold = true;
            header.Style.Fill.PatternType = ExcelFillStyle.Solid;
            header.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(36, 41, 47));
            header.Style.Font.Color.SetColor(System.Drawing.Color.White);
            header.AutoFilter = true;
        }

        if (rowIndex > 2)
        {
            sheet.Cells[2, 13, rowIndex - 1, 17].Style.Numberformat.Format = "0.####";
            sheet.Cells[2, 18, rowIndex - 1, 18].Style.Numberformat.Format = "dd.mm.yyyy";
        }

        sheet.View.FreezePanes(2, 1);
        sheet.Cells[sheet.Dimension.Address].AutoFitColumns(8, 42);
        await package.SaveAsAsync(new FileInfo(path), cancellationToken);
        return path;
    }
}
