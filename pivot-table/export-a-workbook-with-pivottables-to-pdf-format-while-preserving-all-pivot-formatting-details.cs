using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;   // For PdfSaveOptions

namespace AsposeCellsPivotPdfExport
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Bike");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Bike");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Car");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(1500);

            sheet.Cells["A5"].PutValue("Car");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(1300);

            // 3. Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // 4. Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // 5. Preserve formatting when the pivot is refreshed
            pivotTable.PreserveFormatting = true;

            // 6. Create a style and apply it to the data area of the pivot table
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            style.Font.IsBold = true;
            style.ForegroundColor = System.Drawing.Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the data area using PivotFormats
            pivotTable.PivotFormats.FormatArea(
                PivotFieldType.Data,          // Target area
                0,                            // Subtotal index (0 = no subtotal)
                PivotFieldSubtotalType.None, // No subtotal
                PivotTableSelectionType.DataAndLabel,
                false,                        // Apply to row/column headers
                false,                        // Apply to grand totals
                style);

            // 7. Refresh all pivot tables to ensure data and formatting are up‑to‑date
            workbook.Worksheets.RefreshPivotTables();

            // 8. Set PDF save options to retain document structure (helps keep pivot formatting)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.ExportDocumentStructure = true;

            // 9. Save the workbook as PDF
            workbook.Save("PivotTableExport.pdf", pdfOptions);
        }
    }
}