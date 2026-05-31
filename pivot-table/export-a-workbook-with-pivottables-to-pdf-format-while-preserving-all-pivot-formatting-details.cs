using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

namespace AsposeCellsPivotPdfExport
{
    public class ExportPivotToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
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

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Preserve formatting when the pivot table is refreshed
                pivotTable.PreserveFormatting = true;

                // Apply a simple style to the data area (demonstrates formatting preservation)
                Style dataStyle = workbook.CreateStyle();
                dataStyle.Font.IsBold = true;
                dataStyle.ForegroundColor = Color.LightYellow;
                dataStyle.Pattern = BackgroundType.Solid;
                pivotTable.Format(pivotTable.DataBodyRange, dataStyle);

                // Refresh all pivot tables in the workbook to ensure data is up‑to‑date
                workbook.Worksheets.RefreshPivotTables();

                // Set PDF save options to retain document structure (helps keep pivot formatting)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Define output file path
                string outputPath = "PivotTableExport.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}