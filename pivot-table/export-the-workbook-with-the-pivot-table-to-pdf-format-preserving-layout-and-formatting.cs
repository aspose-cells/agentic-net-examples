using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ExportPivotTableToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Food");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["A4"].PutValue("Travel");
                sheet.Cells["B4"].PutValue(1500);
                sheet.Cells["A5"].PutValue("Travel");
                sheet.Cells["B5"].PutValue(700);
                sheet.Cells["A6"].PutValue("Supplies");
                sheet.Cells["B6"].PutValue(400);
                sheet.Cells["A7"].PutValue("Supplies");
                sheet.Cells["B7"].PutValue(600);

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D1", "ExpensePivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Layout the pivot table in tabular form for better PDF appearance
                pivotTable.ShowInTabularForm();

                // Refresh all pivot tables to ensure data is up‑to‑date
                workbook.Worksheets.RefreshPivotTables();

                // Set PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,   // retain structure for accessibility
                    OnePagePerSheet = true,           // fit each sheet onto a single page
                    AllColumnsInOnePagePerSheet = true // keep all columns on that page
                };

                // Define output file path
                string outputPath = "PivotTableExport.pdf";

                // Save the workbook (including the pivot table) to PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook with pivot table exported to PDF successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPivotTableToPdf.Run();
        }
    }
}