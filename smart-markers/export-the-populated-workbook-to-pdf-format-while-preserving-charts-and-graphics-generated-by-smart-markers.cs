using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Sample data class used as a data source for smart markers
    public class SalesRecord
    {
        public string Category { get; set; } = string.Empty;
        public double Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Populate the worksheet with sample data and smart markers
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");

                // Smart marker rows – these will be replaced by the data source during processing
                sheet.Cells["A2"].PutValue("&=$Sales.Category");
                sheet.Cells["B2"].PutValue("&=$Sales.Amount");
                sheet.Cells["A3"].PutValue("&=$Sales.Category");
                sheet.Cells["B3"].PutValue("&=$Sales.Amount");
                sheet.Cells["A4"].PutValue("&=$Sales.Category");
                sheet.Cells["B4"].PutValue("&=$Sales.Amount");
                sheet.Cells["A5"].PutValue("&=$Sales.Category");
                sheet.Cells["B5"].PutValue("&=$Sales.Amount");
                sheet.Cells.CreateRange("A2:B5").Name = "_CellsSmartMarkers";

                // 4. Add a column chart that will reference the data range (initially empty)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";
                chart.Title.Text = "Sales by Category";

                // 5. Prepare the data source for smart markers
                List<SalesRecord> sales = new List<SalesRecord>
                {
                    new SalesRecord { Category = "Fruits",     Amount = 1200 },
                    new SalesRecord { Category = "Vegetables", Amount = 850 },
                    new SalesRecord { Category = "Beverages",  Amount = 430 },
                    new SalesRecord { Category = "Snacks",     Amount = 670 }
                };

                // 6. Process smart markers (lifecycle rule: use WorkbookDesigner.Process)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Sales", sales);
                designer.Process(); // populates the smart marker range and updates the chart data

                // 7. Configure PDF save options to preserve charts and document structure
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true, // retain document structure (charts, graphics)
                    RefreshChartCache = true,       // ensure chart data is up‑to‑date
                    CalculateFormula = true         // calculate any formulas before saving
                };

                // 8. Save the populated workbook to PDF (lifecycle rule: save)
                string outputPath = "SmartMarkersOutput.pdf";

                // Ensure we can write to the target location
                try
                {
                    workbook.Save(outputPath, pdfOptions);
                    Console.WriteLine($"Workbook with smart markers exported to PDF successfully: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.Error.WriteLine($"Error saving PDF: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // General runtime safety
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}