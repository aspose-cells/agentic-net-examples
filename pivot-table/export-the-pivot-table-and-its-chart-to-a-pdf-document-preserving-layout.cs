using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotChartPdfExport
{
    public class ExportPivotAndChartToPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Populate source data for the pivot table
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header row
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("SubCategory");
                dataSheet.Cells["C1"].PutValue("Amount");

                // Sample data
                string[,] sample = new string[,]
                {
                    { "Food", "Fruits", "120" },
                    { "Food", "Vegetables", "80" },
                    { "Beverage", "Tea", "50" },
                    { "Beverage", "Coffee", "70" },
                    { "Food", "Fruits", "150" },
                    { "Beverage", "Tea", "30" }
                };

                for (int i = 0; i < sample.GetLength(0); i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(sample[i, 0]); // Category
                    dataSheet.Cells[i + 1, 1].PutValue(sample[i, 1]); // SubCategory
                    dataSheet.Cells[i + 1, 2].PutValue(Convert.ToDouble(sample[i, 2])); // Amount
                }

                // -------------------------------------------------
                // 2. Add a worksheet that will contain the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Define the source range for the pivot table
                string sourceRange = "Data!A1:C7";

                // Add the pivot table at cell A1 of the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Layout the pivot table in tabular form for better appearance
                pivotTable.ShowInTabularForm();

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 3. Add a chart that visualizes the pivot table data
                // -------------------------------------------------
                // Place the chart on the same pivot sheet, below the pivot table
                int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 15, 0, 30, 8);
                Chart chart = pivotSheet.Charts[chartIndex];

                // Set the chart title
                chart.Title.Text = "Sales by Category and SubCategory";

                // Use the pivot table's data range as the source for the chart
                string chartDataRange = pivotSheet.Cells.MaxDisplayRange.RefersTo;
                chart.SetChartDataRange(chartDataRange, true);

                // -------------------------------------------------
                // 4. Prepare PDF save options to preserve layout
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true,
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true
                };

                // -------------------------------------------------
                // 5. Save the workbook (including pivot table and chart) to PDF
                // -------------------------------------------------
                string outputPdfPath = "PivotTableWithChart.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPdfPath, pdfOptions);
                Console.WriteLine($"Pivot table and chart exported successfully to '{outputPdfPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPivotAndChartToPdf.Run();
        }
    }
}