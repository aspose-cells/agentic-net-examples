// Title: How to export an Excel pivot table with its column chart to a single‑page PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that builds a pivot table from a data range, adds a column chart linked to the pivot data, and saves the worksheet as a PDF with one page per sheet using Aspose.Cells. | Adjust the PDF export options in an Aspose.Cells workbook to enforce landscape orientation, set custom margins, and keep the pivot table and chart layout intact. | Create a reusable C# method that takes a worksheet name and output path, extracts all pivot tables and charts on that sheet, and exports them to a PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# export pivot table and chart to PDF single page | How to save an Excel pivot chart as PDF using Aspose.Cells .NET | PdfSaveOptions one page per sheet for pivot tables Aspose.Cells example | C# generate pivot table, add column chart, and convert to PDF with Aspose | Preserve layout of pivot table and chart when converting to PDF with Aspose.Cells
// Tags: export pivot table to PDF Aspose.Cells | Aspose.Cells PdfSaveOptions one page per sheet | create pivot chart and save as PDF C# | preserve layout Excel to PDF Aspose.Cells | pivot table chart PDF conversion .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

namespace AsposeCellsPivotChartPdfExport
{
    // // Creates a workbook, fills sample data, builds a pivot table with row, column, and data fields, adds a column chart, configures PdfSaveOptions (ExportDocumentStructure, OnePagePerSheet, CalculateFormula), and saves the worksheet containing both the pivot table and chart as a single‑page PDF.
    class Program
    {
        static void Main()
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
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Quantity");

                // Sample data (object array to hold both strings and integers)
                object[,] sample = {
                    {"Fruit", "Apple", 120},
                    {"Fruit", "Banana", 80},
                    {"Fruit", "Apple", 150},
                    {"Vegetable", "Carrot", 200},
                    {"Vegetable", "Broccoli", 90},
                    {"Vegetable", "Carrot", 130}
                };

                for (int i = 0; i < sample.GetLength(0); i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(sample[i, 0]); // Category
                    dataSheet.Cells[i + 1, 1].PutValue(sample[i, 1]); // Product
                    dataSheet.Cells[i + 1, 2].PutValue(sample[i, 2]); // Quantity
                }

                // -------------------------------------------------
                // 2. Create a worksheet to host the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add the pivot table (source range includes header row)
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C7", "A3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields: Category -> Row, Product -> Column, Quantity -> Data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

                // Layout the pivot table in tabular form for better PDF appearance
                pivotTable.ShowInTabularForm();

                // Refresh pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 3. Add a chart that visualizes the pivot table data
                // -------------------------------------------------
                // Place the chart below the pivot table
                int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 20, 0, 35, 8);
                Chart chart = pivotSheet.Charts[chartIndex];

                // Add series using the data sheet range (Quantity values)
                // Categories are taken from the Category column (A) and Product column (B) automatically
                chart.NSeries.Add("Data!C2:C7", true);
                chart.NSeries[0].Name = "Quantity";

                // Set chart title
                chart.Title.Text = "Quantity by Category and Product";

                // -------------------------------------------------
                // 4. Save the workbook (including pivot table and chart) to PDF
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve the document structure (useful for accessibility)
                    ExportDocumentStructure = true,
                    // Fit each worksheet onto a single PDF page to keep layout intact
                    OnePagePerSheet = true,
                    // Optional: calculate formulas before saving
                    CalculateFormula = true
                };

                // Save as PDF; the resulting file contains both the pivot table and its chart
                workbook.Save("PivotTableWithChart.pdf", pdfOptions);

                Console.WriteLine("Pivot table and chart exported to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
