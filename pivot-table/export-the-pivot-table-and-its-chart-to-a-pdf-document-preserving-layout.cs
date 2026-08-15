// Title: Export Pivot Table and Chart to PDF with Layout Preservation using Aspose.Cells for .NET (C#)
// Description: C# sample that builds a workbook, fills a data sheet, creates a pivot table, adds a column chart based on the pivot range, and saves the result to PDF. PdfSaveOptions are configured with OnePagePerSheet and ExportDocumentStructure to keep each sheet on one page and retain bookmarks.
// Keywords: Aspose.Cells PDF export | pivot table to PDF C# | export pivot chart Aspose | OnePagePerSheet option | ExportDocumentStructure | .NET workbook to PDF | preserve worksheet layout | C# Aspose.Cells example | GitHub Aspose.Cells pivot chart
// Common Searches: how to export a pivot table and chart to PDF using Aspose.Cells | Aspose.Cells preserve layout when saving workbook as PDF | C# export pivot chart to PDF one page per sheet | Aspose.Cells PdfSaveOptions for pivot tables | sample code for pivot table PDF export Aspose
// Developer Intent: Generate a PDF that includes both a pivot table and its associated chart while keeping the original worksheet layout and document structure intact.
// Use Cases: Financial reporting: bundle pivot analysis and visual chart into a single‑page PDF for stakeholder distribution. | Sales dashboards: automate creation of printable PDFs that show summary tables and charts on one page per sheet. | Regulatory submissions: produce PDF files with bookmarks for easy navigation of workbooks containing pivot tables and charts.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table from a data range, add a column chart linked to the pivot, and export the sheet to PDF using OnePagePerSheet and ExportDocumentStructure. | Explain the PdfSaveOptions settings required to keep worksheet layout and bookmarks when exporting a workbook that contains a pivot table and chart. | Troubleshoot why a pivot chart might be missing from the PDF output after using Aspose.Cells in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartPdfExport
{
    // C# sample that builds a workbook, fills a data sheet, creates a pivot table, adds a column chart based on the pivot range, and saves the result to PDF. PdfSaveOptions are configured with OnePagePerSheet and ExportDocumentStructure to keep each sheet on one page and retain bookmarks.
    public class ExportPivotAndChart
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
                string[,] sample = {
                    { "Food", "Fruits", "120" },
                    { "Food", "Vegetables", "80" },
                    { "Beverage", "Tea", "50" },
                    { "Beverage", "Coffee", "70" },
                    { "Food", "Fruits", "150" },
                    { "Beverage", "Tea", "30" }
                };

                for (int i = 0; i < sample.GetLength(0); i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(sample[i, 0]);
                    dataSheet.Cells[i + 1, 1].PutValue(sample[i, 1]);
                    dataSheet.Cells[i + 1, 2].PutValue(Convert.ToDouble(sample[i, 2]));
                }

                // -------------------------------------------------
                // 2. Add a worksheet for the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Create the pivot table using the data range
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C7", "A3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure fields: Category (row), SubCategory (column), Amount (data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Optional layout settings
                pivotTable.ShowInTabularForm();
                pivotTable.PrintDrill = true; // show drill indicators when printed

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();          // correct API to refresh source data
                pivotTable.CalculateData();        // recalculate after refresh

                // -------------------------------------------------
                // 3. Add a chart that visualizes the pivot table data
                // -------------------------------------------------
                // Place the chart on the same pivot sheet, below the pivot table
                int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 15, 0, 30, 8);
                Chart chart = pivotSheet.Charts[chartIndex];

                // Use the pivot table range as the chart data source
                CellArea range = pivotTable.TableRange1;
                string chartRange = CellsHelper.CellIndexToName(range.StartRow, range.StartColumn) + ":" +
                                    CellsHelper.CellIndexToName(range.EndRow, range.EndColumn);

                // Set the data range for the chart (isVertical = true for column chart)
                chart.SetChartDataRange(chartRange, true);
                chart.Title.Text = "Pivot Chart";

                // -------------------------------------------------
                // 4. Save the workbook (including pivot table and chart) to PDF
                // -------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve the layout: each sheet on a single page
                    OnePagePerSheet = true,
                    // Export document structure (bookmarks, etc.)
                    ExportDocumentStructure = true
                };

                string outputPath = "PivotTableAndChart.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
