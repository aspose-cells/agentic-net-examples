// Title: Validate Chart Presence Before Exporting a Worksheet to PDF with Aspose.Cells for .NET
// Description: This example creates a workbook, adds sample data and a column chart, checks whether the worksheet contains at least one chart, aborts the PDF conversion if none are found, and otherwise saves the file as PDF using PdfSaveOptions with RefreshChartCache enabled.
// Keywords: Aspose.Cells chart validation | PDF export conditional on chart | RefreshChartCache .NET | Worksheet chart count check | skip PDF conversion without chart
// Common Searches: Aspose.Cells verify chart before PDF | C# export worksheet to PDF only if chart exists | how to prevent PDF save when no chart in Excel | check Excel sheet for charts using Aspose.Cells
// Developer Intent: Export a worksheet to PDF only when it contains at least one chart, otherwise skip the conversion.
// Use Cases: Automated financial reports that include charts only when visual data is present. | Batch processing of multiple sheets, exporting PDFs for chart‑enabled worksheets while ignoring empty ones. | Conditional PDF generation in a dashboard that skips sheets lacking sales trend charts.
// AI Prompts: Generate a reusable C# method that validates a worksheet for at least one chart and throws a custom exception if none are found before calling Workbook.Save with PdfSaveOptions. | Show how to log chart‑validation results for each worksheet in a multi‑sheet workbook and continue processing the remaining sheets. | Provide code that iterates through all worksheets, exports only those with charts to separate PDF files, and creates a summary report of skipped sheets.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartValidation
{
    // This example creates a workbook, adds sample data and a column chart, checks whether the worksheet contains at least one chart, aborts the PDF conversion if none are found, and otherwise saves the file as PDF using PdfSaveOptions with RefreshChartCache enabled.
    public class ExportWithChartCheck
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (replace with Workbook("input.xlsx") if a template file is needed)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a chart (comment out this block to test validation when no chart exists)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Validate that the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("The worksheet does not contain any charts. PDF export aborted.");
                    return;
                }

                // Set PDF save options to refresh chart cache
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    RefreshChartCache = true
                };

                // Define output file path
                string outputPath = "ExportedWithChart.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook exported to PDF successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWithChartCheck.Run();
        }
    }
}
