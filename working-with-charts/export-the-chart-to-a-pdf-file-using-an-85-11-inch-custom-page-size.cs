// Title: Export a Chart to PDF with a Custom 8.5 × 11 in Page Size Using Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart from sample data, and uses the Chart.ToPdf method to generate a PDF named ChartCustomSize.pdf. The PDF page is set to 8.5 × 11 inches (letter size) and the chart is centered horizontally and vertically with PageLayoutAlignmentType.Center.
// Keywords: Aspose.Cells export chart to PDF | Chart.ToPdf custom page size | 8.5 x 11 inch PDF | center chart on PDF page | Aspose.Cells C# example | .NET chart to PDF | letter size PDF Aspose.Cells | PageLayoutAlignmentType Center | export Excel chart as PDF | custom PDF dimensions Aspose
// Common Searches: Aspose.Cells export chart to PDF with specific size | How to set PDF page width and height in Chart.ToPdf | Center chart on a PDF page using Aspose.Cells | C# code for exporting Excel chart to letter‑size PDF | Custom page layout for chart PDF in Aspose.Cells
// Developer Intent: Generate a PDF file that contains a worksheet chart sized to 8.5 × 11 inches and centered on the page.
// Use Cases: Produce printable reports where each chart fills a standard letter‑size page. | Create PDF invoices or statements that embed a centered sales chart. | Automate batch conversion of multiple Excel charts to uniformly sized PDFs.
// AI Prompts: Show how to modify the code for landscape orientation with an 11 × 8.5 in page. | Give an example that saves the chart as PDF and then merges it with other PDFs using Aspose.PDF. | Explain the effect of different PageLayoutAlignmentType values such as TopLeft, BottomRight, and Center when exporting a chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a column chart from sample data, and uses the Chart.ToPdf method to generate a PDF named ChartCustomSize.pdf. The PDF page is set to 8.5 × 11 inches (letter size) and the chart is centered horizontally and vertically with PageLayoutAlignmentType.Center.
    public class ExportChartToPdfCustomSize
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Chart exported to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Fruits");
            worksheet.Cells["A3"].PutValue("Vegetables");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(50);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B3", true);          // Values
            chart.NSeries.CategoryData = "A2:A3";      // Categories

            // Define output PDF path
            string outputPath = "ChartCustomSize.pdf";

            // Ensure the output directory exists (handle possible null from GetDirectoryName)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            try
            {
                // Export the chart to a PDF file with a custom page size (8.5 × 11 inches)
                // and center the chart horizontally and vertically.
                chart.ToPdf(
                    outputPath,
                    8.5f,   // width in inches (float)
                    11f,    // height in inches (float)
                    PageLayoutAlignmentType.Center,
                    PageLayoutAlignmentType.Center);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to export chart to PDF: {ex.Message}");
                throw;
            }
        }
    }
}
