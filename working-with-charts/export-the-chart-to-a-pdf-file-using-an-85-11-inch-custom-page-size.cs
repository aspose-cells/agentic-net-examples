// Title: Export a column chart from Aspose.Cells to a PDF with an 8.5 × 11 inch custom page size using C#
// AI Prompts: Write C# code that creates a column chart with Aspose.Cells and saves only the chart to a PDF using an 8.5 × 11 inch page and centered alignment. | Show how to export an Aspose.Cells chart to PDF in landscape orientation while specifying custom page dimensions in C#. | Provide a C# example that iterates through all charts in a workbook and exports each one to a separate PDF file with its own custom page size.
// Common Searches: Aspose.Cells C# export chart to PDF with custom 8.5x11 page size | How to set page dimensions when using Chart.ToPdf in Aspose.Cells | Center a chart on a PDF page using Aspose.Cells ToPdf method | Export only a chart, not the worksheet, to PDF with Aspose.Cells | Change PDF orientation for a chart exported with Aspose.Cells C#
// Tags: Aspose.Cells ToPdf chart custom page dimensions | C# export chart as PDF centered alignment | column chart PDF generation Aspose.Cells | chart-only PDF output Aspose.Cells | set PDF orientation for Aspose.Cells chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartToPdf
{
    // The example creates a workbook, adds sample data, inserts a column chart, and then uses Chart.ToPdf to export the chart alone to a PDF file named ChartCustomSize.pdf with an 8.5 × 11 inch page, centered horizontally and vertically.
    public class ExportChartPdf
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Fruits");
                sheet.Cells["A3"].PutValue("Vegetables");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Define output PDF file path
                string outputPath = "ChartCustomSize.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the chart to a PDF with custom page size and centered alignment
                chart.ToPdf(outputPath, 8.5f, 11f,
                    PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting chart to PDF: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartPdf.Run();
        }
    }
}
