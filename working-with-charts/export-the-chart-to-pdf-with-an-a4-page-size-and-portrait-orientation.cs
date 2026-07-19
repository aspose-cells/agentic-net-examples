// Title: Export Aspose.Cells Chart to PDF (A4 Portrait) – C# Example
// Description: Demonstrates how to create a workbook, add a column chart, set the chart's PageSetup to A4 paper size with portrait orientation, and export the chart to a PDF file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart export PDF | C# chart to PDF A4 portrait | Aspose.Cells PageSetup PaperA4 | chart ToPdf C# example | export Excel chart as PDF | .NET Aspose.Cells PDF settings
// Common Searches: Aspose.Cells export chart to PDF A4 portrait C# | set chart paper size A4 before PDF export Aspose.Cells | C# Aspose.Cells chart ToPdf orientation | how to export Excel chart as PDF with specific page size | Aspose.Cells chart PDF page setup example
// Developer Intent: Generate a PDF file of a chart with A4 paper size and portrait orientation using Aspose.Cells in C#.
// Use Cases: Create printable sales or financial charts that match standard A4 report layouts. | Automate batch conversion of workbook charts to PDF for distribution to stakeholders. | Produce consistent PDF assets for marketing brochures that require portrait orientation.
// AI Prompts: Write C# code with Aspose.Cells to export a chart to a PDF file using A4 portrait page setup. | Show how to change the example to export the chart in landscape orientation on Letter paper. | Explain how to loop through all charts in a workbook and save each as an A4 portrait PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExport
{
    // Demonstrates how to create a workbook, add a column chart, set the chart's PageSetup to A4 paper size with portrait orientation, and export the chart to a PDF file using Aspose.Cells for .NET.
    public class ExportChartToPdfA4Portrait
    {
        public static void Run()
        {
            try
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

                // Set the data source for the chart
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Configure the chart's page setup for A4 size and portrait orientation
                chart.PageSetup.PaperSize = PaperSizeType.PaperA4;          // A4 paper size
                chart.PageSetup.Orientation = PageOrientationType.Portrait; // Portrait orientation

                // Export the chart to a PDF file
                string outputPath = "Chart_A4_Portrait.pdf";
                chart.ToPdf(outputPath);

                Console.WriteLine($"Chart exported to PDF with A4 portrait layout successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while exporting the chart: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartToPdfA4Portrait.Run();
        }
    }
}
