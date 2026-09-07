// Title: Export an Aspose.Cells column chart to a PDF file with A4 portrait page size using C#
// AI Prompts: Write C# code that builds a workbook, adds a column chart, sets the chart's PageSetup to PaperSizeType.PaperA4 and PageOrientationType.Portrait, and saves the chart as a PDF with Aspose.Cells. | Demonstrate how to configure a chart's page dimensions and orientation before calling ToPdf in Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells export chart to pdf a4 portrait orientation | c# set chart page size to A4 before PDF export using Aspose.Cells | how to change chart orientation to portrait in Aspose.Cells and generate PDF | Aspose.Cells chart ToPdf with specific paper size and orientation C#
// Tags: Aspose.Cells chart export to PDF | chart PageSetup A4 portrait | PaperSizeType.PaperA4 chart configuration | PageOrientationType.Portrait Aspose.Cells | column chart PDF generation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExport
{
    // The sample creates a workbook, inserts sample data, adds a column chart, configures the chart's PageSetup to A4 portrait, and exports the chart to a PDF file named Chart_A4_Portrait.pdf using Aspose.Cells for .NET.
    public class ExportChartToPdfA4Portrait
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
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

                // Configure page setup for A4 size and portrait orientation
                chart.PageSetup.PaperSize = PaperSizeType.PaperA4;
                chart.PageSetup.Orientation = PageOrientationType.Portrait;

                // Export the chart to a PDF file
                chart.ToPdf("Chart_A4_Portrait.pdf");

                Console.WriteLine("Chart exported to PDF with A4 portrait layout successfully.");
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
            ExportChartToPdfA4Portrait.Run();
        }
    }
}
