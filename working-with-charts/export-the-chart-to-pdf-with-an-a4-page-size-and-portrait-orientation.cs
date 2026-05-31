using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExport
{
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
                chart.PageSetup.PaperSize = PaperSizeType.PaperA4;               // A4 paper
                chart.PageSetup.Orientation = PageOrientationType.Portrait;    // Portrait orientation

                // Export the chart to a PDF file
                string outputPath = "Chart_A4_Portrait.pdf";
                chart.ToPdf(outputPath);

                Console.WriteLine("Chart exported to PDF with A4 portrait layout successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during chart export: {ex.Message}");
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