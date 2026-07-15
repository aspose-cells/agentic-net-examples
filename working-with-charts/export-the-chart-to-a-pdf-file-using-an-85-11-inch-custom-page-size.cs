using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartExport
{
    public class ExportChartToPdf
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

                // Configure the chart's page setup for a custom 8.5 x 11 inch page
                chart.PageSetup.PaperSize = PaperSizeType.Custom;
                chart.PageSetup.CustomPaperSize(8.5, 11); // width, height in inches

                // Export the chart to PDF with the desired page size and centered alignment
                string outputPath = "ChartOutput.pdf";
                chart.ToPdf(outputPath, 8.5f, 11f,
                            PageLayoutAlignmentType.Center,
                            PageLayoutAlignmentType.Center);

                Console.WriteLine($"Chart exported to PDF: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartToPdf.Run();
        }
    }
}