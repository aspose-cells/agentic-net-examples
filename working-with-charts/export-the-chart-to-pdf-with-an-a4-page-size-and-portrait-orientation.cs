using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    class ExportChartToPdfA4
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Chart exported successfully.");
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

            // Set the data source for the chart
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Configure the chart's page setup for A4 size and portrait orientation
            chart.PageSetup.PaperSize = PaperSizeType.PaperA4;
            chart.PageSetup.Orientation = PageOrientationType.Portrait;

            // Define output PDF file path
            string outputPath = "ChartA4Portrait.pdf";

            // Ensure we can write to the output location
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Export the chart to PDF with A4 dimensions (8.27 x 11.69 inches) and centered alignment
            chart.ToPdf(outputPath, 8.27f, 11.69f, PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);
        }
    }
}