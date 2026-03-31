using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartToImageDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file containing the chart
            string sourcePath = "input.xlsx";

            // Path for the output image file (PNG format preserves visual fidelity)
            string outputImagePath = "chart_image.png";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart.
            // If no chart exists, create a simple column chart for demonstration.
            if (worksheet.Charts.Count == 0)
            {
                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["B4"].PutValue(150);

                // Add a column chart covering the data range
                int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
            }

            // Retrieve the first chart in the worksheet
            Chart targetChart = worksheet.Charts[0];

            // Export the chart to an image file.
            // The overload with (string, ImageType) determines the image format.
            targetChart.ToImage(outputImagePath, ImageType.Png);

            Console.WriteLine($"Chart successfully exported to image: {outputImagePath}");
        }
    }
}