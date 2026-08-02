using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartToJpeg
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["A4"].PutValue("Banana");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(80);
            worksheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ----- Branding: set custom background color -----
            // Use an opaque background mode so the color is visible
            chart.ChartArea.BackgroundMode = BackgroundMode.Opaque;
            // Example brand color (e.g., deep teal)
            chart.ChartArea.Area.ForegroundColor = Color.FromArgb(0, 102, 102);
            // Optionally set the plot area background as well
            chart.PlotArea.Area.ForegroundColor = Color.FromArgb(230, 255, 255);
            chart.PlotArea.BackgroundMode = BackgroundMode.Opaque;

            // Configure image options for JPEG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Jpeg,   // Output format
                Quality = 90                  // JPEG quality (0-100)
            };

            // Save the chart as a JPEG file with the branding background
            string outputPath = "BrandedChart.jpeg";
            chart.ToImage(outputPath, imgOptions);

            // Optionally save the workbook for reference
            workbook.Save("WorkbookWithChart.xlsx");

            Console.WriteLine($"Chart saved to JPEG with custom background: {outputPath}");
        }
    }
}