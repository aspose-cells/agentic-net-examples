using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsImageSaveDemo
{
    class Program
    {
        static void Main()
        {
            // Define the output directory
            string outputDir = "output_images";
            // Ensure the directory exists
            Directory.CreateDirectory(outputDir);

            // Generate a timestamped filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string imagePath = Path.Combine(outputDir, $"Chart_{timestamp}.png");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["A4"].PutValue("Cherries");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Save the chart as a PNG image using the timestamped filename
            chart.ToImage(imagePath, ImageType.Png);

            Console.WriteLine($"Chart image saved to: {Path.GetFullPath(imagePath)}");
        }
    }
}