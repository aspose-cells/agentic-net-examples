using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ResizeRotatedDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data range
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the first series' data labels
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowValue = true;                 // Show the values
                dataLabels.RotationAngle = 45;                // Rotate label text by 45 degrees

                // Disable automatic shape resizing and set custom dimensions
                dataLabels.IsResizeShapeToFitText = false;    // Prevent auto‑fit
                dataLabels.WidthPixel = 80;                   // Desired width in pixels
                dataLabels.HeightPixel = 30;                  // Desired height in pixels

                // Save the workbook
                workbook.Save("ResizedRotatedDataLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeRotatedDataLabels.Run();
        }
    }
}