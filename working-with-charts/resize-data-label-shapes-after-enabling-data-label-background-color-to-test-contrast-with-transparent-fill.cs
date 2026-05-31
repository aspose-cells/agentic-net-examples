using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDataLabelResizeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
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
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure data labels for the first series
                DataLabels dl = chart.NSeries[0].DataLabels;
                dl.ShowValue = true;                                 // Show the values
                dl.Position = LabelPositionType.Center;              // Position inside the point
                dl.BackgroundMode = BackgroundMode.Transparent;     // Transparent fill for contrast testing

                // Disable auto‑fit so we can set a custom size
                dl.IsResizeShapeToFitText = false;

                // Set custom dimensions (in pixels) that are smaller than the default auto‑fit size
                dl.WidthPixel = 60;
                dl.HeightPixel = 30;

                // Optional: change font color to highlight contrast with transparent background
                dl.Font.Color = Color.Black;
                dl.Font.Size = 10;

                // Save the workbook
                string outputPath = "DataLabelResizeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}