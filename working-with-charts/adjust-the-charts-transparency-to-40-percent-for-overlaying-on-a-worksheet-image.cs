using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartTransparencyDemo
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

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Insert a picture under the chart if the file exists
                const string imagePath = "image.png";
                if (File.Exists(imagePath))
                {
                    // topRow, leftColumn, fileName
                    sheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file \"{imagePath}\" not found. Skipping picture insertion.");
                }

                // Add a column chart that will overlay the picture
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Make the chart background transparent
                chart.ChartArea.BackgroundMode = BackgroundMode.Transparent;

                // Set the chart area transparency to 40%
                chart.ChartArea.Area.Transparency = 0.4;

                // Save the workbook
                workbook.Save("ChartWithTransparency.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}