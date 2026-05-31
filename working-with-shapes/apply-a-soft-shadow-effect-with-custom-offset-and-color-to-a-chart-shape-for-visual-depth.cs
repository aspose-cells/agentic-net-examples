using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
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

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the shape that represents the chart
                Shape chartShape = sheet.Shapes[chartIndex];

                // Obtain the ShadowEffect object
                ShadowEffect shadow = chartShape.ShadowEffect;

                // Set the preset type to Custom to allow custom parameters
                shadow.PresetType = PresetShadowType.Custom;

                // Define custom offset using angle (in degrees) and distance (points)
                shadow.Angle = 45;          // 45 degrees direction
                shadow.Distance = 30;       // 30 points offset from the shape

                // Additional visual parameters
                shadow.Blur = 20;           // Blur radius
                shadow.Size = 1.2;          // Size multiplier
                shadow.Transparency = 0.3;  // 30% transparent

                // Set a custom shadow color (semi‑transparent dark gray)
                CellsColor shadowColor = workbook.CreateCellsColor();
                shadowColor.Color = Color.FromArgb(150, 50, 50, 50); // ARGB with alpha
                shadow.Color = shadowColor;

                // Save the workbook
                string outputPath = "ChartWithSoftShadow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}