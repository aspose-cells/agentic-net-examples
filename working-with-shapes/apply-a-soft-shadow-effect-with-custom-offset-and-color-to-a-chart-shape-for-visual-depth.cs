// Title: Add a Custom Soft Shadow to a Chart Shape with Offset, Blur, Size, Transparency, and Color using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a column chart, and configures the chart's ChartObject.ShadowEffect with a custom preset (angle, distance, blur, size, transparency) and a semi‑transparent dark gray color, then saves the file.
// Keywords: Aspose.Cells | C# | chart soft shadow | custom shadow effect | shadow angle | shadow distance | shadow blur | shadow size | shadow transparency | shadow color | PresetShadowType.Custom
// Common Searches: Aspose.Cells add soft shadow to chart | set custom shadow properties for chart shape .NET | chart shadow angle and distance Aspose.Cells | change chart shadow color and transparency C# | apply blur to Excel chart shadow using Aspose
// Developer Intent: Apply a custom soft shadow with specific offset, blur, size, transparency, and color to a chart shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance the visual depth of a column chart in a financial report by adding a semi‑transparent dark gray soft shadow. | Standardize chart appearance across multiple workbooks to match corporate branding guidelines. | Programmatically iterate through all charts in a worksheet and apply identical shadow settings for consistent styling.
// AI Prompts: Generate C# code with Aspose.Cells that adds a blue soft shadow to a pie chart, using a 45° angle and 15‑point distance. | Show how to loop through every chart in a worksheet and set a uniform shadow effect with custom blur and transparency values. | Explain how to read the existing ShadowEffect of a chart shape, modify its properties, and save the updated workbook.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a workbook, inserts a column chart, and configures the chart's ChartObject.ShadowEffect with a custom preset (angle, distance, blur, size, transparency) and a semi‑transparent dark gray color, then saves the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure a custom soft shadow for the chart shape
                ShadowEffect shadow = chart.ChartObject.ShadowEffect;
                shadow.PresetType = PresetShadowType.Custom;
                shadow.Angle = 135;          // Direction in degrees
                shadow.Distance = 30;        // Distance in points
                shadow.Blur = 20;            // Blur amount in points
                shadow.Size = 1.2;           // Size multiplier (0‑2.0)
                shadow.Transparency = 0.3;   // 30% transparent

                // Set a semi‑transparent dark gray shadow color
                CellsColor shadowColor = workbook.CreateCellsColor();
                shadowColor.Color = Color.FromArgb(128, 0, 0, 0);
                shadow.Color = shadowColor;

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartSoftShadow.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
