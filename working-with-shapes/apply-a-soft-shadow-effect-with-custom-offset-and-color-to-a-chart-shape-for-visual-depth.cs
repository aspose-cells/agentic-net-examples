// Title: Add a Custom Soft Shadow to a Chart Shape with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts a column chart, retrieves its ChartShape, and configures the ShadowEffect (custom preset, angle, distance, blur, size, transparency) with a semi‑transparent dark color before saving the file.
// Keywords: Aspose.Cells | C# chart shadow | ChartShape ShadowEffect | custom soft shadow | Excel chart visual depth | set shadow color Aspose | shadow angle distance blur | Aspose.Cells example | chart formatting .NET
// Common Searches: Aspose.Cells add shadow to chart | C# set chart shape shadow | custom shadow effect Excel chart Aspose | change chart shadow color Aspose.Cells | adjust shadow blur distance Aspose
// Developer Intent: Programmatically apply a custom soft shadow with specific offset, blur, size, transparency, and color to a chart shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Give a column chart a subtle depth effect in automatically generated financial reports. | Standardize chart appearance across a corporate workbook by applying a consistent soft shadow to all chart shapes. | Dynamically match shadow parameters to a brand palette when creating presentation‑style Excel files.
// AI Prompts: Generate C# code with Aspose.Cells that adds a soft shadow (angle 120°, distance 25 pt, blur 15, size 1.1, 40% transparency) to a pie chart shape. | Explain how to change the shadow color of a ChartShape to an opaque blue using the Aspose.Cells API. | Provide step‑by‑step instructions to retrieve a chart shape from a worksheet and set its ShadowEffect properties, including preset type, angle, distance, blur, size, transparency, and color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsShadowDemo
{
    // Creates a workbook, inserts a column chart, retrieves its ChartShape, and configures the ShadowEffect (custom preset, angle, distance, blur, size, transparency) with a semi‑transparent dark color before saving the file.
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Retrieve the shape that hosts the chart
                Shape shape = sheet.Shapes[chartIndex];
                ChartShape chartShape = shape as ChartShape;
                if (chartShape != null)
                {
                    // Apply a custom soft shadow to the chart shape
                    ShadowEffect shadow = chartShape.ShadowEffect;
                    shadow.PresetType = PresetShadowType.Custom;
                    shadow.Angle = 135;          // Diagonal bottom‑right direction
                    shadow.Distance = 30;        // Offset distance in points
                    shadow.Blur = 20;            // Blur radius
                    shadow.Size = 1.2;           // Size multiplier
                    shadow.Transparency = 0.3;   // 30% transparent

                    // Define a semi‑transparent dark color for the shadow
                    CellsColor shadowColor = workbook.CreateCellsColor();
                    shadowColor.Color = Color.FromArgb(128, 0, 0, 0); // 50% opaque black
                    shadow.Color = shadowColor;
                }

                // Save the workbook with the chart that now has a soft shadow
                workbook.Save("ChartWithSoftShadow.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
