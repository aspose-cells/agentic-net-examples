// Title: Create a Light‑Blue Rectangle Behind a Chart with Aspose.Cells for .NET
// Description: Demonstrates how to add a rectangle shape inside a chart using AddShapeInChartByScale, fill it with solid LightBlue, and set its Z‑order to 0 so the shape appears behind all chart elements in a .NET workbook.
// Keywords: Aspose.Cells add shape in chart | AddShapeInChartByScale C# | chart rectangle fill color | set shape ZOrderPosition Aspose.Cells | light blue rectangle Aspose.Cells | place shape behind chart objects
// Common Searches: how to add a rectangle to a chart with Aspose.Cells | change Z‑order of chart shapes in Aspose.Cells .NET | fill chart shape with specific color using Aspose.Cells | Aspose.Cells AddShapeInChartByScale example | send shape to back of chart Aspose.Cells
// Developer Intent: Insert a rectangle into a chart, apply a LightBlue solid fill, and position it behind all other chart components using Aspose.Cells for .NET.
// Use Cases: Highlight a data region by placing a colored background rectangle behind a column chart. | Add a subtle watermark behind chart series for branding or visual consistency. | Group multiple chart elements visually by using a behind‑the‑scenes rectangle as a container.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle to a chart, fills it with LightBlue, and sends it to the back. | Show how to adjust the Z‑order of a shape inside an Aspose.Cells chart so it appears behind axes and series. | Explain the use of AddShapeInChartByScale for proportionally positioning a shape within a chart in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Demonstrates how to add a rectangle shape inside a chart using AddShapeInChartByScale, fill it with solid LightBlue, and set its Z‑order to 0 so the shape appears behind all chart elements in a .NET workbook.
class Program
{
    static void Main()
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Add a rectangle shape inside the chart using percentage coordinates
        Shape rectangle = chart.Shapes.AddShapeInChartByScale(
            MsoDrawingType.Rectangle,
            PlacementType.Move,
            0.2, // left 20% from chart left
            0.2, // top 20% from chart top
            0.4, // right 40% from chart left
            0.4  // bottom 40% from chart top
        );

        // Fill the rectangle with light blue
        rectangle.Fill.FillType = FillType.Solid;
        rectangle.Fill.SolidFill.Color = Color.LightBlue;

        // Place the rectangle behind other chart objects
        rectangle.ZOrderPosition = 0;

        // Save the workbook
        workbook.Save("RectangleBehindChart.xlsx");
    }
}
