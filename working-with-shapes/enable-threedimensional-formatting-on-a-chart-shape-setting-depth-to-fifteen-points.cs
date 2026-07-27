// Title: Set a 15‑point 3D extrusion depth on a chart shape with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a 3‑D column chart, wraps it in a rectangle shape, and uses the shape's ThreeDFormat.ExtrusionHeight property to apply a 15‑point depth before saving the file as an Excel workbook.
// Keywords: Aspose.Cells C# 3D chart shape | ThreeDFormat.ExtrusionHeight | set chart shape depth Aspose.Cells | 3D extrusion points .NET | Excel chart visual depth Aspose
// Common Searches: how to set extrusion height for a chart shape in Aspose.Cells | apply 3D formatting to Excel chart shape C# | Aspose.Cells set chart shape depth to 15 points | add rectangle shape around chart Aspose.Cells .NET | enable 3D extrusion on chart shape example
// Developer Intent: Apply a 15‑point 3D extrusion depth to a chart’s visual shape.
// Use Cases: Enhance a 3‑D column chart with a defined depth for presentation‑grade reports. | Wrap an existing chart in a shape and control its perceived depth using ThreeDFormat. | Generate Excel files where charts display a consistent 3‑D extrusion for visual hierarchy.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape around a chart and sets ThreeDFormat.ExtrusionHeight to 15 points. | Explain the effect of ThreeDFormat.ExtrusionHeight on chart rendering in Aspose.Cells and show a concise example. | Provide step‑by‑step instructions to enable 3‑D formatting on a chart shape and adjust its depth in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, inserts a 3‑D column chart, wraps it in a rectangle shape, and uses the shape's ThreeDFormat.ExtrusionHeight property to apply a 15‑point depth before saving the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable 3‑D formatting on the chart's visual shape by adding a shape
            // and setting its extrusion height (depth) to 15 points.
            Shape chartShape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 200, 100);
            ThreeDFormat threeDFormat = chartShape.ThreeDFormat;
            threeDFormat.ExtrusionHeight = 15; // depth of 15 points

            // Save the workbook
            workbook.Save("ChartWith3DDepth.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
