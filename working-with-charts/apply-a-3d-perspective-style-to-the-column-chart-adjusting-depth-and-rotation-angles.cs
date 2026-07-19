// Title: Add 3‑D Perspective, Depth, Rotation, and Elevation to a Column Chart with Aspose.Cells for .NET
// Description: Creates a new workbook, fills cells A1:B4 with sample sales data, inserts a 3‑D column chart, disables right‑angle axes, and configures Perspective (40), RotationAngle (30), DepthPercent (150) and Elevation (20) before saving as Column3DPerspective.xlsx.
// Keywords: Aspose.Cells 3D column chart | chart perspective .NET | DepthPercent property | RotationAngle setting | Elevation angle Aspose.Cells | RightAngleAxes false | C# Excel chart example | 3D chart customization
// Common Searches: how to set perspective on a 3d column chart using Aspose.Cells | Aspose.Cells DepthPercent and RotationAngle example | change elevation angle of 3D chart in C# | disable right angle axes Aspose.Cells chart | 3D column chart tutorial Aspose.Cells .NET
// Developer Intent: Apply a 3‑D perspective style to a column chart and fine‑tune its depth, rotation, and elevation using Aspose.Cells for .NET.
// Use Cases: Generate sales reports with a realistic 3‑D column chart for presentations. | Programmatically match corporate visual standards by adjusting chart depth and angle. | Create Excel dashboards where 3‑D perspective highlights data trends.
// AI Prompts: Show how to enable perspective and set depth, rotation, and elevation on a 3‑D column chart with Aspose.Cells for .NET. | Provide a C# example that creates a 3‑D column chart and configures RightAngleAxes, Perspective, RotationAngle, DepthPercent, and Elevation. | Explain the valid value ranges for Perspective, DepthPercent, RotationAngle, and Elevation in Aspose.Cells chart settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCells3DPerspectiveDemo
{
    // Creates a new workbook, fills cells A1:B4 with sample sales data, inserts a 3‑D column chart, disables right‑angle axes, and configures Perspective (40), RotationAngle (30), DepthPercent (150) and Elevation (20) before saving as Column3DPerspective.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(3000);

            // Add a 3‑D column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply 3‑D perspective and related settings
            chart.RightAngleAxes = false;   // Enable perspective projection
            chart.Perspective = 40;         // Perspective (0‑100)
            chart.RotationAngle = 30;       // Rotation around Z‑axis (0‑360)
            chart.DepthPercent = 150;       // Depth as % of chart width (20‑2000)
            chart.Elevation = 20;           // Elevation angle (-90‑90)

            // Save the workbook
            workbook.Save("Column3DPerspective.xlsx");
        }
    }
}
