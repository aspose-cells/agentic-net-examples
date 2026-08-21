// Title: Configure 3‑D Perspective, Depth, Rotation & Elevation for a Column Chart using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample sales data, inserts a 3‑D column chart, disables right‑angle axes, sets perspective to 40, depth to 150 % of chart width, rotation to 30°, elevation to 20°, and saves the file as Column3DPerspective.xlsx.
// Keywords: Aspose.Cells 3D column chart | C# set chart perspective | chart rotation angle Aspose.Cells | depth percent 3D chart | elevation property Aspose.Cells | RightAngleAxes false | Excel 3D chart styling | .NET chart customization
// Common Searches: Aspose.Cells C# 3D column chart perspective | How to change rotation angle of a 3D chart in Aspose.Cells | Set depth percent and elevation for Aspose.Cells chart | Disable right angle axes Aspose.Cells 3D | Apply 3D view to Excel chart with Aspose.Cells
// Developer Intent: Create a 3‑D column chart and fine‑tune its perspective, depth, rotation, and elevation using Aspose.Cells for .NET.
// Use Cases: Generate a sales report workbook that displays a stylized 3‑D column chart with specific perspective and rotation settings for a more engaging visual. | Apply uniform 3‑D view parameters (perspective, depth, rotation, elevation) across multiple charts in a workbook by iterating through worksheet.Charts. | Export an Excel file with a pre‑configured 3‑D column chart so that end users see the exact 3‑D orientation when opening the file in Excel.
// AI Prompts: Show C# code that uses Aspose.Cells to create a 3‑D column chart and set perspective, depth percent, rotation angle, and elevation. | Explain the effect of RightAngleAxes, Perspective, DepthPercent, RotationAngle, and Elevation on a 3‑D chart in Aspose.Cells. | Provide a sample that loops through all charts in a worksheet and applies the same 3‑D perspective and rotation settings with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCells3DPerspectiveDemo
{
    // Creates a workbook, adds sample sales data, inserts a 3‑D column chart, disables right‑angle axes, sets perspective to 40, depth to 150 % of chart width, rotation to 30°, elevation to 20°, and saves the file as Column3DPerspective.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["B3"].PutValue(1500);
            worksheet.Cells["B4"].PutValue(1800);

            // Add a 3‑D column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply 3‑D perspective and related visual settings
            chart.RightAngleAxes = false;   // Enable perspective projection
            chart.Perspective = 40;         // Perspective value (0‑100)
            chart.DepthPercent = 150;       // Depth as a percentage of chart width
            chart.RotationAngle = 30;       // Rotation around the Z‑axis (degrees)
            chart.Elevation = 20;           // Elevation angle (degrees)

            // Save the workbook with the configured chart
            workbook.Save("Column3DPerspective.xlsx");
        }
    }
}
