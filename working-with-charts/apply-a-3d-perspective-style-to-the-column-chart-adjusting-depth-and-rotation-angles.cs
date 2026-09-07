// Title: Apply a 3‑D perspective, depth percentage, and rotation angle to a Column3D chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook with sample data and a 3‑D column chart, then set Perspective=40, DepthPercent=150, RotationAngle=30, Elevation=20 using Aspose.Cells in C#. | Update an existing Column3D chart in a .NET workbook to modify its perspective projection, depth percentage, rotation, and elevation via Aspose.Cells properties. | Write C# code that creates a Column3D chart and applies a 3‑D perspective style, adjusting depth and rotation angles through the chart's API.
// Common Searches: Aspose.Cells C# how to change perspective projection of a 3D column chart | set depth percent and rotation angle for Column3D chart using Aspose.Cells .NET | example of adjusting elevation on a 3D column chart with Aspose.Cells | C# code to apply 3D perspective style to Excel chart using Aspose.Cells
// Tags: Aspose.Cells 3D chart perspective | Column3D chart depth percent C# | set rotation angle Aspose.Cells chart | Excel chart elevation Aspose.Cells .NET | configure 3D column chart properties Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills it with sample data, adds a 3‑D column chart, and configures its Perspective, DepthPercent, RotationAngle, and Elevation properties before saving the file as Column3DPerspective.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Q1");
        worksheet.Cells["A3"].PutValue("Q2");
        worksheet.Cells["A4"].PutValue("Q3");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["B4"].PutValue(3000);

        // Add a 3‑D column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply 3‑D perspective style
        chart.RightAngleAxes = false;   // Enable perspective projection
        chart.Perspective = 40;         // Perspective value (0‑100)
        chart.DepthPercent = 150;       // Depth as a percentage of chart width
        chart.RotationAngle = 30;       // Rotation around the Z‑axis (0‑360)
        chart.Elevation = 20;           // Elevation angle (‑90 to 90)

        // Save the workbook with the configured chart
        workbook.Save("Column3DPerspective.xlsx");
    }
}
