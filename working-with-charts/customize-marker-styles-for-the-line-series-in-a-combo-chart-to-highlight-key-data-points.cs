// Title: Aspose.Cells C# – Customize Line Series Markers in a Combo Chart and Highlight a Specific Point
// Description: Creates a workbook with monthly sales and profit data, adds a combo chart (columns for sales, line for profit), applies a circular marker to all profit points, and emphasizes the March point with a larger square‑plus marker in red and orange using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart marker customization | C# combo chart line series marker | highlight specific chart point Aspose | custom marker size color shape | Aspose.Cells line series styling
// Common Searches: Aspose.Cells change line series marker style | C# highlight a point in combo chart Aspose | customize chart markers in Aspose.Cells .NET | set marker shape and color for chart points | Aspose.Cells combo chart marker example
// Developer Intent: Apply custom marker shapes, sizes, and colors to a line series in a combo chart and draw attention to a chosen data point.
// Use Cases: Visually differentiate profit values in a sales‑profit combo chart. | Mark the month with the highest profit using a distinct, larger marker. | Standardize marker appearance across multiple line series in automated reports.
// AI Prompts: Generate C# code with Aspose.Cells that sets a circular marker (size 10, dark blue border, light yellow fill) for all points in a line series of a combo chart. | Show how to change the third point of a line series to a red square‑plus marker with an orange background using Aspose.Cells. | Explain how to access ChartPoint objects and modify their Marker properties programmatically in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ComboChartMarkerDemo
{
    // Creates a workbook with monthly sales and profit data, adds a combo chart (columns for sales, line for profit), applies a circular marker to all profit points, and emphasizes the March point with a larger square‑plus marker in red and orange using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column B – Sales (displayed as columns)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Column C – Profit (displayed as line)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);
            sheet.Cells["C5"].PutValue(70);

            // Add a combo chart (initially a column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (Sales) – stays as column
            chart.NSeries.Add("B2:B5", true);
            // Add the second series (Profit) – will be changed to line
            chart.NSeries.Add("C2:C5", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Change the second series to a line type (creates a combo chart)
            chart.NSeries[1].Type = ChartType.Line;

            // ----- Customize marker style for the line series (Profit) -----
            Series lineSeries = chart.NSeries[1];

            // General marker appearance for all points in the line series
            lineSeries.Marker.MarkerStyle = ChartMarkerType.Circle;
            lineSeries.Marker.MarkerSize = 10;                     // size in points
            lineSeries.Marker.ForegroundColor = Color.DarkBlue;   // border color
            lineSeries.Marker.BackgroundColor = Color.LightYellow;

            // Highlight a specific key data point (e.g., March profit)
            int keyPointIndex = 2; // zero‑based index, March is the third point
            ChartPoint keyPoint = lineSeries.Points[keyPointIndex];
            keyPoint.Marker.MarkerStyle = ChartMarkerType.SquarePlus;
            keyPoint.Marker.MarkerSize = 14;
            keyPoint.Marker.ForegroundColor = Color.Red;
            keyPoint.Marker.BackgroundColor = Color.Orange;

            // Save the workbook
            workbook.Save("ComboChartWithCustomMarkers.xlsx");
        }
    }
}
