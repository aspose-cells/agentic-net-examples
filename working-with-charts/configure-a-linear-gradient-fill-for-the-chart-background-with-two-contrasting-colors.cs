// Title: Aspose.Cells for .NET – Apply a Horizontal Two‑Color Linear Gradient to a Chart Plot Area
// Description: Creates a workbook, adds a column chart, and sets the plot‑area background to a horizontal black‑to‑white two‑color gradient using FillFormat and GradientFill in C#.
// Keywords: Aspose.Cells chart gradient | C# linear gradient fill | plot area background Aspose.Cells | two‑color gradient Excel chart | horizontal gradient chart .NET | GradientFill.SetTwoColorGradient | Aspose.Cells chart styling
// Common Searches: Aspose.Cells set chart background gradient C# | horizontal two‑color gradient for Excel chart using Aspose | how to apply linear gradient to chart plot area .NET | gradient fill format Aspose.Cells example
// Developer Intent: Add a horizontal two‑color linear gradient to the chart’s plot‑area background.
// Use Cases: Design eye‑catching Excel dashboards with gradient‑styled charts. | Apply corporate color schemes to chart backgrounds for branded reports. | Create presentation‑ready worksheets where chart aesthetics highlight data trends.
// AI Prompts: Show how to change the gradient direction to vertical and choose custom colors for the chart background with Aspose.Cells. | Provide C# code for a three‑color gradient with specific color stops on a chart’s plot area using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsGradientDemo
{
    // Creates a workbook, adds a column chart, and sets the plot‑area background to a horizontal black‑to‑white two‑color gradient using FillFormat and GradientFill in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // (Optional) Add some sample data for the chart
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

            // Access the fill format of the chart's plot area (background)
            FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;

            // Set the fill type to gradient to enable gradient properties
            plotAreaFill.FillType = FillType.Gradient;

            // Configure a linear (horizontal) two‑color gradient using contrasting colors
            // Here we use Black and White for maximum contrast
            GradientFill gradient = plotAreaFill.GradientFill;
            gradient.SetTwoColorGradient(
                Color.Black,               // First color
                Color.White,               // Second color
                GradientStyleType.Horizontal, // Linear gradient direction
                1);                        // Variant (1‑4)

            // Save the workbook with the configured chart background gradient
            workbook.Save("ChartBackgroundLinearGradient.xlsx");
        }
    }
}
