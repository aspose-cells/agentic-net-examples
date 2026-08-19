// Title: Set custom slice colors in a pie chart using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds category/value data, inserts a pie chart, disables automatic color variation, maps "High", "Medium", "Low" to Red, Orange, Green, and applies the colors to each slice before saving.
// Keywords: Aspose.Cells pie chart custom colors | C# set individual slice color | disable color variation Aspose.Cells | map categories to colors chart | chart point foreground color .NET | pie chart slice styling Aspose
// Common Searches: How to color specific slices in an Aspose.Cells pie chart | Assign colors to pie chart points based on category in C# | Turn off automatic color variation for Aspose.Cells charts | Dictionary based slice coloring Aspose.Cells | Custom palette for pie chart using Aspose.Cells .NET
// Developer Intent: Apply predefined colors to each pie‑chart slice according to its category label.
// Use Cases: Risk matrix where high, medium, low risks appear in red, orange, green. | Brand‑compliant sales distribution chart with a fixed color map. | Project status donut showing completed, in‑progress, delayed slices in corporate colors.
// AI Prompts: Generate C# code that sets individual pie slice colors in Aspose.Cells based on a category‑to‑color dictionary. | Show how to disable automatic color variation and assign foreground colors to chart points in a .NET pie chart. | Explain reading category names from worksheet cells and applying matching colors to each slice with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPieCustomColors
{
    // Creates a workbook, adds category/value data, inserts a pie chart, disables automatic color variation, maps "High", "Medium", "Low" to Red, Orange, Green, and applies the colors to each slice before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: categories and their values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("High");
            sheet.Cells["A3"].PutValue("Medium");
            sheet.Cells["A4"].PutValue("Low");

            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(20);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure colors are not automatically varied
            chart.NSeries.IsColorVaried = false;

            // Define custom colors based on importance
            Dictionary<string, Color> importanceColors = new Dictionary<string, Color>()
            {
                { "High", Color.Red },
                { "Medium", Color.Orange },
                { "Low", Color.Green }
            };

            // Apply custom colors to each slice (chart point)
            Series series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                string category = sheet.Cells[i + 2, 0].StringValue; // A2, A3, A4 ...

                if (importanceColors.TryGetValue(category, out Color clr))
                {
                    // Set the foreground color of the slice
                    point.Area.ForegroundColor = clr;
                    point.Area.Formatting = FormattingType.Custom;
                }
            }

            // Save the workbook
            workbook.Save("PieChartCustomColors.xlsx");
        }
    }
}
