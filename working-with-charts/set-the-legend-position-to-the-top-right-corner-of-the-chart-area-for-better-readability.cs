// Title: Aspose.Cells for .NET – Position Chart Legend at Top‑Right Corner (C#)
// Description: Creates a workbook, fills sample data, adds a column chart, and moves the legend to the top‑right corner of the plot area using chart.Legend.Position = LegendPositionType.Corner, then saves the file.
// Keywords: Aspose.Cells legend position | C# chart legend top right | LegendPositionType Corner example | Aspose.Cells chart formatting | Excel chart legend placement .NET
// Common Searches: Aspose.Cells set legend to top right | C# move chart legend corner Aspose | How to place legend in chart corner using Aspose.Cells | Aspose.Cells chart legend positioning options
// Developer Intent: Place a chart legend in the top‑right corner of the plot area.
// Use Cases: Improve readability of column charts by anchoring the legend in the corner. | Standardize legend placement across generated Excel reports. | Create a utility method that applies Corner positioning to any Aspose.Cells chart.
// AI Prompts: Generate C# code that sets a line chart legend to the top‑right corner with Aspose.Cells. | Write a reusable function that accepts a Chart object and applies LegendPositionType.Corner with optional offsets. | Compare LegendPositionType.Corner with other legend position values in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendPositionDemo
{
    // Creates a workbook, fills sample data, adds a column chart, and moves the legend to the top‑right corner of the plot area using chart.Legend.Position = LegendPositionType.Corner, then saves the file.
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

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Position the legend at the top‑right corner of the plot area
            chart.Legend.Position = LegendPositionType.Corner;

            // Save the workbook
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}
