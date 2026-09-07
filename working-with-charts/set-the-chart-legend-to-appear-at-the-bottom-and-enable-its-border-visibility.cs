// Title: How to position a chart legend at the bottom and display its border using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample data, generate a column chart, set the legend position to Bottom, and make the legend border visible with a black color using Aspose.Cells in C#. | Modify an existing Aspose.Cells chart to move the legend to the bottom of the chart area and enable the border visibility. | Write C# code that configures a chart's Legend.Position = LegendPositionType.Bottom and sets Legend.Border.IsVisible = true with a specified color.
// Common Searches: Aspose.Cells C# set chart legend to bottom and show border | How to enable legend border visibility in Aspose.Cells chart using .NET | C# Aspose.Cells move chart legend to bottom and change border color | Aspose.Cells chart legend positioning and border styling example | Set Legend.Position = Bottom and Legend.Border.IsVisible in Aspose.Cells
// Tags: chart legend position bottom Aspose.Cells | legend border visibility Aspose.Cells C# | column chart legend styling Aspose.Cells | Aspose.Cells set legend border color | Aspose.Cells chart formatting legend bottom

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLegendDemo
{
    // The example creates a workbook, adds sample data, inserts a column chart, positions the legend at the bottom, makes the legend border visible with a black color, and saves the file as ChartLegendBottomWithBorder.xlsx.
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Enable the legend border visibility (optional: set color for clarity)
            chart.Legend.Border.IsVisible = true;
            chart.Legend.Border.Color = Color.Black; // make the border visible

            // Save the workbook
            workbook.Save("ChartLegendBottomWithBorder.xlsx");
        }
    }
}
