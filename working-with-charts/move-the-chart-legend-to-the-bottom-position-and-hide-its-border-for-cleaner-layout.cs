// Title: Aspose.Cells .NET – Move Chart Legend to Bottom and Remove Border
// Description: C# sample that creates a workbook, adds a column chart, positions the legend at the bottom, disables the legend border, and saves the result as ChartWithBottomLegend.xlsx.
// Keywords: Aspose.Cells legend position | chart legend bottom .NET | hide legend border Aspose | Excel chart formatting C# | Aspose.Cells chart customization | global | US
// Common Searches: set chart legend to bottom using Aspose.Cells | remove legend border in Aspose.Cells chart | Aspose.Cells legend positioning example | C# hide chart legend border Aspose | how to format chart legend Aspose.Cells
// Developer Intent: Programmatically place a chart legend at the bottom and hide its border with Aspose.Cells for .NET.
// Use Cases: Design sales dashboards where the legend sits below the chart for clearer space utilization. | Generate printable Excel reports with a clean look by removing legend outlines on multiple charts. | Apply a consistent bottom‑legend style across all charts in a workbook through automated code.
// AI Prompts: Provide C# code using Aspose.Cells to set a chart's legend to the bottom and make the border invisible for a line chart. | Show how to loop through every chart in a workbook and apply bottom legend positioning while hiding borders with Aspose.Cells. | Explain which Aspose.Cells properties control legend placement and border visibility in chart objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendDemo
{
    // C# sample that creates a workbook, adds a column chart, positions the legend at the bottom, disables the legend border, and saves the result as ChartWithBottomLegend.xlsx.
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
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Move the legend to the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Hide the legend border for a cleaner layout
            chart.Legend.Border.IsVisible = false;

            // Save the workbook
            workbook.Save("ChartWithBottomLegend.xlsx");
        }
    }
}
