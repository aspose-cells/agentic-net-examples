// Title: Make chart legend background transparent while preserving text color with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a column chart, enable its legend, and set each LegendEntry's BackgroundMode to Transparent so the legend has no fill but retains the default text color, then saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | chart legend | transparent legend background | BackgroundMode.Transparent | LegendEntry | remove legend fill | preserve legend text color | Excel chart formatting | global | US
// Common Searches: Aspose.Cells make legend background transparent | remove fill from chart legend C# | set legend entry background mode to transparent Aspose.Cells | keep legend text color while hiding background | Excel chart legend without background using Aspose
// Developer Intent: Set chart legend entries to have no background fill while leaving their existing text color unchanged.
// Use Cases: Designing reports where the legend must blend with a colored worksheet background. | Applying corporate style guidelines that require invisible legend backgrounds in Excel dashboards. | Creating dark‑theme Excel charts where legend text remains readable without a solid fill.
// AI Prompts: Generate C# code with Aspose.Cells that makes all legend entries of a chart transparent while preserving the default font color. | Explain the effect of BackgroundMode.Transparent on LegendEntry objects and any required namespaces. | Provide a step‑by‑step guide to remove legend fill without altering text contrast in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendBackgroundRemoval
{
    // Shows how to create a workbook, add a column chart, enable its legend, and set each LegendEntry's BackgroundMode to Transparent so the legend has no fill but retains the default text color, then saves the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the legend is visible
            chart.ShowLegend = true;

            // Remove background fill from each legend entry while preserving text color
            LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            foreach (LegendEntry entry in legendEntries)
            {
                // Set background to transparent (no fill)
                entry.BackgroundMode = BackgroundMode.Transparent;
                // Text color remains unchanged (default contrast)
            }

            // Save the workbook
            workbook.Save("ChartWithTransparentLegendBackground.xlsx");
        }
    }
}
