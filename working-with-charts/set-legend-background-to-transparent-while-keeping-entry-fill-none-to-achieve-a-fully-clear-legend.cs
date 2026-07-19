// Title: Make Chart Legend Fully Transparent with Aspose.Cells for .NET
// Description: Shows how to generate a workbook, add a column chart, and hide the legend completely by setting `chart.Legend.BackgroundMode` to `BackgroundMode.Transparent` and applying the same setting plus `IsTextNoFill = true` to each `LegendEntry` using Aspose.Cells C# API.
// Keywords: Aspose.Cells | C# | .NET | chart legend transparent | BackgroundMode.Transparent | LegendEntry no fill | Excel chart customization | hide legend Aspose.Cells | transparent legend entries | Excel export C#
// Common Searches: Aspose.Cells set chart legend transparent | C# hide legend background Aspose.Cells | transparent legend entries Aspose.Cells | remove legend fill from Excel chart .NET | make legend invisible without deleting Aspose.Cells
// Developer Intent: Visually hide the chart legend while keeping the legend object intact in the workbook.
// Use Cases: Create clean reports where the legend would obscure data points. | Overlay charts on custom background images and suppress the legend. | Generate printable charts that omit the legend but retain it for later editing. | Prepare dashboards where legends are managed programmatically rather than displayed.
// AI Prompts: Write C# code with Aspose.Cells to set a chart legend and all its entries to transparent while preserving the legend object. | Explain the purpose of `BackgroundMode.Transparent` and `IsTextNoFill` for legend customization in Aspose.Cells. | Show how to verify that a chart legend is fully transparent after saving the workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Shows how to generate a workbook, add a column chart, and hide the legend completely by setting `chart.Legend.BackgroundMode` to `BackgroundMode.Transparent` and applying the same setting plus `IsTextNoFill = true` to each `LegendEntry` using Aspose.Cells C# API.
class TransparentLegendDemo
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
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(50);
        sheet.Cells["B4"].PutValue(70);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Make the whole legend background transparent
        chart.Legend.BackgroundMode = BackgroundMode.Transparent;

        // Ensure each legend entry has no fill (transparent background and no text fill)
        LegendEntryCollection entries = chart.Legend.LegendEntries;
        foreach (LegendEntry entry in entries)
        {
            // Transparent background for the entry
            entry.BackgroundMode = BackgroundMode.Transparent;
            // No fill for the text inside the entry
            entry.IsTextNoFill = true;
        }

        // Save the workbook
        workbook.Save("TransparentLegendDemo.xlsx");
    }
}
