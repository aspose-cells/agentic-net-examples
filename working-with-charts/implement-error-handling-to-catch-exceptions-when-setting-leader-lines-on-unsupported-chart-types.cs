// Title: C# error handling for unsupported leader lines on Aspose.Cells column charts
// Description: Demonstrates how to wrap leader‑line configuration in a try‑catch block when working with a column chart in Aspose.Cells for .NET. The example catches the exception thrown for unsupported chart types, logs a friendly message, and saves the workbook.
// Keywords: Aspose.Cells leader lines | C# chart error handling | unsupported chart type exception | Aspose.Cells column chart | try catch Aspose chart | .NET Excel chart customization | leader line settings Aspose | Excel chart API error handling
// Common Searches: how to catch leader line exception Aspose.Cells | which Aspose chart types support leader lines | C# Aspose.Cells try catch example | leader lines not supported column chart | Aspose.Cells error handling for chart properties
// Developer Intent: Add defensive code that prevents runtime crashes when applying leader‑line properties to chart types that do not support them.
// Use Cases: Validate chart type before setting HasLeaderLines to avoid exceptions. | Log unsupported‑type errors while continuing processing of other charts. | Automatically switch to a compatible chart (e.g., Pie) when leader lines are required.
// AI Prompts: Generate a C# utility method for Aspose.Cells that checks if a chart type supports leader lines and applies the settings only when valid. | Create code that attempts to enable leader lines on every series in a workbook, returns a success flag, and records failures. | Write a script that iterates through all charts, tries to add leader lines, and produces a summary report of supported vs. unsupported chart types.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Demonstrates how to wrap leader‑line configuration in a try‑catch block when working with a column chart in Aspose.Cells for .NET. The example catches the exception thrown for unsupported chart types, logs a friendly message, and saves the workbook.
class LeaderLinesExample
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

        // Add a column chart (leader lines are not supported for column charts)
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        Series series = chart.NSeries[0];

        try
        {
            // Attempt to enable and customize leader lines
            series.HasLeaderLines = true;
            series.LeaderLines.IsAuto = false;
            series.LeaderLines.Style = LineType.Dot;
            series.LeaderLines.WeightPt = 1.5;
            series.LeaderLines.Color = Color.Blue;
        }
        catch (Exception ex)
        {
            // Catch exceptions thrown for unsupported chart types
            Console.WriteLine("Leader lines are not supported for this chart type: " + ex.Message);
        }

        // Save the workbook
        workbook.Save("LeaderLinesHandled.xlsx");
    }
}
