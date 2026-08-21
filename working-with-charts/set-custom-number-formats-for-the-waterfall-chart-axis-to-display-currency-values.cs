// Title: Set Currency Number Format on a Waterfall Chart Axis with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates waterfall data, adds a Waterfall chart, and applies the custom number format "$#,##0" to the chart's value‑axis tick labels before saving the file.
// Keywords: Aspose.Cells C# | waterfall chart axis format | currency number format Excel | custom axis tick label format | chart value axis formatting | Aspose.Cells chart example | Excel financial chart automation
// Common Searches: Aspose.Cells set currency format on chart axis | C# waterfall chart number format Aspose | apply $#,##0 to chart value axis in .NET | format waterfall chart axis as currency using Aspose.Cells | Excel chart axis custom number format C#
// Developer Intent: Apply a custom currency number format to the value axis of a Waterfall chart generated with Aspose.Cells.
// Use Cases: Generate financial waterfall charts where axis labels show dollar amounts. | Automate Excel reports that require currency‑formatted chart axes for accounting dashboards. | Create reusable code snippets for styling chart axes in enterprise .NET applications.
// AI Prompts: Show C# code that adds a Waterfall chart with Aspose.Cells and sets the value axis to the "$#,##0" currency format. | Provide an Aspose.Cells example that creates a workbook, fills waterfall data, adds a chart, formats the axis as currency, and saves the file. | Explain how to change number formats for primary and secondary axes of any chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, populates waterfall data, adds a Waterfall chart, and applies the custom number format "$#,##0" to the chart's value‑axis tick labels before saving the file.
class WaterfallChartCurrencyAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the waterfall chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Start");
        sheet.Cells["B2"].PutValue(5000);
        sheet.Cells["A3"].PutValue("Increase");
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["A4"].PutValue("Decrease");
        sheet.Cells["B4"].PutValue(-1500);
        sheet.Cells["A5"].PutValue("End");
        sheet.Cells["B5"].PutValue(5500);

        // Add a Waterfall chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories

        // Apply a custom currency number format to the value axis tick labels
        chart.ValueAxis.TickLabels.NumberFormat = "$#,##0";

        // Save the workbook to a file
        workbook.Save("WaterfallCurrencyAxis.xlsx");
    }
}
