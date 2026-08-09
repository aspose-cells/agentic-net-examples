// Title: Assign a line series to the secondary vertical axis (AxisGroup = 2) with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add month‑based data, build a line chart, and move the second series to the secondary Y‑axis by setting its AxisGroup property to 2. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells secondary axis | C# line chart secondary Y axis | AxisGroup property Aspose.Cells | dual axis chart .NET | plot series on secondary vertical axis | Aspose.Cells chart example | Aspose.Cells US | Aspose.Cells Europe
// Common Searches: how to use AxisGroup = 2 in Aspose.Cells | Aspose.Cells C# assign chart series to secondary axis | dual‑axis line chart example Aspose.Cells | set secondary vertical axis for series Aspose.Cells .NET | Aspose.Cells chart secondary axis tutorial
// Developer Intent: Move a specific chart series to the secondary vertical axis by configuring its AxisGroup property.
// Use Cases: Financial dashboards that display revenue on the primary axis and profit margin on a secondary axis. | Scientific reports where temperature is plotted on the primary axis and humidity on a secondary axis. | Business presentations that compare unit sales (primary) with market share percentage (secondary) in a single line chart.
// AI Prompts: Generate C# code that assigns a line series to the secondary Y‑axis using AxisGroup = 2 in Aspose.Cells. | Show an Aspose.Cells example of a dual‑axis line chart with one series on each axis. | Explain the difference between PlotOnSecondAxis and AxisGroup when positioning series in Aspose.Cells charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add month‑based data, build a line chart, and move the second series to the secondary Y‑axis by setting its AxisGroup property to 2. The workbook is saved as an Excel file.
class AssignSeriesToSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["A4"].PutValue("Mar");

        worksheet.Cells["B1"].PutValue("Primary");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Secondary");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["C4"].PutValue(300);

        // Add a line chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series: first uses primary axis, second will use secondary axis
        chart.NSeries.Add("B2:B4", true);   // primary series
        chart.NSeries.Add("C2:C4", true);   // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the second series to the secondary vertical axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Save the workbook
        workbook.Save("LineSeriesSecondaryAxis.xlsx");
    }
}
