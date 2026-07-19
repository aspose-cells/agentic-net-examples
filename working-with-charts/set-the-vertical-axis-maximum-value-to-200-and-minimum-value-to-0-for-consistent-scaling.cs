// Title: Aspose.Cells for .NET – Set Fixed 0‑200 Range on a Chart’s Vertical Axis (C#)
// Description: This C# example creates a workbook, adds sample data, inserts a column chart, turns off automatic scaling, sets the value axis minimum to 0 and maximum to 200, and saves the file as ChartWithFixedAxis.xlsx.
// Keywords: Aspose.Cells chart axis limits | C# set chart Y‑axis min max | disable automatic scaling Aspose.Cells | fixed vertical axis range .NET | column chart axis configuration
// Common Searches: Aspose.Cells set chart Y axis minimum | how to fix chart vertical axis range in .NET | Aspose.Cells disable auto scaling for axis | C# chart axis min max Aspose.Cells example | set fixed axis values column chart Aspose
// Developer Intent: Apply explicit minimum and maximum values to a chart’s value axis instead of using auto‑scaling.
// Use Cases: Standardize Y‑axis scaling across quarterly financial charts. | Build dashboards where the data range is predetermined and must remain constant. | Prevent visual jumps when future data points exceed current values.
// AI Prompts: Write C# code with Aspose.Cells that forces a column chart’s vertical axis to run from 0 to 200. | Explain step‑by‑step how to disable automatic axis scaling and set fixed limits on a chart in Aspose.Cells for .NET. | Generate a tutorial for creating a chart with a static Y‑axis range using Aspose.Cells C# API.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds sample data, inserts a column chart, turns off automatic scaling, sets the value axis minimum to 0 and maximum to 200, and saves the file as ChartWithFixedAxis.xlsx.
class SetAxisScale
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
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(120);
        sheet.Cells["B4"].PutValue(180);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Disable automatic scaling and set fixed min/max values for the vertical (value) axis
        chart.ValueAxis.IsAutomaticMinValue = false;
        chart.ValueAxis.MinValue = 0;
        chart.ValueAxis.IsAutomaticMaxValue = false;
        chart.ValueAxis.MaxValue = 200;

        // Save the workbook to a file
        workbook.Save("ChartWithFixedAxis.xlsx");
    }
}
