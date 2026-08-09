// Title: How to Fix the Y‑Axis Range (0‑200) for a Column Chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, disables automatic scaling, and sets the value axis minimum to 0 and maximum to 200 using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# chart axis | set chart Y axis min max | fixed vertical axis Aspose.Cells | column chart scaling Aspose | disable automatic axis scaling | value axis MinValue MaxValue | Excel chart programmatic axis limits
// Common Searches: Aspose.Cells set chart Y axis minimum | Aspose.Cells set chart Y axis maximum | C# Aspose.Cells custom axis range | prevent auto scaling of chart axis in Aspose.Cells | define fixed vertical axis for Excel chart using Aspose
// Developer Intent: Programmatically define a fixed vertical axis range of 0 to 200 for a chart in an Excel file using Aspose.Cells for .NET.
// Use Cases: Standardize Y‑axis across quarterly sales charts in financial reports. | Create KPI dashboards where the expected data range is known and should not auto‑scale. | Generate printable charts with consistent axis limits for presentations. | Prepare templates that must accommodate future data without changing the scale.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart and set its value axis MinValue = 0 and MaxValue = 200. | Show how to apply the same fixed Y‑axis range to every chart in a workbook using Aspose.Cells for .NET. | Explain the purpose of IsAutomaticMinValue, IsAutomaticMaxValue, MinValue, and MaxValue properties for chart axes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, disables automatic scaling, and sets the value axis minimum to 0 and maximum to 200 using Aspose.Cells in C#.
class SetVerticalAxisScale
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(120);
        sheet.Cells["B4"].PutValue(180);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure the vertical (value) axis scaling
        Axis valueAxis = chart.ValueAxis;
        valueAxis.IsAutomaticMinValue = false; // disable automatic min
        valueAxis.MinValue = 0;                // set minimum to 0
        valueAxis.IsAutomaticMaxValue = false; // disable automatic max
        valueAxis.MaxValue = 200;              // set maximum to 200

        // Save the workbook
        workbook.Save("ChartWithFixedAxis.xlsx");
    }
}
