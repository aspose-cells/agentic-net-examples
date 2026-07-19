// Title: Aspose.Cells C# – Create Column Chart, Set Legend Bottom, and Raise It with YRatioToChart
// Description: Demonstrates how to build a column chart in a new workbook, place the legend at the bottom, then lift the legend by 5 % of the chart height using the Legend.YRatioToChart property, and finally save the file as an Excel workbook.
// Keywords: Aspose.Cells legend bottom | YRatioToChart offset | adjust chart legend position C# | column chart legend shift | Aspose.Cells chart customization | Excel legend placement .NET | global Aspose.Cells examples
// Common Searches: Aspose.Cells move legend up after setting bottom | YRatioToChart legend example C# | how to offset chart legend in Aspose.Cells | set legend position bottom then raise it Aspose.Cells | adjust legend vertical position programmatically
// Developer Intent: Place a chart legend at the bottom of a column chart and then raise it by a fixed percentage of the chart height.
// Use Cases: Fine‑tune legend placement in sales or KPI column charts to avoid overlap with axis labels. | Standardize legend positioning across automated Excel reports generated with Aspose.Cells. | Create multi‑series charts where a bottom legend needs a slight upward offset for visual balance.
// AI Prompts: Generate C# code with Aspose.Cells to create a line chart, set the legend to the right, and shift it left by 0.07 using XRatioToChart. | Explain the YRatioToChart property for legend positioning and how to compute a percentage‑based offset. | Show how to read the current Legend.YRatioToChart value and adjust it dynamically based on worksheet data.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to build a column chart in a new workbook, place the legend at the bottom, then lift the legend by 5 % of the chart height using the Legend.YRatioToChart property, and finally save the file as an Excel workbook.
class ChartLegendShiftExample
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

        // Position the legend at the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;

        // Shift the legend upward by a fixed offset (5% of the chart height)
        // YRatioToChart is a fraction (0‑1) representing the vertical position relative to the chart area.
        // Decreasing the value moves the legend upward.
        chart.Legend.YRatioToChart -= 0.05; // Adjust as needed for the desired offset

        // Save the workbook with the configured chart
        workbook.Save("ChartWithLegendShifted.xlsx");
    }
}
