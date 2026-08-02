// Title: Aspose.Cells for .NET: Create a Column Chart and Raise Bottom Legend with YRatioToChart
// Description: Demonstrates how to generate a column chart in a new workbook, set the legend to the bottom, enable overlay, and shift the legend upward using the YRatioToChart property before saving the file as an Excel workbook.
// Keywords: Aspose.Cells chart legend position | C# Aspose.Cells legend offset | YRatioToChart example | LegendPositionType.Bottom | chart legend overlay Aspose.Cells | adjust legend vertical placement | Aspose.Cells .NET chart customization | Excel chart legend shift | Aspose.Cells legend Y ratio | move chart legend upward
// Common Searches: Aspose.Cells move legend upward C# | set chart legend bottom with offset Aspose.Cells | YRatioToChart property usage | how to raise chart legend in Aspose.Cells .NET | legend overlay option Aspose.Cells
// Developer Intent: Place a chart legend at the bottom and raise it by a specific vertical offset.
// Use Cases: Generate a column chart from worksheet data and fine‑tune the legend’s vertical position for a cleaner layout. | Create Excel reports where the legend must stay at the bottom but avoid covering chart elements by using IsOverLay and YRatioToChart. | Automate workbook creation with custom chart styling, ensuring the legend is anchored at the bottom and slightly lifted for visual balance.
// AI Prompts: Show C# code that creates a column chart in Aspose.Cells and moves the bottom legend upward using YRatioToChart. | Explain how IsOverLay and YRatioToChart affect legend placement in Aspose.Cells charts. | Provide an example of adjusting the legend offset by 10% of the chart height in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to generate a column chart in a new workbook, set the legend to the bottom, enable overlay, and shift the legend upward using the YRatioToChart property before saving the file as an Excel workbook.
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Position the legend at the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;

        // Allow the legend to be positioned without overlapping the chart
        chart.Legend.IsOverLay = true;

        // Shift the legend upward by adjusting its relative Y position.
        // The value is a fraction of the chart height (0.0 = top, 1.0 = bottom).
        // Default for Bottom is close to 1.0; setting it to 0.85 moves it up.
        chart.Legend.YRatioToChart = 0.85;

        // Save the workbook
        workbook.Save("ChartLegendShifted.xlsx");
    }
}
