// Title: Aspose.Cells .NET: Move Column Chart Legend to Bottom
// Description: C# sample that builds a workbook, inserts sample data, creates a column chart, positions the legend beneath the plot area, disables overlay to keep the chart clear, and writes the result to ChartWithBottomLegend.xlsx.
// Keywords: Aspose.Cells | C# | chart legend placement | bottom legend | LegendPositionType.Bottom | disable legend overlay | column chart example | Excel file generation | Aspose.Cells chart API
// Common Searches: Aspose.Cells place legend below chart | C# set chart legend to bottom Aspose.Cells | prevent legend overlapping plot area Aspose.Cells | how to move chart legend to bottom in .NET | Aspose.Cells legend positioning options
// Developer Intent: Position the chart legend under the plot area and stop it from covering the data series.
// Use Cases: Automated sales reports where the column chart’s legend is shown under the chart for a cleaner printed layout. | Dashboard exports that require a bottom‑placed legend to improve readability on widescreen displays. | Financial summaries generated programmatically where each chart’s legend must stay outside the plot area.
// AI Prompts: Provide C# code using Aspose.Cells to add a column chart and set the legend at the bottom without overlay. | Explain the effect of Legend.Position and IsOverLay properties on chart layout in Aspose.Cells for .NET. | Show how to ensure a chart legend stays outside the plot area while positioned below the chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# sample that builds a workbook, inserts sample data, creates a column chart, positions the legend beneath the plot area, disables overlay to keep the chart clear, and writes the result to ChartWithBottomLegend.xlsx.
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

        // Position the legend at the bottom of the chart area
        chart.Legend.Position = LegendPositionType.Bottom;

        // Optional: ensure the legend does not overlap the plot area
        chart.Legend.IsOverLay = false;

        // Save the workbook to an XLSX file
        workbook.Save("ChartWithBottomLegend.xlsx");
    }
}
