// Title: Aspose.Cells .NET: Position Chart Legend Bottom‑Right and Apply Calibri 10 Font
// Description: Creates a workbook, adds a column chart, moves the legend to the bottom‑right corner, disables overlay, and sets the legend font to Calibri size 10 before saving the file.
// Keywords: Aspose.Cells chart legend position | bottom right legend .NET | legend font Calibri 10 | Aspose.Cells C# legend formatting | Excel chart legend placement | Aspose.Cells Legend.IsOverLay | set chart legend font Aspose.Cells
// Common Searches: Aspose.Cells move legend to bottom right | change chart legend font to Calibri in C# | set legend position corner Aspose.Cells | disable legend overlay Aspose.Cells chart | Aspose.Cells legend formatting example
// Developer Intent: Place the chart legend in the bottom‑right corner and format its text with Calibri size 10 using Aspose.Cells for .NET.
// Use Cases: Design Excel reports where the legend must stay out of the plot area for better readability. | Apply corporate branding by using the standard Calibri 10 font for all chart legends. | Generate workbooks that require precise legend placement without overlapping chart data.
// AI Prompts: Show how to set a chart legend to the bottom‑right corner and change its font to Calibri 10 with Aspose.Cells in C#. | Provide a C# code example that configures Legend.Position, Legend.IsOverLay, and Legend.Font for an Aspose.Cells chart. | Explain the effect of Legend.IsOverLay when positioning a legend at the corner of an Excel chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, moves the legend to the bottom‑right corner, disables overlay, and sets the legend font to Calibri size 10 before saving the file.
class ChartLegendExample
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.SetChartDataRange("A1:B4", true);

        // Access the legend and configure its position and font
        Legend legend = chart.Legend;

        // Position the legend at the corner (bottom‑right) of the plot area
        legend.Position = LegendPositionType.Corner;

        // Optionally ensure the legend does not overlap the chart area
        legend.IsOverLay = false;

        // Set the legend font to Calibri, size 10
        legend.Font.Name = "Calibri";
        legend.Font.Size = 10;

        // Save the workbook to an XLSX file
        workbook.Save("ChartWithBottomRightLegend.xlsx");
    }
}
