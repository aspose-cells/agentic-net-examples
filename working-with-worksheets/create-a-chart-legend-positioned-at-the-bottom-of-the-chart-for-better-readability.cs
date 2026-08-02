// Title: Aspose.Cells for .NET: Add a Column Chart with Legend at Bottom
// Description: Creates a new Workbook, fills sample category and value data, inserts a column chart, sets the legend position to the bottom using Legend.Position = LegendPositionType.Bottom, and saves the file as ChartWithBottomLegend.xlsx.
// Keywords: Aspose.Cells chart legend bottom | C# set chart legend position | Aspose.Cells column chart example | LegendPositionType.Bottom | save workbook Aspose.Cells | Excel chart legend placement .NET
// Common Searches: Aspose.Cells set legend to bottom C# | how to move chart legend below chart Aspose.Cells | column chart with bottom legend Aspose .NET | chart legend positioning Aspose.Cells | C# Aspose.Cells chart legend placement
// Developer Intent: Place the legend beneath a column chart and generate an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a monthly sales report where the column chart legend appears below the chart for optimal print layout. | Build an executive dashboard that keeps all chart legends at the bottom to maximize data‑visual area. | Automate financial statements with charts that have bottom legends, ensuring a clean appearance when exported.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart whose legend is positioned at the bottom and save it as an .xlsx file. | Show how to set Legend.Position = LegendPositionType.Bottom for any chart type in Aspose.Cells for .NET. | Explain how to adjust the size and alignment of a chart legend after moving it to the bottom using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new Workbook, fills sample category and value data, inserts a column chart, sets the legend position to the bottom using Legend.Position = LegendPositionType.Bottom, and saves the file as ChartWithBottomLegend.xlsx.
class ChartLegendBottomExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
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

        // Save the workbook (lifecycle: save)
        workbook.Save("ChartWithBottomLegend.xlsx");
    }
}
