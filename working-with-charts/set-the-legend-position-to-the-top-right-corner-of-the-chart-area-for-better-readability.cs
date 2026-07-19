// Title: Aspose.Cells for .NET – Move Chart Legend to Top‑Right Corner
// Description: Shows how to create a workbook, add a column chart, define its data range, and place the legend in the chart’s top‑right corner using LegendPositionType.Corner, then save the file as LegendTopRightCorner.xlsx.
// Keywords: Aspose.Cells chart legend position | C# LegendPositionType Corner | top right legend Aspose.Cells | Excel chart customization .NET | Aspose.Cells set legend location | chart legend placement C# | Aspose.Cells example column chart
// Common Searches: Aspose.Cells set chart legend to top right | C# move legend corner Aspose.Cells | LegendPositionType Corner example | how to place chart legend in corner using Aspose.Cells | Aspose.Cells chart legend positioning tutorial
// Developer Intent: Place the chart legend in the top‑right corner of the chart area.
// Use Cases: Generate a column chart with the legend anchored at the top‑right to keep the data series visible. | Create multiple Excel charts where each legend is consistently positioned in the corner for a clean layout. | Export analytical charts to Excel while preventing legend overlap with plotted values by fixing it to the top‑right corner.
// AI Prompts: Write C# code with Aspose.Cells that creates a line chart and sets its legend to the top‑right corner. | Provide a snippet to change the legend position of an existing Aspose.Cells chart to Corner and save the workbook. | Explain how to adjust legend placement for different chart types (column, line, pie) using Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendPositionDemo
{
    // Shows how to create a workbook, add a column chart, define its data range, and place the legend in the chart’s top‑right corner using LegendPositionType.Corner, then save the file as LegendTopRightCorner.xlsx.
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
            chart.SetChartDataRange("A1:B4", true);

            // Position the legend at the top‑right corner of the chart area
            chart.Legend.Position = LegendPositionType.Corner;

            // Save the workbook
            workbook.Save("LegendTopRightCorner.xlsx");
        }
    }
}
