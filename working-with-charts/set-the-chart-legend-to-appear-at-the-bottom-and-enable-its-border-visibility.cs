// Title: Aspose.Cells for .NET: Position Chart Legend at Bottom and Show Its Border
// Description: This C# example creates a workbook, adds a column chart with sample data, moves the legend to the bottom using Legend.Position = Bottom, enables the legend border with Legend.Border.IsVisible = true, and saves the result as ChartLegendBottomWithBorder.xlsx.
// Keywords: Aspose.Cells | C# chart legend position | legend bottom Aspose.Cells | show legend border | Legend.Position Bottom | Legend.Border.IsVisible | Excel chart formatting .NET | Aspose.Cells chart customization | column chart legend styling | Aspose.Cells API
// Common Searches: Aspose.Cells set legend to bottom | Enable legend border in Aspose.Cells chart | C# move chart legend below chart | Aspose.Cells legend formatting example | How to show legend border in Excel using Aspose.Cells
// Developer Intent: Place the chart legend at the bottom of a chart and make its border visible using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports where the legend is positioned beneath the chart for better layout. | Create dashboards that require a visible legend border to separate legend items from the chart area. | Apply consistent legend placement and border styling across multiple charts in a workbook.
// AI Prompts: Write C# code with Aspose.Cells to set a chart legend at the bottom and enable its border. | Show how to customize the legend border color and thickness after making it visible in an Aspose.Cells chart. | Explain how to apply the same bottom legend position and border settings to all charts in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendDemo
{
    // This C# example creates a workbook, adds a column chart with sample data, moves the legend to the bottom using Legend.Position = Bottom, enables the legend border with Legend.Border.IsVisible = true, and saves the result as ChartLegendBottomWithBorder.xlsx.
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Make the legend border visible
            // The Border property returns a Line object; set its visibility to true
            chart.Legend.Border.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartLegendBottomWithBorder.xlsx");
        }
    }
}
