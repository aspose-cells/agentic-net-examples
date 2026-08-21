// Title: Aspose.Cells for .NET – Position Chart Legend at Bottom and Show Border
// Description: Creates a workbook, adds a column chart, moves the legend to the bottom, makes the legend border visible, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | chart legend position | legend border visibility | LegendPositionType.Bottom | Excel chart styling | Aspose.Cells chart API
// Common Searches: Aspose.Cells set legend bottom | show legend border C# | chart legend position .NET | enable legend border Aspose.Cells | Aspose.Cells legend formatting
// Developer Intent: Move the chart legend to the bottom of the chart and enable its border using Aspose.Cells for .NET.
// Use Cases: Standardize Excel reports where the legend must appear below the chart for a clean layout. | Generate dashboards that require a visible legend outline to improve data readability. | Apply corporate style guidelines that dictate legend placement and border styling across multiple charts.
// AI Prompts: Provide C# code with Aspose.Cells that positions a chart legend at the bottom and turns on its border. | Show how to configure legend placement, border visibility, and basic styling for any chart type in Aspose.Cells. | Explain the steps to adjust legend position and enable its border in an Aspose.Cells workbook using .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLegendDemo
{
    // Creates a workbook, adds a column chart, moves the legend to the bottom, makes the legend border visible, and saves the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
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
            chart.Legend.Border.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartWithBottomLegendAndBorder.xlsx");
        }
    }
}
