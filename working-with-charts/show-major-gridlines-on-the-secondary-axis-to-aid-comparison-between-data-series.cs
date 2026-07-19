// Title: Show major gridlines on the secondary axis of a column chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart with two series, plots the second series on the secondary value axis, enables its major gridlines (optionally changing the color), and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | secondary axis | major gridlines | column chart | chart formatting | gridline color | secondary value axis | Excel chart customization | Aspose.Cells chart example
// Common Searches: Aspose.Cells enable secondary axis gridlines | C# show major gridlines on secondary axis | column chart secondary value axis Aspose.Cells | how to set gridline color in Aspose.Cells chart | Aspose.Cells secondary axis major gridlines example
// Developer Intent: Activate and style major gridlines on a chart's secondary value axis.
// Use Cases: Add a secondary value axis to compare two data series and turn on its major gridlines for clearer visual reference. | Change the secondary axis gridline color (e.g., blue) to improve contrast with the chart background. | Generate an Excel file that displays a column chart with customized secondary‑axis gridlines for reporting or dashboards.
// AI Prompts: Write C# code using Aspose.Cells that creates a line chart, plots the second series on a secondary axis, and displays red major gridlines on that axis. | Explain how to hide primary axis gridlines while keeping secondary axis major gridlines visible in an Aspose.Cells chart. | Provide steps to modify the thickness and dash style of secondary axis major gridlines with Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSecondaryAxisGridlinesDemo
{
    // Creates a workbook, adds a column chart with two series, plots the second series on the secondary value axis, enables its major gridlines (optionally changing the color), and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two data series
            chart.NSeries.Add("B2:B4", true);   // Primary series
            chart.NSeries.Add("C2:C4", true);   // Secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary value axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Enable major gridlines on the secondary (value) axis
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.MajorGridLines.IsVisible = true;
            // Optional: set gridline color for better visibility
            secondaryAxis.MajorGridLines.Color = Color.Blue;

            // Save the workbook
            workbook.Save("SecondaryAxisMajorGridlines.xlsx");
        }
    }
}
