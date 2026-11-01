// Title: How to position a chart legend at the bottom‑left and offset its X coordinate to align with chart margins using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart, set the legend position to Bottom, then move the legend left by half the plot area width with Aspose.Cells in C#. | Adjust the X property of a chart's Legend so it aligns with the left edge of the plot area after placing the legend at the bottom using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart legend to bottom left and align with plot area | C# Aspose.Cells move chart legend X offset relative to plot area width | how to shift chart legend left after setting it to bottom in Aspose.Cells | Aspose.Cells chart legend positioning bottom left offset example | C# code to adjust chart legend X coordinate in Aspose.Cells workbook
// Tags: set legend position bottom Aspose.Cells | adjust legend X offset Aspose.Cells | align chart legend with plot area margins C# | Aspose.Cells column chart legend alignment | chart legend positioning bottom left .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart, places the legend at the bottom, shifts its X coordinate left by half the plot area width to match the chart margins, and saves the file as ChartWithLegendBottomLeft.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart (rows 5‑20, columns 0‑10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Place the legend at the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Shift the legend left by half of the plot area width
            chart.Legend.X = (int)(-chart.PlotArea.Width / 2.0);

            // Save the workbook
            string outputPath = "ChartWithLegendBottomLeft.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
