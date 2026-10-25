// Title: Center a chart legend at the bottom and apply a fully transparent fill without altering its alignment using Aspose.Cells for .NET
// Common Searches: c# aspose.cells set chart legend to bottom and keep it centered | apply 100% transparent fill to a chart legend in aspose.cells without moving it | how to verify legend position after changing fill transparency with aspose.cells | asp.net aspose.cells chart legend alignment after setting fill transparency | make chart legend transparent while preserving bottom alignment in aspose.cells
// Tags: set legend position bottom Aspose.Cells | center legend alignment Aspose.Cells | legend fill transparency Aspose.Cells | verify legend placement Aspose.Cells | solid fill full transparency chart legend Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a column chart, positions the legend at the bottom (centered), applies a solid fill with 100 % transparency to the legend area, confirms the legend remains bottom‑centered, and saves the file as LegendAlignmentDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set legend position to bottom (centered)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Apply a fully transparent fill to the legend (Transparency = 1.0 means 100% transparent)
            chart.Legend.Area.FillFormat.FillType = FillType.Solid;
            chart.Legend.Area.FillFormat.Transparency = 1.0;

            // Verify that the legend is positioned at the bottom
            bool isPositionBottom = chart.Legend.Position == LegendPositionType.Bottom;
            Console.WriteLine("Legend position is bottom (centered): " + isPositionBottom);

            // Save the workbook
            workbook.Save("LegendAlignmentDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
