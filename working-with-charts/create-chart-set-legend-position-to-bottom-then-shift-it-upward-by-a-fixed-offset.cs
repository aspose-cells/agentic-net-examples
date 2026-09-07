// Title: Create a column chart in Aspose.Cells for .NET, set the legend to the bottom, and shift it upward by a pixel offset (C#)
// AI Prompts: Generate C# code using Aspose.Cells that adds a column chart, docks the legend at the bottom, and moves the legend up by a specified number of pixels. | Show how to modify the YPixel property of a chart legend after assigning Legend.Position = Bottom in Aspose.Cells.
// Common Searches: Aspose.Cells C# how to move a bottom‑docked chart legend upward | adjust legend YPixel offset in Excel chart using Aspose.Cells .NET | shift chart legend position by pixels after setting Legend.Position in Aspose.Cells
// Tags: Aspose.Cells legend YPixel offset | C# set chart legend bottom | shift chart legend upward | create column chart Aspose.Cells | save workbook Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendShift
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, positions the legend at the bottom, then raises the legend by decreasing its YPixel value by 20 pixels, and finally saves the file as ChartWithShiftedLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
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

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend position to Bottom
            chart.Legend.Position = LegendPositionType.Bottom;

            // Shift the legend upward by a fixed pixel offset (e.g., 20 pixels)
            // Since the legend is docked at the bottom, we adjust its YPixel value directly.
            // Decreasing YPixel moves the legend upward.
            const int offsetPixels = 20;
            chart.Legend.YPixel = chart.Legend.YPixel - offsetPixels;

            // Save the workbook (lifecycle: save)
            workbook.Save("ChartWithShiftedLegend.xlsx");
        }
    }
}
