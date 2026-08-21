// Title: Position a Chart Legend at Exact Pixel Coordinates with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, set the legend to NotDocked, and move it to specific XPixel and YPixel values relative to the chart area, then save the file.
// Keywords: Aspose.Cells legend pixel position | C# chart legend custom coordinates | NotDocked legend Aspose.Cells | XPixel YPixel chart legend | disable legend overlay Aspose.Cells
// Common Searches: Aspose.Cells set legend XPixel YPixel C# | move chart legend to exact pixel location .NET | NotDocked legend positioning example | prevent legend overlay in Aspose.Cells chart | custom legend placement Excel export
// Developer Intent: Place a chart legend at precise pixel offsets inside the chart area.
// Use Cases: Ensure the legend does not obscure data series in automated reports. | Align the legend with surrounding UI components for a uniform layout. | Apply corporate branding rules that require the legend at a fixed screen position.
// AI Prompts: Write C# code using Aspose.Cells to set a chart legend at (150, 60) pixels and turn off overlay. | Explain how Legend.Position = NotDocked combined with XPixel and YPixel controls legend placement. | Provide a step‑by‑step tutorial for moving a chart legend to custom pixel coordinates and avoiding overlap.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendCustomPosition
{
    // Demonstrates how to create a workbook, add a column chart, set the legend to NotDocked, and move it to specific XPixel and YPixel values relative to the chart area, then save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
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

            // Add a column chart (top‑row, left‑column, bottom‑row, right‑column)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the legend
            Legend legend = chart.Legend;

            // To allow free positioning, set the legend to NotDocked
            legend.Position = LegendPositionType.NotDocked;

            // Move the legend to custom pixel coordinates relative to the chart area
            // XPixel and YPixel are measured in pixels from the upper‑left corner of the chart area
            legend.XPixel = 120; // horizontal offset in pixels
            legend.YPixel = 80;  // vertical offset in pixels

            // Optional: ensure the legend does not overlap the chart
            legend.IsOverLay = false;

            // Save the workbook (rule: save)
            workbook.Save("CustomLegendPosition.xlsx");
        }
    }
}
