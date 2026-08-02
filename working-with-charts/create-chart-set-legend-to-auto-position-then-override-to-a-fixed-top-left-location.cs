// Title: Aspose.Cells C# – Create a Column Chart, Set Legend to Auto, Then Move It to a Fixed Top‑Left Position
// Description: This example shows how to build a workbook, add sample data, insert a column chart, apply automatic legend placement with SetPositionAuto(), switch the legend to NotDocked, and assign XPixel/YPixel values (50, 20) to position the legend precisely at the top‑left of the chart before saving the file.
// Keywords: Aspose.Cells chart legend positioning | SetPositionAuto Aspose.Cells | Legend NotDocked C# | Aspose.Cells XPixel YPixel | column chart legend placement | Aspose.Cells C# example | custom chart legend coordinates
// Common Searches: Aspose.Cells set legend auto then custom position | C# move chart legend to specific X Y coordinates Aspose.Cells | How to use Legend.Position NotDocked in Aspose.Cells | Aspose.Cells chart legend pixel placement | auto legend positioning Aspose.Cells chart
// Developer Intent: Create a chart, let the legend auto‑place, then override it with exact pixel coordinates.
// Use Cases: Standardize report layouts by fixing legend locations across multiple worksheets. | Align chart legends with corporate branding guidelines that require precise placement. | Generate templates where the legend must appear at a consistent spot regardless of data size.
// AI Prompts: Generate C# code with Aspose.Cells that adds a line chart, calls SetPositionAuto() on the legend, then repositions it to (100, 30) pixels from the chart's top‑left corner. | Explain why setting Legend.Position = LegendPositionType.NotDocked is required for manual XPixel/YPixel adjustments in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to build a workbook, add sample data, insert a column chart, apply automatic legend placement with SetPositionAuto(), switch the legend to NotDocked, and assign XPixel/YPixel values (50, 20) to position the legend precisely at the top‑left of the chart before saving the file.
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the chart legend
        Legend legend = chart.Legend;

        // 1. Set legend to automatic positioning
        legend.SetPositionAuto();

        // 2. Override to a fixed top‑left location
        //    Use NotDocked so manual X/Y coordinates are respected
        legend.Position = LegendPositionType.NotDocked;
        legend.XPixel = 50;   // X coordinate in pixels from the left edge of the chart area
        legend.YPixel = 20;   // Y coordinate in pixels from the top edge of the chart area

        // Save the workbook with the configured chart
        workbook.Save("ChartLegendAutoThenFixed.xlsx");
    }
}
