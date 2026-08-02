// Title: Set all chart legends to the lower‑right corner with Aspose.Cells for .NET
// Description: The sample builds a workbook, inserts sample data and a column chart, then iterates through every worksheet and chart to move the legend to the lower‑right corner (Corner) and disables overlay, finally saving the file as ChartWithCornerLegend.xlsx.
// Keywords: Aspose.Cells | C# chart legend position | lower right legend | Corner legend Aspose | disable legend overlay | iterate charts workbook | .NET spreadsheet chart styling
// Common Searches: Aspose.Cells set legend to corner | C# move chart legend to lower right | prevent legend overlay in Aspose.Cells | apply legend position to all charts .NET | chart legend placement Aspose.Cells example
// Developer Intent: Automatically place each chart’s legend in the lower‑right corner and turn off overlay for every chart within a workbook.
// Use Cases: Generating multi‑chart reports where legends must stay out of the data area. | Standardizing legend placement across several worksheets in a single spreadsheet. | Creating a template that enforces corner‑positioned legends for any newly added chart.
// AI Prompts: Show C# code that loops through all worksheets in an Aspose.Cells workbook and sets every chart’s legend to the lower‑right corner without overlay. | Provide an Aspose.Cells for .NET example that configures chart legends to use the Corner position for all charts in a file. | Explain how to adjust legend placement and overlay settings for charts created with Aspose.Cells using C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample builds a workbook, inserts sample data and a column chart, then iterates through every worksheet and chart to move the legend to the lower‑right corner (Corner) and disables overlay, finally saving the file as ChartWithCornerLegend.xlsx.
class SetLegendPositionExample
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
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Iterate through all worksheets and their charts
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart ch in ws.Charts)
            {
                // Set legend position to the corner (bottom‑right) to avoid overlapping data
                ch.Legend.Position = LegendPositionType.Corner;

                // Ensure the legend does not overlay the chart area
                ch.Legend.IsOverLay = false;
            }
        }

        // Save the workbook
        workbook.Save("ChartWithCornerLegend.xlsx");
    }
}
