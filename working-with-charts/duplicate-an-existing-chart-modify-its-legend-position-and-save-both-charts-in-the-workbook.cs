// Title: Duplicate a Chart, Change Its Legend Position, and Save the Workbook with Aspose.Cells (C#)
// Description: Creates a workbook, adds a column chart from sample data, clones the chart to a new location, moves the cloned chart's legend to the left, and saves the file containing both charts.
// Keywords: Aspose.Cells | C# | .NET | duplicate chart | clone chart | chart legend position | set legend left | add multiple charts | Excel chart manipulation | save workbook with charts
// Common Searches: Aspose.Cells copy chart C# | How to duplicate a chart in Aspose.Cells | Set legend position left Aspose.Cells | Add two charts with same data range Aspose.Cells | Clone Excel chart using Aspose.Cells .NET
// Developer Intent: Copy an existing chart, adjust the legend placement of the copy, and persist both charts in the same workbook.
// Use Cases: Show a summary chart alongside a detailed version for side‑by‑side comparison. | Reuse a chart template in different report sections while customizing legend orientation. | Generate a workbook that contains the original chart and a duplicated chart with a left‑aligned legend for presentation.
// AI Prompts: Generate C# code with Aspose.Cells to duplicate a chart and keep its data range. | Show how to set the legend position to left for a cloned chart in Aspose.Cells. | Explain adding multiple charts from the same source and saving the workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart from sample data, clones the chart to a new location, moves the cloned chart's legend to the left, and saves the file containing both charts.
class DuplicateChartExample
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

        // -----------------------------------------------------------------
        // 1. Add the original chart
        // -----------------------------------------------------------------
        int originalChartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart originalChart = sheet.Charts[originalChartIndex];
        originalChart.NSeries.Add("B2:B4", true);          // Values
        originalChart.NSeries.CategoryData = "A2:A4";     // Categories

        // -----------------------------------------------------------------
        // 2. Duplicate the original chart
        // -----------------------------------------------------------------
        // Retrieve the type and data range of the original chart
        ChartType chartType = originalChart.Type;
        string dataRange = originalChart.GetChartDataRange(); // e.g., "A1:B4"

        // Add a new chart with the same type and data range at a different position
        int duplicatedChartIndex = sheet.Charts.Add(chartType, dataRange, true, 16, 0, 26, 5);
        Chart duplicatedChart = sheet.Charts[duplicatedChartIndex];

        // -----------------------------------------------------------------
        // 3. Modify the legend position of the duplicated chart
        // -----------------------------------------------------------------
        duplicatedChart.Legend.Position = LegendPositionType.Left; // Move legend to the left side

        // -----------------------------------------------------------------
        // 4. Save the workbook containing both charts
        // -----------------------------------------------------------------
        workbook.Save("DuplicatedChart.xlsx");
    }
}
