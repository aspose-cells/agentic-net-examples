// Title: C# – Create a Waterfall Chart with a Total Segment using Aspose.Cells
// Description: Demonstrates how to build a workbook, add month labels and values, insert a Waterfall chart, bind series and categories, and mark a specific point as a total segment via LayoutProperties.Subtotals, then save as WaterfallChart.xlsx.
// Keywords: Aspose.Cells | C# waterfall chart | total segment | LayoutProperties.Subtotals | chart subtotal index | .NET charting | financial waterfall
// Common Searches: Aspose.Cells set total bar waterfall chart C# | how to mark subtotal in Aspose.Cells waterfall | waterfall chart total segment .NET | LayoutProperties.Subtotals example | create waterfall chart Aspose.Cells
// Developer Intent: Add a Waterfall chart to a workbook and define a chosen data point as the total segment.
// Use Cases: Financial statements where the final month shows the cumulative total. | Sales performance analysis highlighting the overall result as a total bar. | Project cost breakdown with the last column representing the total expense.
// AI Prompts: Generate C# code that creates a Waterfall chart in Aspose.Cells and sets the last point as a total segment. | Show how to assign multiple total segments in an Aspose.Cells Waterfall chart. | Explain customization options for the appearance of total bars in Aspose.Cells waterfall charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsWaterfallDemo
{
    // Demonstrates how to build a workbook, add month labels and values, insert a Waterfall chart, bind series and categories, and mark a specific point as a total segment via LayoutProperties.Subtotals, then save as WaterfallChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            // Category labels
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["A6"].PutValue("May");

            // Corresponding values (positive and negative)
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(-5);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(-3);
            sheet.Cells["B6"].PutValue(20);

            // Add a Waterfall chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series and the category axis
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Designate the last data point (index 4, zero‑based) as a total segment
            // using the Subtotals property of LayoutProperties
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 4 };

            // Recalculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook with the waterfall chart
            workbook.Save("WaterfallChart.xlsx");
        }
    }
}
