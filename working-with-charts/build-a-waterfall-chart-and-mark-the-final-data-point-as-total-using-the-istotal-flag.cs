// Title: Aspose.Cells C# – Build a Waterfall Chart and Flag the Final Bar as Total
// Description: Shows how to create a workbook, fill category and value columns, add a Waterfall chart, bind the data ranges, and use Series.LayoutProperties.Subtotals to designate the last point as a total before saving.
// Keywords: Aspose.Cells | C# | Waterfall chart | Series.LayoutProperties.Subtotals | mark total point | chart subtotal flag | programmatic Excel chart | Aspose.Cells chart API
// Common Searches: Aspose.Cells set total bar in waterfall chart | C# waterfall chart with subtotal flag | mark last point as total Aspose.Cells | how to use Subtotals property in Aspose chart | create waterfall chart programmatically
// Developer Intent: Create a Waterfall chart in a .NET workbook and use the Subtotals property to label the final data point as a total column.
// Use Cases: Financial statements where the ending balance appears as a highlighted total column. | Inventory reports that show adjustments and emphasize the final stock level. | Project cost analysis where cumulative expenses are presented with a total bar at the end.
// AI Prompts: Generate C# code with Aspose.Cells that builds a waterfall chart from a range and marks the last bar as a total. | Explain how Series.LayoutProperties.Subtotals can be used to flag multiple points as totals in a waterfall chart. | Outline steps to add a waterfall chart to an existing workbook, bind categories and values, and customize total bars.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // Shows how to create a workbook, fill category and value columns, add a Waterfall chart, bind the data ranges, and use Series.LayoutProperties.Subtotals to designate the last point as a total before saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A – Categories, Column B – Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);   // Initial value

            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(30);    // Positive change

            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-20);   // Negative change

            sheet.Cells["A5"].PutValue("End");
            sheet.Cells["B5"].PutValue(110);   // Final total (should be marked as total)

            // Add a waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Mark the last data point as a total using Subtotals property
            // Subtotals expects the zero‑based index of points that are totals
            Series series = chart.NSeries[0];
            int lastPointIndex = series.Points.Count - 1;
            series.LayoutProperties.Subtotals = new int[] { lastPointIndex };

            // Optional: force chart calculation so that layout properties take effect
            chart.Calculate();

            // Save the workbook with the waterfall chart
            workbook.Save("WaterfallChartWithTotal.xlsx");
        }
    }
}
