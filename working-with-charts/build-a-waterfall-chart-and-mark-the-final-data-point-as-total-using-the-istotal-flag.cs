// Title: Create a Waterfall Chart with a Total Bar Using Aspose.Cells for .NET (C#)
// Description: This example builds a new workbook, adds category and value data, inserts a Waterfall chart, binds the series to B2:B6 and categories A2:A6, and uses SeriesLayoutProperties.Subtotals to flag the last point as a total before saving the file as WaterfallChart.xlsx.
// Keywords: Aspose.Cells | C# waterfall chart | Waterfall chart total flag | SeriesLayoutProperties.Subtotals | Aspose.Cells chart subtotal | Excel waterfall chart .NET | mark total point Aspose.Cells
// Common Searches: Aspose.Cells set total point waterfall chart C# | How to mark subtotal in Aspose.Cells waterfall chart | Waterfall chart total bar Aspose.Cells .NET example | SeriesLayoutProperties.Subtotals usage Aspose.Cells
// Developer Intent: Generate a Waterfall chart and designate the final bar as a total using Aspose.Cells for .NET.
// Use Cases: Financial statements that need a clear total column in a waterfall visualization. | Automated Excel reports where the ending balance is highlighted as a subtotal. | Dashboard workbooks that update dynamically and emphasize the final total bar.
// AI Prompts: Show how to set multiple subtotal indices in an Aspose.Cells waterfall chart. | Explain the behavior of SeriesLayoutProperties.Subtotals and how to apply it to different data points. | Provide code that adds data labels, custom colors, and a legend to the waterfall chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // This example builds a new workbook, adds category and value data, inserts a Waterfall chart, binds the series to B2:B6 and categories A2:A6, and uses SeriesLayoutProperties.Subtotals to flag the last point as a total before saving the file as WaterfallChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Categories, Column B – Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["A4"].PutValue("Expense");
            sheet.Cells["A5"].PutValue("Profit");
            sheet.Cells["A6"].PutValue("Total");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(0);    // Start point
            sheet.Cells["B3"].PutValue(120);  // Revenue
            sheet.Cells["B4"].PutValue(-40);  // Expense (negative)
            sheet.Cells["B5"].PutValue(80);   // Profit
            sheet.Cells["B6"].PutValue(0);    // Total (will be marked as subtotal)

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Mark the final data point (index 4, zero‑based) as a total/subtotal
            // Using SeriesLayoutProperties.Subtotals to specify the index of the total point
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 4 };

            // Calculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("WaterfallChart.xlsx");
        }
    }
}
