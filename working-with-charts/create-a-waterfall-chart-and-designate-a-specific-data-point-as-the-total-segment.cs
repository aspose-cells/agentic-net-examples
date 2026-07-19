// Title: C# – Create a Waterfall Chart with a Total Segment using Aspose.Cells
// Description: Learn how to generate a Waterfall chart in Aspose.Cells for .NET, bind category and value ranges, and mark a specific data point (e.g., Net Profit) as a total segment by using the LayoutProperties.Subtotals property, then save the workbook as an Excel file.
// Keywords: Aspose.Cells waterfall chart C# | set total segment Aspose.Cells | LayoutProperties.Subtotals example | waterfall chart subtotal index | Aspose.Cells chart tutorial | .NET Excel waterfall chart
// Common Searches: Aspose.Cells how to mark total in waterfall chart | C# waterfall chart subtotal property | Create waterfall chart with total segment Aspose.Cells | Aspose.Cells Waterfall chart Subtotals example | Set total column in Excel waterfall using Aspose
// Developer Intent: Generate a Waterfall chart and flag a chosen data point as the total segment in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Financial statements: display revenue, costs, and highlight net profit as the total column. | Project budgeting: show incremental expenses and emphasize the final balance as a subtotal. | Automated reporting: embed a Waterfall chart with a designated total segment into generated Excel reports.
// AI Prompts: Write C# code with Aspose.Cells to create a Waterfall chart from a range and set the last point as a total using LayoutProperties.Subtotals. | Explain how to configure multiple total segments in an Aspose.Cells Waterfall chart by providing an array of indices to Subtotals. | Show how to change the color and style of the total segment in a Waterfall chart created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // Learn how to generate a Waterfall chart in Aspose.Cells for .NET, bind category and value ranges, and mark a specific data point (e.g., Net Profit) as a total segment by using the LayoutProperties.Subtotals property, then save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            // Column A – Category names
            // Column B – Values (positive, negative, and a total)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(5000);   // Positive

            sheet.Cells["A3"].PutValue("Cost of Goods Sold");
            sheet.Cells["B3"].PutValue(-2000); // Negative

            sheet.Cells["A4"].PutValue("Operating Expenses");
            sheet.Cells["B4"].PutValue(-1500); // Negative

            sheet.Cells["A5"].PutValue("Net Profit");
            sheet.Cells["B5"].PutValue(1500);  // This will be marked as the total segment

            // Add a Waterfall chart to the worksheet
            // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Designate the "Net Profit" point (index 3, zero‑based) as the total segment
            // Using SeriesLayoutProperties.Subtotals property
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 3 };

            // Optional: calculate the chart to ensure all internal data is up‑to‑date
            chart.Calculate();

            // Save the workbook with the chart
            workbook.Save("WaterfallChartWithTotal.xlsx");
        }
    }
}
