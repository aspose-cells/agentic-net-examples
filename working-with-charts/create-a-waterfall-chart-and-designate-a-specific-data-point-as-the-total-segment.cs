// Title: Create a Waterfall Chart with a Total Segment using Aspose.Cells for .NET (C#)
// Description: Learn how to generate a Waterfall chart in C# with Aspose.Cells, bind category and value ranges, and mark a specific data point as a total/subtotal segment using SeriesLayoutProperties.Subtotals before saving the workbook.
// Keywords: Aspose.Cells waterfall chart C# | waterfall chart total segment | Aspose.Cells subtotal bar | SeriesLayoutProperties.Subtotals | C# Excel chart example | Aspose.Cells chart customization
// Common Searches: Aspose.Cells mark total bar waterfall chart | C# set subtotal index Aspose.Cells waterfall | how to create waterfall chart with total segment in .NET | Aspose.Cells waterfall chart example
// Developer Intent: Generate a Waterfall chart and flag a chosen point as the total/subtotal bar.
// Use Cases: Financial statements: highlight the final total after a series of gains and losses. | Project budgeting: emphasize the cumulative budget after incremental changes. | Sales analysis: display quarterly contributions with the overall total bar distinguished.
// AI Prompts: Show C# code to add multiple subtotal points to an Aspose.Cells Waterfall chart. | Explain how to customize the color and label of the total segment in a Waterfall chart with Aspose.Cells. | Provide a script that reads data from an existing worksheet and automatically sets the last point as a total in a Waterfall chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartDemo
{
    // Learn how to generate a Waterfall chart in C# with Aspose.Cells, bind category and value ranges, and mark a specific data point as a total/subtotal segment using SeriesLayoutProperties.Subtotals before saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            // Column A – Category labels
            // Column B – Values (positive and negative)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(30);   // Initial value

            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(20);   // Positive change

            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-10);  // Negative change

            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(40);   // Total segment (will be marked as subtotal)

            // Add a Waterfall chart to the worksheet
            // Parameters: chart type, top row, left column, bottom row, right column
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Designate the "Total" data point (index 3, zero‑based) as a subtotal/total segment
            // This uses the SeriesLayoutProperties.Subtotals property
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 3 };

            // Optional: force chart layout calculation before saving
            chart.Calculate();

            // Save the workbook with the chart
            workbook.Save("WaterfallChartWithTotal.xlsx");
        }
    }
}
