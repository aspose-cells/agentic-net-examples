// Title: Aspose.Cells for .NET – Set IsTotal on Final Points of a Stacked Area Chart (C#)
// Description: This C# example creates a workbook, inserts monthly product data, adds a stacked area chart, and uses LayoutProperties.Subtotals to flag the last point of each series as a total. After calling chart.Calculate, the workbook is saved, displaying cumulative totals at the chart’s end.
// Keywords: Aspose.Cells | .NET | C# | stacked area | IsTotal flag | LayoutProperties.Subtotals | chart subtotal | cumulative total | Excel chart automation | chart calculation
// Common Searches: Aspose.Cells set IsTotal flag C# | How to mark last point as total in Aspose.Cells chart | LayoutProperties.Subtotals usage example | C# stacked area chart cumulative total Aspose | Aspose.Cells chart subtotal property
// Developer Intent: Apply the IsTotal flag to the final data point of each series so the chart shows a cumulative total.
// Use Cases: Financial statements where the final month displays total sales. | Project timeline dashboards that highlight the last milestone. | Marketing analytics sheets that present total reach at the end of a trend chart.
// AI Prompts: Generate C# code with Aspose.Cells that creates a stacked area chart and sets the IsTotal flag on the last point of each series. | Explain the purpose of LayoutProperties.Subtotals and how it influences chart rendering in Aspose.Cells. | Provide a step‑by‑step tutorial for adding a stacked area chart with cumulative totals using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, inserts monthly product data, adds a stacked area chart, and uses LayoutProperties.Subtotals to flag the last point of each series as a total. After calling chart.Calculate, the workbook is saved, displaying cumulative totals at the chart’s end.
    public class StackedAreaChartIsTotalDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // First series values
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Second series values
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(45);

            // Add a stacked area chart
            int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two data series (by column)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.Add("C2:C5", true);

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Mark the last data point of each series as a subtotal (IsTotal flag)
            int lastIndexSeries0 = chart.NSeries[0].Points.Count - 1;
            int lastIndexSeries1 = chart.NSeries[1].Points.Count - 1;

            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { lastIndexSeries0 };
            chart.NSeries[1].LayoutProperties.Subtotals = new int[] { lastIndexSeries1 };

            // Force chart calculation so that the subtotal flag is applied
            chart.Calculate();

            // Save the workbook
            string outputPath = "StackedAreaChartIsTotalDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
