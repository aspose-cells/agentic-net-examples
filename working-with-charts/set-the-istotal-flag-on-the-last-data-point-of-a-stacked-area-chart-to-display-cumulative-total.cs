// Title: Aspose.Cells for .NET: Mark the last point as cumulative total (IsTotal) in a stacked area chart (C#)
// Description: Shows how to build an Excel workbook, fill category and series values, add a stacked area chart, and apply LayoutProperties.Subtotals to set the IsTotal flag on the final data point of each series before recalculating and saving the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | stacked area chart | IsTotal flag | LayoutProperties.Subtotals | cumulative total chart | Excel chart automation | chart subtotal | GitHub example
// Common Searches: Aspose.Cells set IsTotal flag stacked area chart | mark last data point as total Aspose.Cells .NET | LayoutProperties.Subtotals usage example | cumulative total in Excel chart using Aspose.Cells | stacked area chart subtotal C#
// Developer Intent: Apply the IsTotal flag to the last point of each series in a stacked area chart with Aspose.Cells for .NET.
// Use Cases: Financial reports where the final month displays the total revenue across all categories. | Performance dashboards that highlight the overall cumulative value at the end of a time‑series chart. | Automated sales analysis workbooks that emphasize the aggregated sales figure in the chart’s last point.
// AI Prompts: Generate C# code using Aspose.Cells to set the IsTotal flag on the last point of a stacked area chart. | Explain how LayoutProperties.Subtotals works for marking cumulative totals in Aspose.Cells charts and note any constraints. | Suggest an alternative method to show a cumulative total in a stacked area chart with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build an Excel workbook, fill category and series values, add a stacked area chart, and apply LayoutProperties.Subtotals to set the IsTotal flag on the final data point of each series before recalculating and saving the file with Aspose.Cells for .NET.
    public class StackedAreaChartIsTotalDemo
    {
        public static void Run()
        {
            try
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
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Second series values
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // Add a stacked area chart
                int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set data ranges for the series
                chart.NSeries.Add("B2:B5", true); // Series1
                chart.NSeries.Add("C2:C5", true); // Series2
                chart.NSeries.CategoryData = "A2:A5";

                // Determine the index of the last data point (zero‑based)
                int lastPointIndex = chart.NSeries[0].Points.Count - 1;

                // Mark the last point as a subtotal (cumulative total) for each series
                chart.NSeries[0].LayoutProperties.Subtotals = new int[] { lastPointIndex };
                chart.NSeries[1].LayoutProperties.Subtotals = new int[] { lastPointIndex };

                // Recalculate the chart so that the subtotal markers are generated
                chart.Calculate();

                // Save the workbook
                workbook.Save("StackedAreaChartIsTotalDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StackedAreaChartIsTotalDemo.Run();
        }
    }
}
