// Title: Set a Data Point as Total in an Aspose.Cells Column Chart (C#)
// Description: Shows how to build an Excel workbook, add a column chart, and flag a specific series point as a total/subtotal using the LayoutProperties.Subtotals property in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | column chart total point | C# | .NET | LayoutProperties.Subtotals | IsTotal flag | chart series total | Excel automation | data visualization | Aspose.Cells chart example
// Common Searches: Aspose.Cells column chart total point C# | How to set IsTotal flag in Aspose.Cells chart | Mark subtotal data point in Excel chart using Aspose.Cells | C# Aspose.Cells chart total data point example | LayoutProperties.Subtotals Aspose.Cells tutorial
// Developer Intent: Find the code pattern to mark a specific chart point as a total in an Aspose.Cells column chart with C#.
// Use Cases: Highlight the final quarter's sales column as a cumulative total in a performance dashboard. | Display an intermediate subtotal column to separate product groups in a financial report. | Create a stacked column chart where the last column represents the overall project cost.
// AI Prompts: Generate C# code that sets the IsTotal flag for a data point in an Aspose.Cells column chart. | Explain how LayoutProperties.Subtotals works compared to manually setting IsTotal in Aspose.Cells. | Provide an example of marking multiple points as totals in a single Aspose.Cells chart series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build an Excel workbook, add a column chart, and flag a specific series point as a total/subtotal using the LayoutProperties.Subtotals property in Aspose.Cells for .NET.
    public class ColumnChartTotalPointDemo
    {
        public static void Run()
        {
            try
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

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Mark the third data point (index 2) as a subtotal/total point
                chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 2 };

                // Save the workbook
                string outputPath = "ColumnChartWithTotalPoint.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ColumnChartTotalPointDemo.Run();
        }
    }
}
