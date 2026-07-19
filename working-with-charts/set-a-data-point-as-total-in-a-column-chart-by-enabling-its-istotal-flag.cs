// Title: Set a Data Point as Total in an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to flag a specific column‑chart data point as a total/subtotal in Aspose.Cells for .NET by assigning the zero‑based index to LayoutProperties.Subtotals and saving the workbook.
// Keywords: Aspose.Cells C# column chart total | LayoutProperties.Subtotals | mark chart data point as total | Excel chart subtotal Aspose.Cells | set data point total Aspose.Cells | chart series total flag | Aspose.Cells chart API
// Common Searches: Aspose.Cells set data point total | C# column chart subtotal Aspose.Cells | LayoutProperties.Subtotals example | how to mark total bar in Aspose.Cells chart | Aspose.Cells chart series total flag
// Developer Intent: Apply the IsTotal flag to a chosen data point in a column chart using Aspose.Cells for .NET.
// Use Cases: Highlight a grand‑total column in a financial chart for clearer reporting. | Programmatically emphasize a specific bar in a sales chart as a subtotal. | Validate that the total index was applied by reading LayoutProperties.Subtotals after assignment.
// AI Prompts: Show how to set multiple data points as totals in an Aspose.Cells chart and style them differently. | Explain the role of LayoutProperties.Subtotals and how to retrieve the set indices at runtime. | Provide a C# example that adds a line chart and marks its last point as a total using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to flag a specific column‑chart data point as a total/subtotal in Aspose.Cells for .NET by assigning the zero‑based index to LayoutProperties.Subtotals and saving the workbook.
    public class SetDataPointAsTotalDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a column chart
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

                // Mark the second data point (index 1) as a subtotal/total
                // The Subtotals property receives an array of zero‑based indices
                chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 1 };

                // Optional: verify that the subtotal index was set
                Console.WriteLine("Subtotal index set to: " + chart.NSeries[0].LayoutProperties.Subtotals[0]);

                // Save the workbook
                string outputPath = "SetDataPointAsTotal.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetDataPointAsTotalDemo.Run();
        }
    }
}
