// Title: Set a Data Point as Total in an Aspose.Cells Column Chart (C#)
// Description: Demonstrates how to create a workbook, add a column chart, and flag a specific data point as a total/subtotal by using the LayoutProperties.Subtotals array in Aspose.Cells for .NET.
// Keywords: Aspose.Cells column chart total point | C# Aspose.Cells IsTotal flag | LayoutProperties.Subtotals example | mark data point as total Aspose.Cells | .NET chart subtotal flag | Aspose.Cells chart customization
// Common Searches: Aspose.Cells set total flag column chart C# | How to mark a subtotal point in Aspose.Cells chart | C# column chart total data point Aspose | Enable IsTotal for a chart series in Aspose.Cells | Aspose.Cells column chart subtotal example
// Developer Intent: Add a column chart and designate a chosen data point as a total/subtotal in a .NET workbook.
// Use Cases: Generate a sales report where the final column shows the overall total. | Create a financial dashboard that highlights cumulative totals within a column chart. | Build a project‑status workbook that emphasizes the overall progress column as a total.
// AI Prompts: Show C# code to set the IsTotal flag for a column chart data point using Aspose.Cells. | Provide an Aspose.Cells example that marks the third column as a total with LayoutProperties.Subtotals. | Explain how to retrieve and modify the total/subtotal setting of a chart series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a column chart, and flag a specific data point as a total/subtotal by using the LayoutProperties.Subtotals array in Aspose.Cells for .NET.
    public class ColumnChartTotalPointDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
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

            // Mark the third data point (index 2) as a subtotal/total
            chart.NSeries[0].LayoutProperties.Subtotals = new int[] { 2 };

            // Save the workbook
            string outputPath = "ColumnChartWithTotalPoint.xlsx";
            workbook.Save(outputPath);
        }
    }
}
