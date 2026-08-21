// Title: C# – Remove a Chart Control from a Chart Sheet using Shapes.RemoveAt in Aspose.Cells
// Description: Creates a workbook, adds a column chart to the first worksheet (as a chart sheet), deletes the chart shape with Shapes.RemoveAt(0), and saves the file to confirm removal.
// Keywords: Aspose.Cells remove chart shape C# | Shapes.RemoveAt chart sheet | delete chart control Aspose.Cells .NET | remove chart from worksheet programmatically | Aspose.Cells chart sheet manipulation | C# chart shape removal
// Common Searches: how to delete a chart shape from a chart sheet using Aspose.Cells | remove chart control from worksheet C# Aspose.Cells | Shapes.RemoveAt example for charts Aspose.Cells | Aspose.Cells delete chart from workbook programmatically
// Developer Intent: Delete the chart control (shape) on a chart sheet by invoking Shapes.RemoveAt with the correct index.
// Use Cases: Erase a temporary chart after data analysis to keep the workbook lightweight. | Clean up unwanted chart shapes before publishing a final report. | Programmatically remove a specific chart when a user cancels a chart‑creation operation.
// AI Prompts: Generate C# code that adds multiple charts to a worksheet and removes the chart at shape index 2 using Aspose.Cells. | Explain how to determine the shape index of a particular chart before calling Shapes.RemoveAt in Aspose.Cells. | Provide a complete example that removes a chart control from a chart sheet and saves the workbook without the chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart to the first worksheet (as a chart sheet), deletes the chart shape with Shapes.RemoveAt(0), and saves the file to confirm removal.
    public class RemoveChartControlFromChartSheet
    {
        // Entry point for the console application
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (used as a chart sheet for this demo)
            Worksheet chartSheet = workbook.Worksheets[0];

            // Add sample data for the chart
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["A2"].PutValue("A");
            chartSheet.Cells["A3"].PutValue("B");
            chartSheet.Cells["A4"].PutValue("C");
            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["B2"].PutValue(10);
            chartSheet.Cells["B3"].PutValue(20);
            chartSheet.Cells["B4"].PutValue(30);

            // Add a chart to the sheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = chartSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Remove the chart control (shape) from the Shapes collection
            // The chart we just added is the first (and only) shape, so its index is 0.
            chartSheet.Shapes.RemoveAt(0);

            // Save the workbook to verify that the chart has been removed
            string outputPath = "RemoveChartControlDemo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Chart control removed and workbook saved to '{outputPath}'.");
        }
    }
}
