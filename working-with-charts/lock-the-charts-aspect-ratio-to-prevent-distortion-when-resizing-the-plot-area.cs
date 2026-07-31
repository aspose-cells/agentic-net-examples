// Title: Lock Chart Aspect Ratio in Aspose.Cells for .NET to Prevent Distortion
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and lock its aspect ratio using ChartObject.IsAspectRatioLocked (or SetLockedProperty) so the chart stays proportional when the plot area is resized.
// Keywords: Aspose.Cells chart aspect ratio | IsAspectRatioLocked C# | Aspose.Cells prevent chart distortion | .NET Excel chart resizing | ChartObject SetLockedProperty | lock chart size Aspose.Cells | Excel chart scaling C#
// Common Searches: lock aspect ratio of chart Aspose.Cells .NET | prevent chart distortion when resizing Excel with Aspose | ChartObject.IsAspectRatioLocked example | set locked property chart Aspose.Cells C# | how to keep chart proportions in Aspose.Cells
// Developer Intent: Ensure a chart retains its original proportions by locking its aspect ratio before resizing the plot area.
// Use Cases: Create a column chart from worksheet data and lock its aspect ratio before exporting the workbook. | Apply aspect‑ratio locking to all charts in a template to maintain consistent visual scaling across reports. | Toggle the lock at runtime based on user preferences, then generate the final Excel file.
// AI Prompts: Generate C# code that adds a line chart with Aspose.Cells and locks its aspect ratio. | Show how to unlock a chart’s aspect ratio and programmatically adjust the plot area using Aspose.Cells. | Provide a script that iterates through every chart in an existing workbook and sets IsAspectRatioLocked to true.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, and lock its aspect ratio using ChartObject.IsAspectRatioLocked (or SetLockedProperty) so the chart stays proportional when the plot area is resized.
    public class LockChartAspectRatioDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Lock the aspect ratio of the chart to prevent distortion when resizing
                chart.ChartObject.IsAspectRatioLocked = true;
                // Alternatively:
                // chart.ChartObject.SetLockedProperty(ShapeLockType.AspectRatio, true);

                // Save the workbook
                string outputPath = "LockChartAspectRatioDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LockChartAspectRatioDemo.Run();
        }
    }
}
