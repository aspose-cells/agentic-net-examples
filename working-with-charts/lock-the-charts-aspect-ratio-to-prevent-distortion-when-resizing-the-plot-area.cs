// Title: Lock chart aspect ratio in Aspose.Cells for .NET to avoid distortion
// Description: Creates a workbook, adds sample data, inserts a column chart and locks its aspect ratio with ChartObject.IsAspectRatioLocked (or SetLockedProperty) so the chart stays proportional when the plot area is resized.
// Keywords: Aspose.Cells | .NET | lock chart aspect ratio | IsAspectRatioLocked | SetLockedProperty | chart resizing | prevent chart distortion | Excel chart shape lock | ChartObject | column chart
// Common Searches: Aspose.Cells lock chart aspect ratio | prevent chart distortion in Excel using Aspose.Cells | set chart shape lock .NET | ChartObject IsAspectRatioLocked example | resize chart without changing proportions Aspose.Cells
// Developer Intent: Ensure a chart retains its original proportions when the plot area is resized by locking its aspect ratio.
// Use Cases: Generate Excel reports with column charts that keep consistent proportions across pages. | Programmatically add multiple charts to a workbook and lock each one to maintain visual integrity. | Prepare workbooks for downstream applications where chart distortion must be avoided.
// AI Prompts: Show how to lock a chart's aspect ratio in Aspose.Cells for .NET and then resize it programmatically. | Provide code that uses ChartObject.SetLockedProperty to lock aspect ratios for different chart types. | Explain the difference between IsAspectRatioLocked and SetLockedProperty when working with chart objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart and locks its aspect ratio with ChartObject.IsAspectRatioLocked (or SetLockedProperty) so the chart stays proportional when the plot area is resized.
    public class LockChartAspectRatioDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
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

            // Save the workbook (lifecycle rule: save)
            workbook.Save("LockChartAspectRatioDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                LockChartAspectRatioDemo.Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
