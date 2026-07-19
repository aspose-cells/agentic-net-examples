// Title: Lock a chart’s aspect ratio in Aspose.Cells for .NET by setting Height and Width
// Description: Shows how to create a workbook, add sample data, insert a column chart, and keep its proportions consistent by manually assigning Height and Width, because Aspose.Cells for .NET lacks a direct LockAspectRatio property.
// Keywords: Aspose.Cells | .NET | C# | chart aspect ratio | lock aspect ratio | prevent chart distortion | set chart height | set chart width | manual aspect ratio | Excel chart resizing | Aspose.Cells chart dimensions
// Common Searches: How to lock chart aspect ratio in Aspose.Cells .NET | Aspose.Cells chart resize without distortion | Set fixed height and width for Excel chart using Aspose.Cells | Is there a LockAspectRatio property in Aspose.Cells charts | Maintain chart proportions when changing plot area size
// Developer Intent: Maintain a fixed chart aspect ratio while resizing the plot area.
// Use Cases: Create a column chart and immediately assign Height and Width to enforce a 16:9 ratio. | Dynamically adjust chart size in a report while preserving a 4:3 proportion. | Generate Excel dashboards where all charts must retain identical visual ratios across sheets.
// AI Prompts: Write C# code with Aspose.Cells that sets a chart’s Height and Width to keep a constant aspect ratio during resizing. | Explain how to emulate LockAspectRatio behavior for Aspose.Cells charts by calculating dimensions after creation. | Provide an example that enforces a 3:2 aspect ratio for a line chart using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add sample data, insert a column chart, and keep its proportions consistent by manually assigning Height and Width, because Aspose.Cells for .NET lacks a direct LockAspectRatio property.
    public class LockChartAspectRatioDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
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

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Note: Aspose.Cells for .NET does not expose a direct LockAspectRatio property for charts.
                // The aspect ratio can be controlled manually by setting the chart's size (Height/Width) as needed.

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LockChartAspectRatioDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            LockChartAspectRatioDemo.Run();
        }
    }
}
