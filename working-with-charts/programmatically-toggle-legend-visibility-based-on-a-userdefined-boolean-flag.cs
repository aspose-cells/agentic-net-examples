// Title: Toggle Chart Legend Visibility in Aspose.Cells for .NET (C#)
// Description: A concise C# example that creates a workbook, adds a column chart, and uses the Chart.ShowLegend property to turn the legend on or off based on a Boolean parameter before saving the file.
// Keywords: Aspose.Cells C# chart legend | Chart.ShowLegend property | toggle legend Aspose.Cells | hide Excel chart legend .NET | programmatic chart customization | Excel chart generation Aspose | dynamic legend visibility
// Common Searches: how to hide legend in Aspose.Cells chart C# | set chart legend visibility programmatically Aspose.Cells | toggle Excel chart legend using .NET | Chart.ShowLegend example Aspose | conditional legend display Aspose.Cells
// Developer Intent: Enable or disable a chart's legend in an Excel workbook by passing a Boolean flag to the code.
// Use Cases: Display the legend only when a chart contains multiple series to keep the layout clean. | Suppress the legend for single‑series reports to conserve space. | Provide an end‑user setting that remembers their preference for legend visibility in generated Excel files.
// AI Prompts: Generate a C# method that accepts a bool and applies it to chart.ShowLegend, then saves the workbook with Aspose.Cells. | Add comprehensive error handling to the legend‑toggle sample while preserving performance. | Show how to read a configuration value (e.g., from appsettings.json) and use it to control Chart.ShowLegend in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // A concise C# example that creates a workbook, adds a column chart, and uses the Chart.ShowLegend property to turn the legend on or off based on a Boolean parameter before saving the file.
    public class ToggleLegendVisibilityDemo
    {
        // Demonstrates toggling the legend visibility of a chart
        public static void Run(bool showLegend)
        {
            try
            {
                // Create a new workbook
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Toggle legend visibility
                chart.ShowLegend = showLegend;
                Console.WriteLine("Legend visibility set to: " + chart.ShowLegend);

                // Save the workbook
                string outputPath = "ToggleLegendVisibility.xlsx";
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
            // Example usage: show the legend
            ToggleLegendVisibilityDemo.Run(true);
        }
    }
}
