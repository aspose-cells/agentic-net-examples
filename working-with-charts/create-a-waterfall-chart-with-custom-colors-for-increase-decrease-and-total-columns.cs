// Title: Create a Waterfall chart in C# with Aspose.Cells and apply distinct colors for increase, decrease, and total columns
// AI Prompts: Write C# code that uses Aspose.Cells to build a waterfall chart and set green for positive bars, red for negative bars, and blue for the total bar. | Show how to iterate over chart series points in Aspose.Cells and assign a foreground color based on the underlying cell value. | Provide a complete, runnable example that creates an Excel workbook, populates waterfall data, adds the chart, customizes point colors, and saves the file.
// Common Searches: aspnet c# assign green red blue colors to waterfall chart points using Aspose.Cells | example of coloring increase and decrease columns differently in Aspose.Cells waterfall chart | how to highlight total column with a distinct color in Aspose.Cells waterfall chart C#
// Tags: Aspose.Cells waterfall chart point foreground colors | C# set series point color Aspose.Cells | increase decrease bar colors Aspose.Cells chart | total column distinct color Aspose.Cells | generate Excel waterfall chart Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills cells with category and numeric data for a waterfall chart, adds a Waterfall chart, iterates over each data point to apply green to positive values, red to negative values, and blue to the final total column, then saves the workbook as WaterfallChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet's cells collection
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Populate data for the waterfall chart
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Categories
            cells["A2"].PutValue("Start");
            cells["A3"].PutValue("Revenue");
            cells["A4"].PutValue("Cost");
            cells["A5"].PutValue("Profit");
            cells["A6"].PutValue("Total");

            // Values (positive = increase, negative = decrease, zero placeholders for calculated totals)
            cells["B2"].PutValue(1000);   // Start
            cells["B3"].PutValue(3000);   // Increase
            cells["B4"].PutValue(-1500);  // Decrease
            cells["B5"].PutValue(0);      // Placeholder (will be calculated by the chart)
            cells["B6"].PutValue(0);      // Total placeholder

            // Add a Waterfall chart
            var worksheet = workbook.Worksheets[0];
            int chartIndex = worksheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            var chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Define custom colors
            Color increaseColor = Color.Green;
            Color decreaseColor = Color.Red;
            Color totalColor = Color.Blue;

            // Apply custom colors to each data point based on its type
            var series = chart.NSeries[0];
            for (int i = 0; i < series.Points.Count; i++)
            {
                var point = series.Points[i];
                double value = cells[i + 2, 1].DoubleValue; // B column values (row offset by 2)

                // Last point is treated as the total column
                if (i == series.Points.Count - 1)
                {
                    point.Area.ForegroundColor = totalColor;
                }
                else if (value >= 0)
                {
                    point.Area.ForegroundColor = increaseColor;
                }
                else
                {
                    point.Area.ForegroundColor = decreaseColor;
                }
            }

            // Save the workbook with the chart
            string outputPath = "WaterfallChart.xlsx";

            // Ensure the directory exists before saving (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
