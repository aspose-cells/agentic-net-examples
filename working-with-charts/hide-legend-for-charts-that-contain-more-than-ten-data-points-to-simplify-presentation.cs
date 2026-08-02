// Title: Hide chart legend when a chart has more than 10 data points with Aspose.Cells for .NET
// Description: C# example that loads an Excel workbook, walks through every worksheet and chart, totals the points in all series, disables the chart's ShowLegend property if the count exceeds ten, and saves the updated file. Demonstrates conditional legend control using Aspose.Cells.
// Keywords: Aspose.Cells | C# chart legend hide | conditional legend visibility | count chart data points | ShowLegend property | Excel chart manipulation | large data set chart | automate legend settings
// Common Searches: Aspose.Cells hide legend based on data points | C# count points in Excel chart series | conditional chart legend .NET | iterate charts in workbook Aspose.Cells | remove legend from complex Excel charts programmatically
// Developer Intent: Automatically suppress the legend of any Excel chart that contains more than ten data points when generating or processing workbooks with Aspose.Cells.
// Use Cases: Clean up dashboards by removing legends from densely populated charts. | Apply a uniform rule across all sheets in a report to avoid visual clutter. | Prepare Excel files for presentation where legends are unnecessary for large data sets.
// AI Prompts: Generate C# code using Aspose.Cells that hides a chart legend when the total number of points across all series exceeds a configurable limit. | Explain how to retrieve the point count for each series in an Aspose.Cells chart and use it to set ShowLegend conditionally. | Suggest ways to improve performance when processing many worksheets and charts for legend visibility rules.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendControl
{
    // C# example that loads an Excel workbook, walks through every worksheet and chart, totals the points in all series, disables the chart's ShowLegend property if the count exceeds ten, and saves the updated file. Demonstrates conditional legend control using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputData.xlsx";
            const string outputPath = "OutputData.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        int totalDataPoints = 0;

                        // Sum data points across all series of the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Use series.Points.Count to get the number of points in the series
                            totalDataPoints += series.Points.Count;
                        }

                        // Hide legend if the chart has more than ten data points
                        if (totalDataPoints > 10)
                        {
                            chart.ShowLegend = false;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
