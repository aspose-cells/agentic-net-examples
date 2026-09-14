// Title: Remove all trendlines from an Excel chart using Aspose.Cells for .NET before adding new analytical lines
// AI Prompts: Write C# code with Aspose.Cells that iterates through every series in a chart and deletes all existing trendlines. | Show how to clear trendlines from a chart's NSeries and then insert a new linear trendline using Aspose.Cells. | Provide a method that removes chart series trendlines, adds a fresh trendline, and saves the workbook to a new file.
// Common Searches: aspocells how to delete all trendlines from a chart in C# | remove existing trendlines before adding new ones with Aspose.Cells .NET | clear chart series trendlines programmatically using Aspose.Cells | C# Aspose.Cells example to reset trendlines on an Excel chart | delete trendlines from Excel chart and add linear trendline Aspose.Cells
// Tags: Aspose.Cells remove chart trendlines C# | clear Excel chart trendlines programmatically | add linear trendline after clearing Aspose.Cells | chart NSeries trendline manipulation Aspose.Cells | reset chart trendlines .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, accesses the first worksheet and its first chart, iterates through each series to remove all trendlines, optionally adds a new linear trendline to the first series, and saves the modified workbook.
class RemoveTrendlinesExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (adjust index as needed)
            Chart chart = worksheet.Charts[0];

            // Remove all existing trendlines from every series in the chart
            foreach (dynamic series in chart.NSeries)
            {
                try
                {
                    // Remove trendlines in reverse order to avoid index shifting
                    for (int i = series.Trendlines.Count - 1; i >= 0; i--)
                    {
                        series.Trendlines.RemoveAt(i);
                    }
                }
                catch
                {
                    // If the Series type does not support Trendlines, ignore
                }
            }

            // OPTIONAL: Add a new linear trendline to the first series
            if (chart.NSeries.Count > 0)
            {
                try
                {
                    dynamic firstSeries = chart.NSeries[0];
                    firstSeries.Trendlines.Add(TrendlineType.Linear);
                }
                catch
                {
                    // If Trendlines are not supported, skip adding
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
