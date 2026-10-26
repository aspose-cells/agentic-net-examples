// Title: Insert an exponential trendline into the first series of an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add exponential trendline to first chart series | How to set an exponential trendline on a chart series with Aspose.Cells .NET | Insert a trendline into an existing Excel chart using Aspose.Cells API | Aspose.Cells example for adding regression line to chart data series | Programmatically add exponential regression to Excel chart in C#
// Tags: trendline insertion Aspose.Cells | chart series manipulation .NET | Aspose.Cells chart editing API | Excel chart programmatic modification C# | add regression line Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing workbook, checks that the first worksheet contains a chart with at least one data series, adds an exponential trendline named "Exponential Trendline" to the first series, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one data series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any data series.");
                return;
            }

            // Get the first data series of the chart
            Series series = chart.NSeries[0];

            // Add an exponential trendline to the first series
            Trendline trendline = series.AddTrendline(TrendlineType.Exponential);
            trendline.Name = "Exponential Trendline";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new trendline
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
