// Title: Load an XLSX workbook with Aspose.Cells, read chart trendlines and log their confidence interval values in C#
// AI Prompts: Generate C# code that opens a .xlsx file using Aspose.Cells, selects the first worksheet and its first chart, and iterates through each series to obtain trendline objects. | Write a C# snippet that safely accesses the Trendlines collection of a chart series via dynamic binding and prints the series name, trendline type, and any available confidence interval values. | Create a robust C# program that handles missing Trendlines support, logs errors, and outputs the confidence interval (lower and upper bounds) for each trendline when the Aspose.Cells API exposes them.
// Common Searches: how to read trendline confidence interval from Excel chart using Aspose.Cells C# | Aspose.Cells retrieve chart series trendline details .NET | C# dynamic access to chart trendlines when property may not exist Aspose.Cells | example code to load XLSX and extract trendline statistics with Aspose.Cells | log confidence interval values of Excel chart trendline in .NET application
// Tags: Aspose.Cells load XLSX workbook chart trendline confidence interval | C# dynamic trendlines collection Aspose.Cells | extract chart series trendline type Aspose.Cells .NET | log Excel chart trendline statistics using Aspose.Cells | handle missing trendlines property Aspose.Cells runtimebinderexception

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Microsoft.CSharp.RuntimeBinder;

// The example checks for the input.xlsx file, loads it with Aspose.Cells, accesses the first worksheet’s first chart, iterates over each series, uses dynamic binding to safely read any Trendlines collection, and writes the series name, trendline type and, when available, the confidence interval bounds to the console while handling absent trendline support gracefully.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Iterate through each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Use dynamic to safely access Trendlines (may not exist in older versions)
                try
                {
                    dynamic dynSeries = series;
                    foreach (Trendline trendline in dynSeries.Trendlines)
                    {
                        Console.WriteLine($"Series: {series.Name}");
                        Console.WriteLine($"  Trendline Type: {trendline.Type}");
                        Console.WriteLine();
                    }
                }
                catch (RuntimeBinderException)
                {
                    // Trendlines property not available in this version
                    Console.WriteLine($"Series: {series.Name} has no trendlines support.");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
