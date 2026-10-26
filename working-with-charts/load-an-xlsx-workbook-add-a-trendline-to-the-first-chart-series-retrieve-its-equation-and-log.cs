// Title: Add a linear trendline to the first series of an Excel chart and retrieve its equation using Aspose.Cells for .NET (C#)
// Common Searches: c# add linear trendline to first chart series using Aspose.Cells | how to get trendline equation from an Excel chart with Aspose.Cells .NET | display trendline equation and R-squared on a chart programmatically Aspose.Cells | using dynamic API to add trendlines in Aspose.Cells C# example | save workbook after modifying chart trendline in C# Aspose.Cells
// Tags: Aspose.Cells chart series trendline addition | Aspose.Cells trendline equation extraction | Aspose.Cells XLSX chart editing | dynamic API usage Aspose.Cells trendlines | saving workbook after chart update Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing XLSX workbook, checks for a chart and its first series, then adds a linear trendline via the dynamic Trendlines API, enables equation and R² display, reads the generated trendline equation, writes it to the console, and finally saves the modified workbook as a new file while handling potential errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            try
            {
                // Use dynamic to access Trendlines API (may vary between versions)
                dynamic dynSeries = series;
                dynamic trendline = dynSeries.Trendlines.Add(TrendlineType.Linear);
                trendline.DisplayEquation = true;   // Show the equation on the chart
                trendline.DisplayRSquared = true;   // Optionally show R² value

                // Retrieve and display the trendline equation (if available)
                string equation = trendline.TrendlineEquation;
                Console.WriteLine($"Trendline equation: {equation}");
            }
            catch (Exception ex)
            {
                // Trendline feature may not be supported in the current library version
                Console.WriteLine($"Trendline could not be added: {ex.Message}");
            }

            // Save the modified workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
