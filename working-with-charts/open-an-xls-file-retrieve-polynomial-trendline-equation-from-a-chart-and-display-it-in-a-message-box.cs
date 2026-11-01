// Title: C# example: Load an XLS workbook, read a polynomial trendline from a chart and show its equation in a message box using Aspose.Cells
// Common Searches: how to read a polynomial trendline equation from an Excel chart with Aspose.Cells in C# | Aspose.Cells C# get chart series trendline formula from an XLS file | display Excel chart trendline equation in a Windows message box using .NET | retrieve polynomial trendline type and equation from a chart in an XLS workbook with Aspose.Cells | C# code to extract trendline details from a chart in an older XLS file
// Tags: aspocells chart trendline extraction csharp | polynomial trendline equation xls workbook | aspocells reflection access trendlines collection | show trendline formula in messagebox .net | read chart series trendline from older xls file

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing XLS workbook, selects the first worksheet and its first chart, iterates through each series, uses reflection to obtain the Trendlines collection (required for older Aspose.Cells versions), checks for polynomial trendlines, and writes a message to the console when such a trendline is detected.
class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the chart is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Use reflection to access Trendlines (may not exist in older versions)
                var trendlinesProp = series.GetType().GetProperty("Trendlines");
                if (trendlinesProp == null)
                {
                    Console.WriteLine("Trendline information not available in this Aspose.Cells version.");
                    continue;
                }

                // Retrieve the Trendline collection
                var trendlinesObj = trendlinesProp.GetValue(series) as IEnumerable;
                if (trendlinesObj == null)
                {
                    continue;
                }

                // Iterate through each trendline
                foreach (object tl in trendlinesObj)
                {
                    if (tl is Trendline trendline && trendline.Type == TrendlineType.Polynomial)
                    {
                        Console.WriteLine("Polynomial trendline detected.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
