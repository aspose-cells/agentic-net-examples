// Title: Read and log a chart trendline equation from an XLSX workbook after recalculating formulas with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file using Aspose.Cells, call CalculateFormula, then extract the first chart's first series trendline equation and output it. | Extend the code to iterate through all worksheets, charts, and series, gathering each trendline equation into a collection. | Add version‑check logic that detects if the Trendline property is unavailable and writes a clear warning to the console.
// Common Searches: aspnet read chart trendline equation after workbook.CalculateFormula | c# Aspose.Cells get trendline equation from chart | how to display trendline equation programmatically with Aspose.Cells | iterate over all charts in an XLSX and extract trendline formulas using Aspose.Cells | check Aspose.Cells version for trendline support in .NET
// Tags: chart trendline equation extraction Aspose.Cells | recalculate formulas before reading chart data | enumerate worksheets and charts for trendline equations | fallback handling for unsupported trendline API | log trendline equation from first series C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an XLSX workbook with Aspose.Cells, recalculates all formulas to refresh chart data, accesses the first worksheet’s first chart, obtains the first series’ trendline, forces the equation to be displayed, prints the equation text, and includes error handling for missing files or unsupported Trendline API.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas to ensure chart data is up‑to‑date
            workbook.CalculateFormula();

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Use dynamic to access Trendline members (avoids compile‑time errors if API is unavailable)
            dynamic series = chart.NSeries[0];

            try
            {
                // Attempt to retrieve the Trendlines collection
                var trendlines = series.Trendlines;

                // Ensure the series has at least one trendline
                if (trendlines == null || trendlines.Count == 0)
                {
                    Console.WriteLine("The series does not contain any trendlines.");
                    return;
                }

                // Get the first trendline
                dynamic trendline = trendlines[0];

                // Ensure the equation is displayed (required to retrieve it)
                trendline.DisplayEquation = true;

                // Retrieve and display the equation text
                string equationText = trendline.TrendlineEquation;
                Console.WriteLine("Trendline Equation: " + equationText);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // Trendline API not available in the referenced Aspose.Cells version
                Console.WriteLine("Trendline functionality is not supported by the current Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
