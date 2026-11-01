// Title: C# – Load an Excel workbook with Aspose.Cells, add a linear trendline to the first chart series, retrieve its equation, and print details to the console
// AI Prompts: Generate C# code that opens a .xlsx file using Aspose.Cells, selects the first worksheet, adds a linear trendline to the first series of the first chart, and writes the trendline type to the console. | Write C# that, after adding a linear trendline with Aspose.Cells, reads the trendline's equation formula and displays it together with the trendline type. | Create robust C# that checks for the existence of charts and series before adding a trendline, and logs clear messages when they are missing.
// Common Searches: aspnet add linear trendline to chart series using Aspose.Cells | how to get trendline equation from Excel chart in C# Aspose.Cells | check for charts in worksheet before adding trendline Aspose.Cells .NET | retrieve trendline type and formula from Excel chart with Aspose.Cells | error handling when workbook has no charts Aspose.Cells C#
// Tags: add linear trendline to chart series Aspose.Cells | extract trendline equation Aspose.Cells C# | check for charts before trendline Aspose.Cells | load workbook and modify chart Aspose.Cells | error handling missing series Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook using Aspose.Cells, verifies that the first worksheet contains a chart and a series, adds a linear trendline to that series, retrieves the trendline's equation formula, and prints both the trendline type and formula to the console while handling missing elements and potential errors.
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
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Ensure the chart contains at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("No series found in the chart.");
                return;
            }

            // Use dynamic to access Trendlines (avoids compile‑time binding issues)
            dynamic series = chart.NSeries[0];

            try
            {
                // Add a linear trendline to the series
                int trendlineIndex = series.Trendlines.Add(TrendlineType.Linear);

                // Retrieve the newly added trendline
                dynamic trendline = series.Trendlines[trendlineIndex];

                // Output basic information about the trendline
                Console.WriteLine($"Linear Trendline added. Type: {trendline.Type}");
            }
            catch (Exception ex)
            {
                // Handle errors related to trendline operations
                Console.WriteLine($"Trendline operation failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
