// Title: Add a linear trendline to the first series of the first chart in an XLSX workbook and display its equation using Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX workbook, locate the first chart, add a linear type trendline to its first series, enable equation display, and output the equation with Aspose.Cells in C#. | Use Aspose.Cells to programmatically obtain and display the formula of a chart series trendline in a .NET workbook.
// Common Searches: c# insert linear trendline into Excel chart using Aspose.Cells | aspnet retrieve equation of a trendline from an existing XLSX chart | display trendline equation in Aspose.Cells chart series example | apply trendline to first chart series in workbook with Aspose.Cells .NET | dynamic call Trendlines.Add Aspose.Cells version compatibility
// Tags: linear trendline insertion Aspose.Cells chart | display trendline formula C# | first chart series access Aspose.Cells | dynamic Trendlines.Add call | modify existing XLSX chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing XLSX file, checks for at least one worksheet, chart, and series, then uses a dynamic call to add a linear trendline to the first series of the first chart, enables equation display, and writes status messages to the console while handling missing elements and potential version‑specific limitations.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("No series found in the chart.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            try
            {
                // Use dynamic to call Trendlines.Add, avoiding compile‑time binding issues
                dynamic dynSeries = series;
                Trendline trendline = dynSeries.Trendlines.Add(TrendlineType.Linear);
                trendline.DisplayEquation = true;
                Console.WriteLine("Trendline added to the series.");
            }
            catch (Exception ex)
            {
                // Trendline feature may not be supported in older Aspose.Cells versions
                Console.WriteLine("Unable to add trendline: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
