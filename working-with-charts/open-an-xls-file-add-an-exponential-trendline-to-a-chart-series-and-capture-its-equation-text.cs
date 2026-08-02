// Title: Add an Exponential Trendline to an XLS Chart and Display Its Equation with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLS workbook, accesses the first worksheet and its first chart, adds an exponential trendline to the first series, enables the equation display, and saves the updated file. Includes error handling for missing files or charts.
// Keywords: Aspose.Cells exponential trendline C# | add trendline to XLS chart | display trendline equation Aspose | modify chart series .xls | Aspose.Cells chart automation | C# Excel trendline example
// Common Searches: how to add exponential trendline to an XLS chart using Aspose.Cells | Aspose.Cells C# show trendline equation on chart | retrieve trendline equation text from Excel chart with Aspose | add trendline to existing workbook Aspose.Cells | C# code for chart trendline in XLS file
// Developer Intent: Programmatically insert an exponential trendline into the first series of a chart in an existing XLS workbook and make the equation visible on the chart.
// Use Cases: Forecast sales growth by adding an exponential trendline to a chart in a legacy XLS report. | Generate financial statements where the trendline equation must appear for audit documentation. | Automate template updates for engineering data by inserting predictive trendlines into charts before distribution.
// AI Prompts: Write C# code with Aspose.Cells that opens an XLS file, adds an exponential trendline to a specified chart series, sets DisplayEquation = true, and returns the equation string if available. | Explain how to detect whether the TrendlineEquation property exists in the current Aspose.Cells version and provide a fallback to display the equation on the chart. | Create a robust C# example that adds exponential trendlines to all series of the first chart in a worksheet, handles missing files or charts, and saves the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLS workbook, accesses the first worksheet and its first chart, adds an exponential trendline to the first series, enables the equation display, and saves the updated file. Includes error handling for missing files or charts.
class AddExponentialTrendline
{
    static void Main()
    {
        try
        {
            // Verify input file exists
            string inputPath = "input.xls";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Add an exponential trendline to the first series
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

            // Display the equation on the chart
            trendline.DisplayEquation = true;

            // Since TrendlineEquation property may not be available in all versions,
            // we simply acknowledge that the equation will be shown on the chart.
            Console.WriteLine("Exponential trendline added and its equation will be displayed on the chart.");

            // Save the modified workbook
            string outputPath = "output.xls";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
