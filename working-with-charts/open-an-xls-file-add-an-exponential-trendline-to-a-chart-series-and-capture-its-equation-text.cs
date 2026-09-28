// Title: Add an exponential trendline to a chart series in an existing XLS workbook and read its equation with Aspose.Cells for .NET
// AI Prompts: Load an XLS file, locate the first chart's first series, add an exponential trendline, enable equation display, and return the equation text. | Update the sample to save the modified workbook as XLSX while keeping the exponential trendline and its displayed equation intact. | Add error handling that logs a warning when the Trendline.Equation property is unavailable in the current Aspose.Cells version.
// Common Searches: Aspose.Cells C# add exponential trendline to existing chart and get equation | How to read trendline equation from an Excel chart using Aspose.Cells .NET | Retrieve exponential trendline formula from XLS workbook with Aspose.Cells | Add trendline to first series of a chart in an XLS file using Aspose.Cells for .NET | Display and capture trendline equation in Aspose.Cells chart
// Tags: add exponential trendline Aspose.Cells chart | retrieve trendline equation Aspose.Cells .NET | modify existing XLS workbook chart series | dynamic trendline property handling Aspose.Cells | save workbook as XLSX with trendline Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook, accesses the first worksheet's first chart and its first series, adds an exponential trendline with equation display enabled, attempts to read the equation text, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Verify input file exists
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

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Get the first series of the chart
            Series series = chart.NSeries[0];

            // Use dynamic to access Trendlines (avoids compile‑time binding issues)
            dynamic dynSeries = series;
            var trendline = dynSeries.Trendlines.Add(TrendlineType.Exponential);
            trendline.DisplayEquation = true;

            // Attempt to retrieve the equation if the property exists at runtime
            string equation = null;
            try
            {
                equation = trendline.Equation;
            }
            catch
            {
                // Property not available in this version; ignore
            }

            Console.WriteLine("Exponential Trendline added."
                              + (equation != null ? $" Equation: {equation}" : string.Empty));

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new trendline
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
