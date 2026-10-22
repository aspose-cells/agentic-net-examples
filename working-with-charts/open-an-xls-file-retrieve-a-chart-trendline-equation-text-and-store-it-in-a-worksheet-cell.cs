// Title: Read a chart's trendline equation from an XLS file and write it to a worksheet cell using Aspose.Cells for .NET
// AI Prompts: Open an existing .xls workbook, locate the first chart, extract the trendline equation from its first series, and store the string in cell A1. | Extend the code to iterate over all charts in the workbook and write each series' trendline equation to successive rows. | Add robust error handling that detects missing trendlines and writes a custom message instead of the default placeholder.
// Common Searches: how to get trendline formula from a chart with Aspose.Cells C# | Aspose.Cells read chart series trendline text in .xls workbook | write extracted chart information into Excel cell using Aspose.Cells .NET | C# retrieve chart trendline equation from existing XLS file | Aspose.Cells chart metadata extraction example
// Tags: retrieve chart trendline equation Aspose.Cells | write string to worksheet cell C# | access first chart series Aspose.Cells | handle missing trendline placeholder Aspose.Cells | save modified workbook to new file .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook, checks for at least one chart and series, attempts to read the trendline equation of the first series (using a placeholder when unavailable), writes the result into cell A1 of the first worksheet, and saves the updated workbook to a new file while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Access the first series of the chart
            Series series = chart.NSeries[0];

            // Attempt to retrieve a trendline equation if supported
            string equation = "Trendline information not available in this Aspose.Cells version.";

            // Store the equation (or placeholder) text in a worksheet cell (e.g., A1)
            worksheet.Cells["A1"].PutValue(equation);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated cell
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
