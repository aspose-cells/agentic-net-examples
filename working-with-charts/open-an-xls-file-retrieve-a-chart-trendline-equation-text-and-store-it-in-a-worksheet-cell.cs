// Title: Read a chart trendline equation from an XLS file and write it to a cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get trendline equation from existing XLS chart | How to write chart trendline formula into Excel cell with Aspose.Cells | Extract first series trendline equation and save to A1 using Aspose.Cells .NET | Aspose.Cells read trendline equation from .xls workbook programmatically | Store chart trendline text in worksheet cell with Aspose.Cells C# example
// Tags: Aspose.Cells chart trendline extraction | Aspose.Cells write text to Excel cell | Aspose.Cells XLS workbook trendline handling | Aspose.Cells C# chart series API | Aspose.Cells modify worksheet after reading chart data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook, verifies that a chart and its first series contain a trendline, retrieves the TrendlineEquation string, writes that string into cell A1 of the same worksheet, and saves the updated file, with error handling for missing elements.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart contains at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("No series found in the chart.");
                return;
            }

            // Use dynamic to access Trendlines (avoids compile‑time binding issues)
            dynamic series = chart.NSeries[0];

            // Ensure the series contains at least one trendline
            if (series.Trendlines == null || series.Trendlines.Count == 0)
            {
                Console.WriteLine("No trendlines found in the series.");
                return;
            }

            dynamic trendline = series.Trendlines[0];

            // Retrieve the trendline equation text
            string equation = trendline.TrendlineEquation as string ?? string.Empty;

            // Store the equation text in cell A1 of the same worksheet
            worksheet.Cells["A1"].PutValue(equation);

            // Save the workbook with the updated cell
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
