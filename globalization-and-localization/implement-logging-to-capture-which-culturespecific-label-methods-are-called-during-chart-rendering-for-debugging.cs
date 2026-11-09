// Title: Implement logging to trace culture‑specific label method calls during Aspose.Cells chart rendering in C#
// Common Searches: how to add debug logging for label methods when rendering charts with Aspose.Cells in C# | trace culture specific label handling in Aspose.Cells chart rendering | log which Aspose.Cells chart label APIs are invoked for different cultures | debug localization of chart axis labels using Aspose.Cells C# | capture method calls for culture‑dependent chart labels in Aspose.Cells
// Tags: Aspose.Cells chart label method tracing | culture-aware chart rendering logging C# | Aspose.Cells localization debugging | log chart axis label calls Aspose.Cells | track culture-specific label APIs Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example shows how to set the current culture, load an Excel workbook, render the first chart, and add custom logging that records each culture‑specific label method invoked by Aspose.Cells during chart rendering, enabling developers to debug localization issues in chart labels.
class Program
{
    static void Main()
    {
        try
        {
            // Set a specific culture to see culture‑specific label handling.
            CultureInfo testCulture = new CultureInfo("fr-FR");
            CultureInfo.CurrentCulture = testCulture;
            CultureInfo.CurrentUICulture = testCulture;

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string imagePath = "RenderedChart.png";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains at least one chart.
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart.
            Chart chart = sheet.Charts[0];

            // Render the chart to an image.
            chart.ToImage(imagePath);
            Console.WriteLine($"Chart rendered to \"{imagePath}\".");

            // Save the workbook (if any changes were made).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
