// Title: Render an Excel chart to a PNG file using Aspose.Cells Chart.ToImage in C#
// Common Searches: C# Aspose.Cells export chart to PNG file | How to use Chart.ToImage to save an Excel chart as an image in .NET | Saving the first worksheet chart as PNG with Aspose.Cells library | Render an Excel chart to a bitmap image using Aspose.Cells C# example
// Tags: Aspose.Cells chart to PNG conversion | Chart.ToImage method C# | export Excel chart as image .NET | render worksheet chart to bitmap using Aspose | save chart as PNG with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, checks for a chart on the first worksheet, creates the output folder if needed, and uses the Chart.ToImage method to render the chart directly to a PNG file, handling any errors that may occur.
class ChartToPngExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "chart.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (change index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart in the worksheet (change index if needed)
            Chart chart = sheet.Charts[0];

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the chart directly to a PNG file
            chart.ToImage(outputPath);
            Console.WriteLine($"Chart rendered to \"{outputPath}\" successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
