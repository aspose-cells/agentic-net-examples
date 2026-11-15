// Title: Move an Excel chart legend to the bottom and hide its border with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend to bottom position | remove border from chart legend using Aspose.Cells .NET | how to reposition and style Excel chart legend programmatically | C# code to place chart legend at bottom and disable its border | Aspose.Cells example for customizing chart legend layout
// Tags: chart legend bottom placement Aspose.Cells | disable legend border .NET | Excel chart legend formatting Aspose.Cells | modify chart legend layout C# | set legend position enum Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, accesses the first chart, moves its legend to the bottom, hides the legend border, and saves the updated file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Move the legend to the bottom of the chart
            // Use LegendPositionType enum (compatible with various Aspose.Cells versions)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Hide the legend border for a cleaner layout
            chart.Legend.Border.IsVisible = false;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
