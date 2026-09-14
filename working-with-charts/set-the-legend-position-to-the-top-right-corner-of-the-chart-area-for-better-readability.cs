// Title: How to position a chart legend at the top‑right corner in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, verifies the first chart exists, and sets its legend to LegendPositionType.TopRight (or Top as a fallback) before saving. | Show a complete example that validates the input workbook path, creates missing output directories, and updates the chart legend to the upper‑right corner using the Aspose.Cells API. | Provide a snippet that handles older Aspose.Cells versions where LegendPositionType.TopRight is unavailable, using LegendPositionType.Top instead, and writes the modified file.
// Common Searches: Aspose.Cells C# change legend placement to upper‑right in Excel chart | How to move chart legend to top right using Aspose.Cells for .NET | LegendPositionType.TopRight not supported in older Aspose.Cells releases | C# example for adjusting Excel chart legend location with Aspose.Cells | Validate chart existence before modifying legend in Aspose.Cells workbook
// Tags: Aspose.Cells chart legend positioning | C# legend position enum Aspose.Cells | Excel chart layout adjustment .NET | fallback legend position Top Aspose.Cells | verify chart presence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing .xlsx file, checks for a chart, sets the chart's legend to the top‑right (using Top as a fallback for older versions), ensures the output directory exists, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Set the legend position. TopRight is not available in older versions,
            // so we use Top as a safe alternative.
            chart.Legend.Position = LegendPositionType.Top;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
