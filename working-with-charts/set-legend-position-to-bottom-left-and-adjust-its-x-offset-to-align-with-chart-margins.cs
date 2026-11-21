// Title: Set an Excel chart legend to the bottom‑left corner and align its X offset using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend to bottom left and adjust X offset | C# code to move Excel chart legend to bottom left with margin alignment using Aspose.Cells | How to change legend position and offset in an Aspose.Cells chart | Set legend position to bottom left in an Excel workbook with Aspose.Cells for .NET | Align chart legend X offset with worksheet margins using Aspose.Cells C#
// Tags: chart legend positioning Aspose.Cells C# | set legend bottom left Aspose.Cells | adjust legend X offset Excel chart | Aspose.Cells chart formatting C# | modify chart legend alignment .NET | Excel workbook legend customization Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, checks for a chart on the first worksheet, sets the chart's legend to the bottom‑left position, optionally adjusts its X offset to match worksheet margins, and saves the updated workbook, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Set the legend position to Bottom (valid enum value)
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
