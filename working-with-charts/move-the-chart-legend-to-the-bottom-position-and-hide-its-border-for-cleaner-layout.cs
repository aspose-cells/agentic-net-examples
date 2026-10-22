// Title: How to move an Excel chart legend to the bottom and hide its border using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, sets the first chart's legend position to Bottom, and disables the legend border. | Show a minimal Aspose.Cells example that repositions a chart legend to the bottom and turns off its border visibility. | Provide a step‑by‑step C# snippet to adjust a chart's legend placement and border using the Aspose.Cells Chart.Legend API.
// Common Searches: Aspose.Cells C# move chart legend to bottom | Hide legend border in Excel chart using Aspose.Cells | Change legend position of a chart with Aspose.Cells .NET | Aspose.Cells example for customizing chart legend appearance | C# code to set chart legend position and border visibility in Aspose.Cells
// Tags: set chart legend position bottom Aspose.Cells | disable legend border Aspose.Cells C# | customize Excel chart legend Aspose.Cells | Aspose.Cells chart legend formatting C# | modify chart legend placement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, accesses the first chart, moves its legend to the bottom, hides the legend border, and saves the updated workbook.
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

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Move the legend to the bottom of the chart
            chart.Legend.Position = LegendPositionType.Bottom;

            // Hide the legend border
            chart.Legend.Border.IsVisible = false;

            // Save the workbook with the updated chart
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
