// Title: How to move all chart legends to the bottom of each worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates over every worksheet and chart, sets each chart's Legend.Position to Bottom, and saves the updated file. | Write a .NET console program that opens a specified .xlsx file, changes the legend placement of all charts to the bottom to avoid overlap using Aspose.Cells, and writes the result to a new location. | Create a C# script that verifies an input Excel file exists, updates every chart legend to the bottom with Aspose.Cells, handles any chart‑specific errors, and saves the modified workbook.
// Common Searches: Aspose.Cells C# set chart legend position to bottom for all charts in workbook | bulk update Excel chart legends using Aspose.Cells .NET | C# iterate through worksheets and change chart legend placement with Aspose.Cells | prevent chart legend overlap in Excel by moving legends to bottom via Aspose.Cells
// Tags: Aspose.Cells set chart legend position | C# bulk modify Excel chart legends | Aspose.Cells iterate worksheets charts | prevent legend overlap Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an input .xlsx file, loops through each worksheet and its charts, sets every chart's legend to the bottom to avoid overlapping data, ensures the output directory exists, saves the workbook to a new file, and includes error handling for missing files and chart updates.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and its charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Position the legend at the bottom of the chart
                        chart.Legend.Position = LegendPositionType.Bottom;
                        // No need for IsAutomaticPosition; setting Position is sufficient
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not modify legend for chart \"{chart.Name}\". {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
