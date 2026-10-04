// Title: Export Excel worksheet charts to PNG files named after each worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, loops through each worksheet, and saves every chart as a PNG file whose name is derived from the worksheet title, adding an index when a sheet has more than one chart. | Show a robust C# example that checks for the source .xlsx file, catches errors, and uses Aspose.Cells to generate PNG images for all charts while creating unique filenames per worksheet.
// Common Searches: C# Aspose.Cells save each chart from an Excel file as a separate PNG using worksheet names | How to export multiple charts on one sheet to PNG with distinct filenames in .NET | Aspose.Cells generate PNG images for all charts in a workbook programmatically | Create unique image files for Excel charts based on sheet name in C# | Iterate over worksheets and export charts to PNG with Aspose.Cells example
// Tags: Aspose.Cells export chart to PNG | chart image naming by worksheet | C# iterate workbook charts | multiple chart PNG export Aspose.Cells | dynamic filename generation for Excel chart images

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample verifies that 'input.xlsx' exists, loads it with Aspose.Cells, iterates through every worksheet and each chart on those sheets, and uses Chart.ToImage to export each chart as a PNG file named after its worksheet (appending an index when a sheet contains several charts). It logs each export and handles exceptions gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIndex = 0; // counter for multiple charts on the same sheet

                // Iterate through all charts on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Build the PNG file name using the worksheet name (and index if needed)
                    string fileName = chartIndex == 0
                        ? $"{sheet.Name}.png"
                        : $"{sheet.Name}_{chartIndex}.png";

                    // Export the chart to a PNG image
                    chart.ToImage(fileName);

                    Console.WriteLine($"Chart exported to '{fileName}'.");
                    chartIndex++;
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
