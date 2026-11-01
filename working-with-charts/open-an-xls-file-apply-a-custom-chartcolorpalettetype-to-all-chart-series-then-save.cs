// Title: Apply a custom color palette to all chart series in an XLS workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing .xls workbook with Aspose.Cells, define a Color[] palette, iterate over every worksheet and its charts, assign each series' Area.ForegroundColor from the palette, and save the modified file. | Write a helper method that accepts input and output file paths plus a Color[] array, opens the workbook, and updates the foreground fill color of all chart series using Aspose.Cells. | Extend the sample to also set each series' marker fill color from the same palette and output a log of the charts and series that were processed.
// Common Searches: Aspose.Cells set series foreground color for charts in existing XLS file C# | apply custom color palette to all chart series in a workbook using Aspose.Cells .NET | change chart series colors programmatically in an XLS workbook Aspose.Cells | iterate over worksheets and charts to update series colors with Aspose.Cells C# | save modified XLS workbook after updating chart colors using Aspose.Cells
// Tags: apply custom series colors Aspose.Cells | set chart series foreground fill .NET | iterate worksheets and charts Aspose.Cells | custom color palette XLS workbook | update chart colors programmatically C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an existing XLS workbook, defines a custom Color[] palette, walks through each worksheet and its charts, assigns each series' foreground fill color from the palette, ensures the output directory exists, and saves the updated workbook to a new file.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS file
            Workbook workbook = new Workbook(inputPath);

            // Define a custom color palette (example colors)
            Color[] customPalette = new Color[]
            {
                Color.FromArgb(255, 0, 0),      // Red
                Color.FromArgb(0, 255, 0),      // Green
                Color.FromArgb(0, 0, 255),      // Blue
                Color.FromArgb(255, 165, 0),    // Orange
                Color.FromArgb(128, 0, 128)     // Purple
            };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Apply custom colors to each series in the chart
                        int colorIdx = 0;
                        // NSeries holds the collection of series for the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Cycle through the custom palette
                            Color clr = customPalette[colorIdx % customPalette.Length];
                            // Set the series fill (foreground) color
                            series.Area.ForegroundColor = clr;
                            colorIdx++;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart '{chart.Name}': {exChart.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
