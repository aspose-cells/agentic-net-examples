// Title: Validate that all chart legends use the updated Dark2 theme color after changing the theme in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Load a workbook, change the Dark2 theme color to a custom RGB value, then loop through every worksheet and chart to assign the new color to each legend's Font.Color. | After applying the new Dark2 shade, compare each chart legend's Font.Color with the expected value, log any mismatches, and save the workbook.
// Common Searches: c# aspocells change dark2 theme color and update all chart legend fonts | how to confirm chart legend colors reflect a new theme color in Aspose.Cells | iterate over all charts in a workbook to set legend font color using Aspose.Cells .NET
// Tags: update Dark2 theme color Aspose.Cells | set chart legend font color C# | validate chart legend color after theme change | iterate worksheets and charts Aspose.Cells | apply custom theme shade to Excel legends

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, changes the Dark2 theme color to a custom shade, iterates through every worksheet and chart to set each legend's font color to the new shade, verifies that the update succeeded, logs any mismatches, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Define the new shade for the Dark2 theme color (example: DodgerBlue)
                Color newDark2Color = Color.FromArgb(255, 30, 144, 255);

                // Flag to track validation result
                bool allLegendsUpdated = true;

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            // Access the chart legend
                            Legend legend = chart.Legend;

                            // Update legend font color to the new Dark2 color
                            legend.Font.Color = newDark2Color;

                            // Verify the update
                            if (legend.Font.Color != newDark2Color)
                            {
                                allLegendsUpdated = false;
                                Console.WriteLine(
                                    $"Mismatch in chart '{chart.Name}' on sheet '{sheet.Name}'. " +
                                    $"Legend color: {legend.Font.Color}, Expected: {newDark2Color}");
                            }
                        }
                        catch (Exception exChart)
                        {
                            // Handle any chart‑specific errors without stopping the whole process
                            Console.WriteLine($"Error processing chart '{chart.Name}' on sheet '{sheet.Name}': {exChart.Message}");
                            allLegendsUpdated = false;
                        }
                    }
                }

                // Output validation result
                if (allLegendsUpdated)
                {
                    Console.WriteLine("All chart legends reflect the updated Dark2 theme color.");
                }
                else
                {
                    Console.WriteLine("Some chart legends do not reflect the updated Dark2 theme color.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
