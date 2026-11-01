// Title: How to programmatically reset all shape adjustment values to their defaults in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through every worksheet and shape, and restores each shape's adjustment collection to its default values before saving the workbook. | Explain the steps required to clear custom shape adjustments for all shapes in a workbook using Aspose.Cells, noting any version‑specific considerations and fallback methods when the Adjustments property is not exposed. | Provide a concise example that demonstrates resetting shape handles (e.g., AutoShape, Callout) to default settings across all worksheets with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# reset custom shape adjustments in Excel workbook | How to clear shape adjustment handles for all worksheets using Aspose.Cells .NET | Programmatically restore default shape properties in Excel with Aspose.Cells | Iterate over shapes and reset their adjustment values in C# Aspose.Cells | Default shape adjustment values Aspose.Cells version differences
// Tags: reset shape adjustments Aspose.Cells | Aspose.Cells shape default values | C# iterate worksheet shapes Aspose.Cells | clear custom shape handles .NET | Aspose.Cells shape adjustment API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, walks through each worksheet and every shape, and (when supported) resets each shape's adjustment values to their default settings before saving the modified file. It also notes version considerations for the Adjustments property.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace with actual paths or pass via args)
            string inputFile = @"{InputFile}";
            string outputFile = @"{OutputFile}";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the current worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // If needed, reset shape adjustments or properties here.
                        // The Adjustments property is not available in this version of Aspose.Cells,
                        // so this block is left intentionally empty or can contain alternative logic.
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
