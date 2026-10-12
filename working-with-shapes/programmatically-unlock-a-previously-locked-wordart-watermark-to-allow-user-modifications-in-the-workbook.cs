// Title: Unlock WordArt watermarks in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that opens a .xlsx file, iterates through all shapes on a specified worksheet, checks the IsWordArt flag, and sets IsLocked to false for each WordArt shape. | Adapt the sample to receive the worksheet name as an argument and unlock WordArt shapes only on that sheet while leaving other sheets untouched. | Enhance the example with robust try‑catch logic that logs a warning when a shape does not support the IsWordArt property and then continues processing the remaining shapes.
// Common Searches: C# Aspose.Cells unlock WordArt shape in Excel worksheet | how to remove WordArt watermark protection using Aspose.Cells .NET | iterate over Excel shapes and change IsLocked property with Aspose.Cells | unlock WordArt objects in a specific sheet programmatically in .xlsx | Aspose.Cells example for editing locked WordArt in Excel file
// Tags: unlock WordArt shape Aspose.Cells | set shape IsLocked false Excel | iterate worksheet shapes C# | handle IsWordArt exception Aspose.Cells | modify Excel watermark programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UnlockWordArtExample
{
    // The program loads an Excel workbook, loops through all shapes on the first worksheet, detects WordArt objects via the IsWordArt property, unlocks them by setting IsLocked to false, and saves the updated file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Unlock WordArt shapes (if any) in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    try
                    {
                        // Some Shape objects may not support IsWordArt; guard against exceptions
                        if (shape.IsWordArt)
                        {
                            shape.IsLocked = false;
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        // Log shape-specific errors and continue processing other shapes
                        Console.WriteLine($"Warning: Unable to process a shape. Details: {shapeEx.Message}");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // General error handling
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
