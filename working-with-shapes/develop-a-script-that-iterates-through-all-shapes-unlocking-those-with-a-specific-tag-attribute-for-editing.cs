// Title: Unlock Excel shapes with a specific AlternativeText tag using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds shapes whose AlternativeText equals a given tag, and clears the IsLocked flag. | Add logging to the shape‑unlocking script so that each unlocked shape’s name and cell address are written to the console. | Refactor the example to read a user‑defined document property and compare it with each shape’s AlternativeText before clearing IsLocked.
// Common Searches: C# Aspose.Cells how to unlock shapes based on AlternativeText tag | iterate over all shapes in an Excel workbook and change IsLocked property using Aspose.Cells | filter shapes by custom tag and unlock them in .NET | Aspose.Cells unlock specific shapes in multiple worksheets programmatically
// Tags: unlock shape IsLocked Aspose.Cells | filter shapes by AlternativeText Aspose.Cells | iterate worksheet shapes C# | Aspose.Cells shape editing by tag | batch unlock Excel shapes .NET

using System;
using System.IO;
using Aspose.Cells;

namespace UnlockShapesByTag
{
    // The program loads an input workbook, walks through each worksheet and its collection of shapes, checks each shape's AlternativeText against a target tag, sets IsLocked to false for matches, and saves the result to a new file while handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string targetTag = "EditableShape";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the current worksheet
                    foreach (Aspose.Cells.Drawing.Shape shape in sheet.Shapes)
                    {
                        // Use AlternativeText as a tag equivalent
                        if (!string.IsNullOrEmpty(shape.AlternativeText) &&
                            shape.AlternativeText.Equals(targetTag, StringComparison.OrdinalIgnoreCase))
                        {
                            // Unlock the shape for editing
                            shape.IsLocked = false;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
