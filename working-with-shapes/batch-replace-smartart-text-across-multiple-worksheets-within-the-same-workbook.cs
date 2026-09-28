// Title: Batch replace text in non‑SmartArt shapes across all worksheets of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates every worksheet and each shape, replaces a given string in shape.Text, and saves the updated workbook. | Demonstrate how to use the IsSmartArt property to identify and skip SmartArt shapes while performing a bulk text replacement in Excel shapes. | Add a try‑catch block around each shape update to log errors and continue processing when modifying shape text with Aspose.Cells.
// Common Searches: aspnet replace text in all shapes of an Excel file using Aspose.Cells | how to skip SmartArt when updating shape text in a workbook with C# | batch modify shape.Text property across multiple worksheets Aspose.Cells | C# iterate workbook shapes and replace specific string in Excel
// Tags: bulk shape text replacement Aspose.Cells | skip SmartArt shapes .NET | iterate worksheets shapes Aspose.Cells | shape text find and replace Excel C# | per‑shape error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, loops through each worksheet and its shapes, skips SmartArt shapes, replaces occurrences of "OldText" with "NewText" in shape.Text, logs any shape‑level exceptions, and saves the modified workbook.
class SmartArtBatchReplace
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Text to find and its replacement
            const string textToFind = "OldText";
            const string replacementText = "NewText";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Aspose.Cells does not provide direct SmartArt manipulation.
                        // If the shape is SmartArt, log and skip it.
                        if (shape.IsSmartArt)
                        {
                            Console.WriteLine($"Skipping SmartArt shape on sheet '{sheet.Name}'.");
                            continue;
                        }

                        // Process regular shapes that contain text
                        if (!string.IsNullOrEmpty(shape.Text) && shape.Text.Contains(textToFind))
                        {
                            shape.Text = shape.Text.Replace(textToFind, replacementText);
                        }
                    }
                    catch (Exception exShape)
                    {
                        // Log shape‑level errors but continue processing other shapes
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {exShape.Message}");
                    }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
