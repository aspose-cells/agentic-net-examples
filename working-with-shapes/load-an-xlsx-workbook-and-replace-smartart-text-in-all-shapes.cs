// Title: Replace text in all SmartArt and other text‑containing shapes of an XLSX workbook using Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file with Aspose.Cells, loop through every worksheet and each Shape, set the Shape.Text property to a specified string, and save the workbook to a new file. | Create a reusable method that takes inputPath, outputPath, and replacementText, then updates the Text property of every shape that contains text (including SmartArt) in the workbook. | Modify the code to filter only SmartArt shapes before replacing their text, leaving other shape types unchanged.
// Common Searches: aspnet replace smartart text in existing excel workbook using aspose.cells | c# iterate over worksheet shapes and modify text with Aspose.Cells | how to bulk update shape text in an xlsx file programmatically | asp.net read xlsx and change all smartart labels with Aspose.Cells | replace all shape text in excel using Aspose.Cells C#
// Tags: replace shape text Aspose.Cells | smartart text modification .net | iterate worksheet shapes C# | bulk edit excel shapes Aspose | update xlsx shape content programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX workbook, iterates through each worksheet and every shape, replaces any non‑empty Text property with a new string, logs shape‑level errors, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Replace text in shapes that support text
                        if (!string.IsNullOrEmpty(shape.Text))
                        {
                            shape.Text = "New Text";
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {exShape.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
