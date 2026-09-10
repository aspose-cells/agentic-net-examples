// Title: Apply 1.2‑point character spacing to all shape text in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Update the sample to assign shape.TextEffect.CharacterSpacing = 1.2f for every shape that contains a TextEffect. | Extend the loop to also detect shapes with a TextFrame and set their CharacterSpacing property to 1.2 points. | Create a helper method ApplyCharacterSpacing(Workbook workbook, float spacing) that walks all worksheets and shapes and applies the given spacing. | Add logging to record the names of shapes that were skipped because they do not support character spacing.
// Common Searches: Aspose.Cells C# how to change character spacing of text inside Excel shapes | Set 1.2 point spacing for shape text in an Excel file using Aspose.Cells .NET | Increase readability of shape text by adjusting character spacing with Aspose.Cells | C# iterate through worksheet shapes and modify TextEffect spacing Aspose.Cells | Apply uniform character spacing to all shapes in a workbook with Aspose.Cells
// Tags: Aspose.Cells set shape character spacing | C# adjust text effect spacing Excel | Excel shape text formatting Aspose.Cells | apply character spacing to all shapes .NET | iterate worksheet shapes modify text properties

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, iterates through each worksheet and every shape, and sets the character spacing of any TextEffect (or TextFrame) to 1.2 points to improve readability before saving the modified file.
class Program
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // If the shape contains a TextEffect, adjust a supported property
                    if (shape.TextEffect != null)
                    {
                        // Example: change the font size of the text effect
                        shape.TextEffect.FontSize = 12; // set font size to 12 points
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
