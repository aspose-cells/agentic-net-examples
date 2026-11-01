// Title: Unlock Excel shapes that have a specific AlternativeText tag using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that scans every worksheet, finds shapes whose AlternativeText equals a given string, and sets their IsLocked property to false. | Create a script that loads an .xlsx file, iterates through all shapes, unlocks those tagged with a custom AlternativeText value, and saves the workbook. | Generate a robust C# example that checks file existence, handles exceptions, and unlocks shapes by tag using Aspose.Cells.
// Common Searches: C# Aspose.Cells unlock shapes by AlternativeText tag | How to programmatically unlock specific shapes in an Excel file using Aspose.Cells | Iterate over worksheet shapes and change IsLocked property based on custom tag in .NET | Aspose.Cells example to filter shapes by AlternativeText and modify lock status | Unlock shapes with a particular tag in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells unlock shape by AlternativeText | iterate worksheet shapes C# Aspose.Cells | set shape IsLocked property Aspose.Cells | filter Excel shapes using custom tag | modify shape lock status .NET | batch unlock shapes in workbook Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The C# program loads 'input.xlsx', verifies the file exists, iterates through every worksheet and each shape, checks if the shape's AlternativeText matches the target tag "UnlockMe", sets the shape's IsLocked property to false for matching shapes, and saves the updated workbook as 'output.xlsx' with exception handling.
class UnlockShapesByTag
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Tag to search for (stored in AlternativeText)
            string targetTag = "UnlockMe";

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.AlternativeText == targetTag)
                    {
                        // Unlock the shape using the correct property
                        shape.IsLocked = false;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
