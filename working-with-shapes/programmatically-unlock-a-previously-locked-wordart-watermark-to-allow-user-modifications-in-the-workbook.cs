// Title: Programmatically unlock a locked WordArt watermark shape in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to unlock a locked WordArt watermark in an existing Excel file | how to remove shape lock from WordArt watermark with Aspose.Cells .NET | unlocking protected shapes in Excel workbook using Aspose.Cells API | set IsLocked false for WordArt shapes in Aspose.Cells C# example | modify Excel watermark shape after loading workbook with Aspose.Cells
// Tags: unlock WordArt shape Aspose.Cells | remove shape lock Excel .NET | modify workbook watermark programmatically | iterate worksheet shapes C# Aspose.Cells | set shape IsLocked false Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing Excel file, iterates through each worksheet and its shapes, checks the IsLocked property of every shape (including WordArt watermarks), sets it to false to unlock the shape, and saves the workbook to a new file while handling file‑not‑found and shape‑processing errors.
class UnlockWatermark
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_unlocked.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Unlock any locked shape (including WordArt) to allow modifications
                        if (shape.IsLocked)
                        {
                            shape.IsLocked = false;
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Failed to process a shape on sheet \"{sheet.Name}\": {exShape.Message}");
                    }
                }
            }

            // Save the workbook with the unlocked watermark
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
