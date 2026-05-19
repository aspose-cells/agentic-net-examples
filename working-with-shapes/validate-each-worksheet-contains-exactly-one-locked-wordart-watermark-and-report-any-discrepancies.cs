using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermarkValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be validated
            string inputPath = "input.xlsx";

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

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int lockedWordArtCount = 0;

                    // Examine each shape in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Check if the shape is a WordArt (TextEffect) and is locked
                        // Use string comparison to avoid reliance on ShapeType enum availability
                        if (shape.Type.ToString() == "TextEffect" && shape.IsLocked)
                        {
                            lockedWordArtCount++;
                        }
                    }

                    // Report discrepancy if the count is not exactly one
                    if (lockedWordArtCount != 1)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" has {lockedWordArtCount} locked WordArt watermark(s). Expected exactly 1.");
                    }
                    else
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains exactly one locked WordArt watermark.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}