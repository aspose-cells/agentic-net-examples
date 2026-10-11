// Title: Check each Excel worksheet for exactly one locked WordArt watermark using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through all worksheets, counts only locked WordArt shapes, and prints the names of sheets where the count differs from one. | Modify the sample to also output the total number of WordArt objects (locked and unlocked) found on each worksheet.
// Common Searches: how to verify a single locked WordArt watermark on every sheet with Aspose.Cells C# | Aspose.Cells count WordArt shapes per worksheet .NET | detect missing WordArt watermark in an Excel workbook using Aspose.Cells | validate that each worksheet contains exactly one locked WordArt shape in C#
// Tags: Aspose.Cells shape enumeration | Excel worksheet WordArt tally | C# verify single WordArt watermark | shape lock inspection Aspose | Excel workbook watermark validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file, walks through each worksheet, counts shapes of type WordArt that are locked, and reports any worksheet that does not contain exactly one locked WordArt watermark.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int lockedWordArtCount = 0;

                try
                {
                    // Examine all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Identify WordArt shapes that are locked
                        if (shape.Type.ToString() == "WordArt" && shape.IsLocked)
                        {
                            lockedWordArtCount++;
                        }
                    }
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Error processing shapes in worksheet '{sheet.Name}': {exShape.Message}");
                }

                // Report any worksheet that does not have exactly one locked WordArt watermark
                if (lockedWordArtCount != 1)
                {
                    Console.WriteLine(
                        $"Worksheet '{sheet.Name}' contains {lockedWordArtCount} locked WordArt watermark(s); expected exactly 1.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
