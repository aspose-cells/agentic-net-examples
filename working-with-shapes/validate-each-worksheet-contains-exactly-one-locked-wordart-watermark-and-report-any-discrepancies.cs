// Title: C# – Verify a Single Locked WordArt Watermark per Worksheet with Aspose.Cells
// Description: Loads an Excel workbook, iterates through each worksheet, counts shapes that are both WordArt and locked, and outputs which sheets have exactly one locked WordArt watermark or report any discrepancy.
// Keywords: Aspose.Cells C# WordArt watermark validation | locked WordArt shape count | single watermark per worksheet | Excel shape IsWordArt IsLocked | automated watermark check Aspose.Cells
// Common Searches: how to check for locked WordArt watermark in Excel using Aspose.Cells | C# count WordArt shapes per sheet Aspose.Cells | validate single watermark in each worksheet .NET | detect extra or missing WordArt watermarks with Aspose.Cells
// Developer Intent: Identify worksheets that do not contain exactly one locked WordArt watermark.
// Use Cases: Quality‑control script for corporate Excel templates to enforce a single locked watermark on every sheet. | Pre‑release validation step that flags missing or duplicate watermarks before publishing reports. | CI/CD integration that ensures generated Excel files maintain consistent watermark placement.
// AI Prompts: Create a method that returns a list of worksheet names lacking exactly one locked WordArt watermark using Aspose.Cells. | Refactor the example to throw a custom exception when a worksheet has zero or multiple locked WordArt watermarks. | Add CSV logging of worksheet names and their locked WordArt watermark counts to the provided code.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through each worksheet, counts shapes that are both WordArt and locked, and outputs which sheets have exactly one locked WordArt watermark or report any discrepancy.
class ValidateWatermarks
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int lockedWordArtCount = 0;

            // Examine all drawing shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Count shapes that are WordArt and are locked
                // Use IsWordArt property to identify WordArt shapes
                if (shape.IsWordArt && shape.IsLocked)
                {
                    lockedWordArtCount++;
                }
            }

            // Report the result for the current worksheet
            if (lockedWordArtCount == 1)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}': OK (exactly one locked WordArt watermark).");
            }
            else
            {
                Console.WriteLine($"Worksheet '{sheet.Name}': Discrepancy - found {lockedWordArtCount} locked WordArt watermarks (expected 1).");
            }
        }
    }
}
