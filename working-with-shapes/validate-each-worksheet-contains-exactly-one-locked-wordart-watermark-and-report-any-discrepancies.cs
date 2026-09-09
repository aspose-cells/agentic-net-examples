// Title: Validate that every Excel worksheet has exactly one locked WordArt (TextEffect) watermark with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that scans all worksheets and returns a list of sheet names where the count of locked WordArt shapes is not equal to one. | Create a reusable Aspose.Cells method that throws an exception if a worksheet contains zero or multiple locked TextEffect shapes. | Extend the watermark validator to log detailed information (sheet name, locked WordArt count) to a file instead of the console.
// Common Searches: Aspose.Cells C# ensure each worksheet has a single locked WordArt watermark | list Excel sheets with missing or extra WordArt shapes using Aspose.Cells | how to count locked TextEffect shapes per worksheet in Aspose.Cells .NET | detect watermark discrepancies in an Excel workbook with Aspose.Cells
// Tags: locked WordArt watermark validation Aspose.Cells | count TextEffect shapes per worksheet | Excel worksheet watermark check C# | shape.IsLocked detection Aspose.Cells | detect extra WordArt shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that iterates through every worksheet in an Excel file, counts locked WordArt (TextEffect) shapes, and reports any sheets that do not contain exactly one such watermark.
class WatermarkValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – \"{inputPath}\"");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold discrepancy messages
        List<string> discrepancies = new List<string>();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int lockedWordArtCount = 0;

            // Examine all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Determine if the shape is a WordArt (TextEffect) and is locked.
                    // Some older versions expose IsWordArt; otherwise fall back to checking the ShapeType value.
                    bool isWordArt = false;

                    // Prefer the explicit property if available
                    var isWordArtProp = shape.GetType().GetProperty("IsWordArt");
                    if (isWordArtProp != null && isWordArtProp.PropertyType == typeof(bool))
                    {
                        isWordArt = (bool)isWordArtProp.GetValue(shape);
                    }
                    else
                    {
                        // Fallback: compare the underlying enum value for TextEffect (value 6 in current API)
                        // This avoids direct dependency on the ShapeType enum which may be missing in some versions.
                        const int TextEffectEnumValue = 6;
                        isWordArt = (int)shape.Type == TextEffectEnumValue;
                    }

                    if (isWordArt && shape.IsLocked)
                    {
                        lockedWordArtCount++;
                    }
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Warning: Unable to evaluate shape on sheet \"{sheet.Name}\": {exShape.Message}");
                }
            }

            // Validate that exactly one locked WordArt watermark exists
            if (lockedWordArtCount != 1)
            {
                discrepancies.Add($"Worksheet \"{sheet.Name}\" contains {lockedWordArtCount} locked WordArt watermark(s).");
            }
        }

        // Report the results
        if (discrepancies.Count == 0)
        {
            Console.WriteLine("All worksheets contain exactly one locked WordArt watermark.");
        }
        else
        {
            Console.WriteLine("Discrepancies found:");
            foreach (string msg in discrepancies)
            {
                Console.WriteLine(msg);
            }
        }
    }
}
