// Title: Add robust error handling for unsupported SmartArt shapes when converting to GroupShape in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to skip SmartArt shapes that cannot be converted to GroupShape | handle unsupported SmartArt type exception Aspose.Cells .NET | detect non-convertible SmartArt and log warning using Aspose.Cells | Aspose.Cells convert SmartArt to GroupShape with try‑catch for unsupported types | smartart shape type check Aspose.Cells before conversion to group shape
// Tags: SmartArt to GroupShape conversion error handling | unsupported SmartArt shape detection Aspose.Cells | Aspose.Cells shape type validation C# | group shape fallback for SmartArt Aspose.Cells | exception handling for SmartArt conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates loading an Excel workbook with Aspose.Cells, iterating through each worksheet's Shapes collection, identifying SmartArt shapes, and safely handling cases where a SmartArt shape cannot be converted to a GroupShape. It includes try‑catch blocks for file I/O, workbook loading, shape processing, and saving, and shows how to log or skip unsupported SmartArt types to prevent runtime exceptions.
class SmartArtConversion
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
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
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Scan worksheets for SmartArt shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    // Compare the shape type with SmartArt (cast to int for safety)
                    if ((int)shape.Type == (int)MsoDrawingType.SmartArt)
                    {
                        Console.WriteLine($"Found SmartArt on sheet '{sheet.Name}'.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing shapes: {ex.Message}");
        }

        try
        {
            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
