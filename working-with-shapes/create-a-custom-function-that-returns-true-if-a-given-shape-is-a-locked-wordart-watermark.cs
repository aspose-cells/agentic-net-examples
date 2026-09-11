// Title: How to determine if a shape is a locked WordArt watermark using Aspose.Cells for .NET
// AI Prompts: Write a C# helper method that returns true when an Aspose.Cells Shape has IsWordArt set, IsLocked set, and its Name or AlternativeText contains the word "watermark". | Generate code that loads an Excel workbook with Aspose.Cells, loops through Worksheet.Shapes, and uses the helper to output which shapes are locked WordArt watermarks. | Create C# unit tests that validate the IsLockedWordArtWatermark function against shapes with different IsWordArt, IsLocked, and watermark keyword combinations.
// Common Searches: aspnet c# check if Excel shape is a locked WordArt watermark with Aspose.Cells | detect watermark shapes in a workbook using Aspose.Cells Drawing API | filter worksheet shapes by IsWordArt and IsLocked properties in C# | how to use AlternativeText to identify watermarks in Aspose.Cells shapes
// Tags: shape.IsWordArt locked watermark detection Aspose.Cells | Aspose.Cells shape alternative text keyword search | C# iterate worksheet shapes Aspose.Cells | identify locked WordArt watermark in Excel workbook | custom shape helper function Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Provides a static ShapeHelper class with an IsLockedWordArtWatermark method that checks a Shape for non‑null, IsWordArt, IsLocked, and the presence of the word "watermark" in its Name or AlternativeText, returning true only when all conditions are met; includes sample code that loads a workbook, iterates over worksheet shapes, and prints the detection result.
public static class ShapeHelper
{
    /// <param name="shape">The shape to evaluate.</param>
    /// <returns>True if the shape is a locked WordArt watermark; otherwise, false.</returns>
    public static bool IsLockedWordArtWatermark(Shape shape)
    {
        if (shape == null)
            throw new ArgumentNullException(nameof(shape));

        // Check if the shape is a WordArt object.
        bool isWordArt = shape.IsWordArt;

        // Check if the shape is locked.
        bool isLocked = shape.IsLocked;

        // Many watermarks are identified by the alternative text or name containing the word "Watermark".
        // Perform a case‑insensitive search in both properties.
        string altText = shape.AlternativeText ?? string.Empty;
        string name = shape.Name ?? string.Empty;
        bool containsWatermarkKeyword = altText.IndexOf("watermark", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                        name.IndexOf("watermark", StringComparison.OrdinalIgnoreCase) >= 0;

        // The shape is considered a locked WordArt watermark only when all three conditions are met.
        return isWordArt && isLocked && containsWatermarkKeyword;
    }
}

public static class Program
{
    public static void Main()
    {
        try
        {
            // Example usage: load a workbook if it exists.
            string filePath = "sample.xlsx";

            if (System.IO.File.Exists(filePath))
            {
                Workbook workbook = new Workbook(filePath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through all shapes in the worksheet.
                foreach (Shape shape in worksheet.Shapes)
                {
                    bool isWatermark = ShapeHelper.IsLockedWordArtWatermark(shape);
                    Console.WriteLine($"Shape '{shape.Name}' locked WordArt watermark: {isWatermark}");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
