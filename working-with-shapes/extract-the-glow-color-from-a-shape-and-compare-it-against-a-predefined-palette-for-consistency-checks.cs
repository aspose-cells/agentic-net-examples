// Title: How to extract a shape's glow color and validate it against a custom palette using Aspose.Cells for .NET
// Common Searches: C# read glow effect color from Excel shape with Aspose.Cells | Aspose.Cells compare shape GlowColor to a predefined ARGB list | Check if Excel shape glow matches brand colors in a .NET application | Retrieve EffectFormat GlowColor property from workbook shapes using Aspose.Cells | Validate consistency of shape glow colors across worksheets in C#
// Tags: shape glow color extraction Aspose.Cells | compare shape GlowColor with palette C# | EffectFormat GlowColor usage Aspose.Cells | Excel shape color validation .NET | dynamic access to shape EffectFormat Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, defines a list of acceptable ARGB glow colors, iterates through each worksheet and its shapes, dynamically accesses each shape's EffectFormat to obtain the GlowColor (when supported), compares the retrieved color to the predefined palette, outputs the comparison result for each shape, and saves the workbook.
class GlowColorChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
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

        // Predefined palette of acceptable glow colors
        List<Color> palette = new List<Color>
        {
            Color.FromArgb(255, 255, 0, 0),   // Red
            Color.FromArgb(255, 0, 255, 0),   // Green
            Color.FromArgb(255, 0, 0, 255),   // Blue
            Color.FromArgb(255, 255, 255, 0)  // Yellow
        };

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Use dynamic to access EffectFormat at runtime (covers versions where the property may be missing)
                Color glowColor = Color.Empty;
                try
                {
                    dynamic dynShape = shape;
                    glowColor = dynShape.EffectFormat?.GlowColor ?? Color.Empty;
                }
                catch
                {
                    // Property not available; keep glowColor as Empty
                }

                if (glowColor != Color.Empty)
                {
                    // Compare against the predefined palette
                    bool isInPalette = palette.Exists(c => c.ToArgb() == glowColor.ToArgb());

                    // Output the result
                    Console.WriteLine($"Shape '{shape.Name}' on sheet '{sheet.Name}':");
                    Console.WriteLine($"  Glow Color = ARGB({glowColor.A}, {glowColor.R}, {glowColor.G}, {glowColor.B})");
                    Console.WriteLine($"  Consistent with palette: {isInPalette}");
                }
                else
                {
                    Console.WriteLine($"Shape '{shape.Name}' on sheet '{sheet.Name}' has no glow effect or effect not supported.");
                }
            }
        }

        // Save the workbook (optional, as no modifications are made)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
