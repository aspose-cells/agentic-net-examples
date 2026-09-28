// Title: How to extract a shape's glow color and validate it against a custom palette using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve the Glow effect color of a specific shape on an Excel worksheet with Aspose.Cells and return it as a System.Drawing.Color. | Determine if the retrieved glow color exists in a predefined list of allowed colors and output a true/false result. | Save the workbook after performing the glow‑color validation to preserve any changes.
// Common Searches: aspocells get shape glow color c# example | compare shape glow effect to custom color palette aspocells | using reflection to read EffectFormat Glow property in Aspose.Cells | validate Excel shape glow color against allowed colors in .NET | how to check shape effect consistency with Aspose.Cells
// Tags: read shape glow color Aspose.Cells | validate shape glow against palette | C# reflection EffectFormat Aspose.Cells | Excel shape effect extraction .NET | color consistency check Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, accesses the first shape on the first worksheet, uses reflection to obtain the shape's Glow effect color, compares that color to a predefined palette of acceptable colors, prints the result, and saves the workbook.
class ShapeGlowChecker
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Retrieve the first shape (adjust index as needed)
            Shape shape = sheet.Shapes[0];

            // Attempt to read the glow color using reflection (covers versions without EffectFormat)
            Color glowColor = Color.Empty;
            try
            {
                var effectFormatObj = shape.GetType().GetProperty("EffectFormat")?.GetValue(shape);
                if (effectFormatObj != null)
                {
                    var glowObj = effectFormatObj.GetType().GetProperty("Glow")?.GetValue(effectFormatObj);
                    if (glowObj != null)
                    {
                        var colorObj = glowObj.GetType().GetProperty("Color")?.GetValue(glowObj);
                        if (colorObj is Color c)
                            glowColor = c;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to retrieve glow color via EffectFormat. {ex.Message}");
            }

            // Define a palette of acceptable glow colors
            List<Color> palette = new List<Color>
            {
                Color.FromArgb(255, 255, 0, 0),   // Red
                Color.FromArgb(255, 0, 255, 0),   // Green
                Color.FromArgb(255, 0, 0, 255),   // Blue
                Color.FromArgb(255, 255, 255, 0)  // Yellow
            };

            // Determine consistency with the palette (if glowColor was retrieved)
            bool isConsistent = glowColor != Color.Empty &&
                                palette.Exists(c => c.ToArgb() == glowColor.ToArgb());

            // Output results
            Console.WriteLine($"Glow Color: {(glowColor == Color.Empty ? "None/Unavailable" : glowColor.ToString())}");
            Console.WriteLine($"Consistency with palette: {(isConsistent ? "Yes" : "No")}");

            // Save the workbook (even if unchanged) to the output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
