// Title: Extract the glow color of an Excel shape and validate it against a custom palette using Aspose.Cells for .NET
// Common Searches: how to read the glow color of a shape in an Excel file with Aspose.Cells C# | using reflection to get EffectFormat.Glow.Color from a worksheet shape in .NET | compare a shape's glow color to a predefined list of colors in C# | verify consistency of Excel shape glow effects against a color palette using Aspose.Cells | retrieve hidden shape effect properties in Aspose.Cells via reflection
// Tags: extract shape glow color via reflection Aspose.Cells | validate shape glow against custom color palette C# | read EffectFormat Glow property .NET | check Excel shape effect consistency Aspose.Cells | compare extracted glow color with predefined palette

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, accesses the first shape on the first worksheet, and uses reflection to read the hidden EffectFormat.Glow.Color property. It then checks whether the extracted glow color matches any entry in a predefined palette of acceptable colors and outputs the result.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

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

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Retrieve the first shape (or use a specific index/name as required)
            Shape shape = worksheet.Shapes[0];

            // Attempt to extract the glow color via reflection (EffectFormat not directly exposed)
            Color glowColor = Color.Empty;
            try
            {
                PropertyInfo effectProp = shape.GetType().GetProperty("EffectFormat");
                if (effectProp != null)
                {
                    object effectFormat = effectProp.GetValue(shape);
                    PropertyInfo glowProp = effectFormat?.GetType().GetProperty("Glow");
                    object glow = glowProp?.GetValue(effectFormat);
                    PropertyInfo colorProp = glow?.GetType().GetProperty("Color");
                    object colorObj = colorProp?.GetValue(glow);
                    if (colorObj is Color c)
                    {
                        glowColor = c;
                    }
                }
            }
            catch (Exception reflEx)
            {
                // Reflection failed; glow effect may not be supported in this version
                Console.WriteLine($"Reflection error while retrieving glow color: {reflEx.Message}");
            }

            // Define a predefined palette of acceptable glow colors
            List<Color> palette = new List<Color>
            {
                Color.FromArgb(255, 255, 0, 0), // Red
                Color.FromArgb(255, 0, 255, 0), // Green
                Color.FromArgb(255, 0, 0, 255)  // Blue
            };

            // Compare the extracted glow color against the palette
            bool isConsistent = palette.Contains(glowColor);

            // Output the results
            Console.WriteLine($"Extracted glow color: {glowColor}");
            Console.WriteLine($"Is the glow color consistent with the palette? {isConsistent}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
