// Title: Add a uniform yellow glow to every shape whose name contains "Important" in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set glow effect on shapes with a specific name | add yellow glow to Excel shapes that contain 'Important' using Aspose.Cells | C# iterate through workbook shapes and apply uniform glow based on shape name | how to use EffectFormat.Glow with Aspose.Cells for .NET | apply visual effect to selected shapes in an .xlsx file via Aspose.Cells API
// Tags: Aspose.Cells shape glow effect | apply EffectFormat.Glow to Excel shapes | filter shapes by name in workbook | C# set shape visual style in .xlsx | reflection access EffectFormat property Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, defines a 5‑point yellow glow, loops through each worksheet and its shapes, and for every shape whose Name contains the word "Important" (case‑insensitive) it obtains the EffectFormat.Glow object via reflection and sets its Size and Color before saving the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

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

            // Define the uniform glow effect parameters
            float glowSize = 5f;                 // Glow size in points
            Color glowColor = Color.Yellow;      // Glow color

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Apply the glow only to shapes whose name contains 'Important' (case‑insensitive)
                    if (!string.IsNullOrEmpty(shape.Name) &&
                        shape.Name.IndexOf("Important", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Use reflection to access EffectFormat (may not exist in older Aspose.Cells versions)
                        var effectFormatProp = shape.GetType().GetProperty("EffectFormat");
                        if (effectFormatProp != null)
                        {
                            var effectFormat = effectFormatProp.GetValue(shape);
                            if (effectFormat != null)
                            {
                                var glowProp = effectFormat.GetType().GetProperty("Glow");
                                if (glowProp != null)
                                {
                                    var glow = glowProp.GetValue(effectFormat);
                                    if (glow != null)
                                    {
                                        var sizeProp = glow.GetType().GetProperty("Size");
                                        var colorProp = glow.GetType().GetProperty("Color");
                                        sizeProp?.SetValue(glow, glowSize);
                                        colorProp?.SetValue(glow, glowColor);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
