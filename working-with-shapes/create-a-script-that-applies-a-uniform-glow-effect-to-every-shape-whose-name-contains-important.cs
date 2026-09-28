// Title: Apply a yellow 5‑point glow to every Excel shape whose name contains "Important" using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets, finds shapes whose Name includes "Important", and sets EffectFormat.Glow.Radius to 5 points and Glow.Color to Yellow. | Demonstrate how to use dynamic typing to access the EffectFormat.Glow property safely when the Aspose.Cells version may not expose it, including error handling that logs shapes that cannot be updated. | Provide a full example that loads input.xlsx, applies the conditional glow effect to matching shapes, and saves the workbook as output.xlsx while handling missing files and unexpected exceptions.
// Common Searches: how to add a glow effect to specific shapes in an Excel file using Aspose.Cells C# | Aspose.Cells set shape EffectFormat.Glow radius 5 color yellow | filter shapes by name containing Important and apply formatting with Aspose.Cells .NET | C# iterate through workbook shapes and apply conditional visual effects Aspose.Cells | handle missing EffectFormat API in older Aspose.Cells versions
// Tags: Aspose.Cells shape glow effect | conditional shape formatting Aspose.Cells | EffectFormat.Glow C# Aspose.Cells | apply visual effect to Excel shapes .NET | dynamic API access Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates every worksheet and its shapes, and for each shape whose Name contains "Important" it attempts to set a yellow glow with a 5‑point radius via the EffectFormat.Glow API. It uses dynamic typing to stay compatible with older library versions, logs any shapes that cannot be updated, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Apply glow only to shapes whose name contains 'Important'
                    if (!string.IsNullOrEmpty(shape.Name) && shape.Name.Contains("Important"))
                    {
                        // Use dynamic to access EffectFormat which may not exist in older versions
                        try
                        {
                            dynamic dynShape = shape;
                            dynShape.EffectFormat.Glow.Radius = 5;          // radius in points
                            dynShape.EffectFormat.Glow.Color = Color.Yellow;
                        }
                        catch (Exception ex)
                        {
                            // Log but continue if the EffectFormat API is unavailable
                            Console.WriteLine($"Unable to apply glow to shape '{shape.Name}': {ex.Message}");
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
