// Title: Apply a yellow glow effect to every Excel shape whose name contains "Important" using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells set glow effect on shapes with specific name | how to add yellow glow to Excel shapes containing Important using Aspose.Cells | filter Excel shapes by name and apply EffectFormat glow in C# | handle missing EffectFormat support when applying shape effects with Aspose.Cells | apply uniform glow radius to multiple shapes in a workbook using Aspose.Cells
// Tags: Aspose.Cells shape EffectFormat glow | apply glow to Excel shapes C# | filter shapes by name Aspose.Cells | dynamic EffectFormat handling Aspose.Cells | set glow radius for Excel shape

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# program loads an Excel workbook, scans each worksheet for shapes whose Name includes "Important", and applies a yellow glow with a 5‑point radius via the shape's EffectFormat. It uses dynamic access to gracefully skip shapes or library versions that lack EffectFormat, then saves the updated workbook.
class ApplyGlowToImportantShapes
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
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

        // Glow settings
        Color glowColor = Color.Yellow;
        float glowRadius = 5f; // points

        // Iterate through worksheets and shapes
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only shapes whose name contains "Important"
                if (!string.IsNullOrEmpty(shape.Name) &&
                    shape.Name.IndexOf("Important", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try
                    {
                        // Use dynamic to access EffectFormat which may not be present in older versions
                        dynamic dynShape = shape;
                        dynShape.EffectFormat.Glow.Color = glowColor;
                        dynShape.EffectFormat.Glow.Radius = glowRadius;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // EffectFormat not supported in this version; skip gracefully
                        Console.WriteLine($"Shape '{shape.Name}' does not support EffectFormat. Skipping.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error applying glow to shape '{shape.Name}': {ex.Message}");
                    }
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
