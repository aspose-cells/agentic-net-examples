// Title: Set shape glow intensity by importance level in Excel workbooks with Aspose.Cells for .NET
// Common Searches: how to change shape glow size based on a custom importance value using Aspose.Cells C# | using reflection to set EffectFormat.Glow on Excel shapes in a .NET application | apply conditional glow color to shapes from the AlternativeText property with Aspose.Cells
// Tags: shape glow intensity Aspose.Cells | EffectFormat reflection C# | importance level to glow size mapping | alternative text property shape Aspose.Cells | batch apply glow to workbook shapes

using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates how to read an 'Importance' value from each shape's AlternativeText, map it to a glow size and color, and apply the Glow effect via reflection on the shape's EffectFormat. It processes all worksheets in a workbook, updates every shape accordingly, and saves the modified file.
public class ShapeGlowHelper
{
    // Map importance level to glow size (intensity)
    private static float GetGlowSize(int importance)
    {
        return importance switch
        {
            1 => 0.5f,   // Low importance
            2 => 1.0f,   // Medium importance
            3 => 1.5f,   // High importance
            _ => 0f      // No glow for undefined levels
        };
    }

    // Sets the glow effect of a shape based on its importance level.
    // Uses reflection to stay compatible with versions that may not expose EffectFormat.
    public static void SetShapeGlow(Shape shape)
    {
        if (shape == null) return;

        int importance = 0;
        string alt = shape.AlternativeText;

        if (!string.IsNullOrEmpty(alt))
        {
            var parts = alt.Split('=');
            if (parts.Length == 2 && parts[0].Trim().Equals("Importance", StringComparison.OrdinalIgnoreCase))
            {
                int.TryParse(parts[1].Trim(), out importance);
            }
        }

        float glowSize = GetGlowSize(importance);
        if (glowSize <= 0) return;

        try
        {
            // Try to obtain the EffectFormat property via reflection
            PropertyInfo effectProp = shape.GetType().GetProperty("EffectFormat", BindingFlags.Public | BindingFlags.Instance);
            if (effectProp != null)
            {
                object effectFormat = effectProp.GetValue(shape);
                if (effectFormat != null)
                {
                    // Set Glow.Size
                    PropertyInfo glowProp = effectFormat.GetType().GetProperty("Glow", BindingFlags.Public | BindingFlags.Instance);
                    if (glowProp != null)
                    {
                        object glow = glowProp.GetValue(effectFormat);
                        if (glow != null)
                        {
                            // Set Size
                            PropertyInfo sizeProp = glow.GetType().GetProperty("Size", BindingFlags.Public | BindingFlags.Instance);
                            sizeProp?.SetValue(glow, glowSize);

                            // Set Color based on importance
                            PropertyInfo colorProp = glow.GetType().GetProperty("Color", BindingFlags.Public | BindingFlags.Instance);
                            if (colorProp != null)
                            {
                                Color glowColor = importance == 3 ? Color.Red : Color.Yellow;
                                colorProp.SetValue(glow, glowColor);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // If reflection fails, log and continue without applying glow.
            Console.WriteLine($"Unable to apply glow effect: {ex.Message}");
        }
    }

    // Example method that loads a workbook, applies glow settings to all shapes,
    // and saves the workbook.
    public static void ProcessWorkbook(string inputPath, string outputPath)
    {
        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load workbook
            Workbook wb = new Workbook(inputPath);

            // Iterate through each worksheet and its shapes
            foreach (Worksheet ws in wb.Worksheets)
            {
                foreach (Shape shape in ws.Shapes)
                {
                    SetShapeGlow(shape);
                }
            }

            // Ensure output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save workbook
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }
}

// Entry point for the console application
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage:
        // args[0] = input workbook path, args[1] = output workbook path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ShapeGlowHelper <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            ShapeGlowHelper.ProcessWorkbook(inputPath, outputPath);
            Console.WriteLine("Workbook processed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to process workbook: {ex.Message}");
        }
    }
}
