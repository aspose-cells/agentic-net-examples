// Title: Reset all shape visual effects to default values in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to reset every shape's visual effects in an Excel file | Remove shadows, glow, and reflection from all shapes across worksheets using Aspose.Cells | Clear 3D formatting and line borders on shapes in a workbook with Aspose.Cells .NET | C# utility to iterate through worksheets and disable shape effects with Aspose.Cells | Programmatically reset shape formatting to defaults in Excel using Aspose.Cells API
// Tags: reset shape visual effects Aspose.Cells | clear shape shadows Aspose.Cells | disable shape glow Aspose.Cells | remove shape reflection Aspose.Cells | reset three‑d format shapes Aspose.Cells | hide shape line borders Aspose.Cells | iterate workbook shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, loops through each worksheet and its shapes, and uses reflection to turn off shadow, glow, soft edges, reflection, 3‑D format, and line visibility properties when they exist, then saves the workbook with all shape effects reset to their defaults.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Reset all shape effects in the workbook
            ResetAllShapeEffects(workbook);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Utility method to reset shape effects to their defaults
    static void ResetAllShapeEffects(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Disable shadow effect if available
                    var shadow = shape.GetType().GetProperty("Shadow")?.GetValue(shape);
                    if (shadow != null)
                    {
                        var isEnabledProp = shadow.GetType().GetProperty("IsEnabled");
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                            isEnabledProp.SetValue(shadow, false);
                    }

                    // Disable glow effect if available
                    var glow = shape.GetType().GetProperty("Glow")?.GetValue(shape);
                    if (glow != null)
                    {
                        var isEnabledProp = glow.GetType().GetProperty("IsEnabled");
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                            isEnabledProp.SetValue(glow, false);
                    }

                    // Disable soft edges effect if available
                    var softEdges = shape.GetType().GetProperty("SoftEdges")?.GetValue(shape);
                    if (softEdges != null)
                    {
                        var isEnabledProp = softEdges.GetType().GetProperty("IsEnabled");
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                            isEnabledProp.SetValue(softEdges, false);
                    }

                    // Disable reflection effect if available
                    var reflection = shape.GetType().GetProperty("Reflection")?.GetValue(shape);
                    if (reflection != null)
                    {
                        var isEnabledProp = reflection.GetType().GetProperty("IsEnabled");
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                            isEnabledProp.SetValue(reflection, false);
                    }

                    // Disable 3‑D format if available
                    var threeD = shape.GetType().GetProperty("ThreeDFormat")?.GetValue(shape);
                    if (threeD != null)
                    {
                        var isEnabledProp = threeD.GetType().GetProperty("IsEnabled");
                        if (isEnabledProp != null && isEnabledProp.CanWrite)
                            isEnabledProp.SetValue(threeD, false);
                    }

                    // Disable line format (border) if available
                    var line = shape.GetType().GetProperty("Line")?.GetValue(shape);
                    if (line != null)
                    {
                        var isVisibleProp = line.GetType().GetProperty("IsVisible");
                        if (isVisibleProp != null && isVisibleProp.CanWrite)
                            isVisibleProp.SetValue(line, false);
                    }
                }
                catch (Exception ex)
                {
                    // Log but continue processing other shapes
                    Console.WriteLine($"Warning: Unable to reset effects for a shape on sheet '{sheet.Name}'. Details: {ex.Message}");
                }
            }
        }
    }
}
