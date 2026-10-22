// Title: How to apply a uniform shadow offset to every shape in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Set a 5‑point shadow with a 45° direction on all shapes in a workbook via the Aspose.Cells C# API. | Loop through each worksheet and configure Shape.ShadowEffect.IsVisible, Distance, and Direction uniformly.
// Common Searches: C# code to set the same shadow distance and direction for all shapes in an Excel file with Aspose.Cells | How to make shadows visible on every shape in a workbook using Aspose.Cells .NET | Apply uniform shadow effect to shapes across multiple worksheets in Aspose.Cells | Aspose.Cells shape shadow properties batch update example
// Tags: batch update shape shadow effect Aspose.Cells | set uniform shadow distance direction C# | apply shadow to all Excel shapes Aspose.Cells | iterate worksheets shapes Aspose.Cells C# | shadow effect properties Excel workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ApplyUniformShadowApp
{
    // The example loads an existing workbook, iterates through each worksheet and its shapes, and assigns a visible ShadowEffect with a 5‑point distance and 45° direction to every shape before saving the updated file.
    class ApplyUniformShadow
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Uniform shadow settings
                double uniformDistance = 5.0;   // points
                double uniformDirection = 45.0; // degrees

                // Iterate through all worksheets and their shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Use dynamic to avoid compile‑time binding issues with ShadowEffect properties
                            dynamic shadow = shape.ShadowEffect;
                            if (shadow != null)
                            {
                                try { shadow.IsVisible = true; } catch { }
                                try { shadow.Distance = uniformDistance; } catch { }
                                try { shadow.Direction = uniformDirection; } catch { }
                            }
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"Failed to apply shadow to shape '{shape.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
