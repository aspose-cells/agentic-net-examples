// Title: Set Shape Glow Intensity by Importance Level with Aspose.Cells for .NET (C#)
// Description: C# example that clamps an importance rating (1‑5), converts it to a glow radius, chooses green for low importance and red for high importance, applies a 30% transparency, and saves the workbook. Demonstrates how to programmatically highlight worksheet shapes using Aspose.Cells.
// Keywords: Aspose.Cells | C# | shape glow | glow intensity | importance level | glow size | glow color | worksheet shape | Excel automation | visual emphasis
// Common Searches: Aspose.Cells set shape glow size | C# change shape glow color based on importance | how to add glow effect to Excel shape using Aspose | map importance rating to shape glow radius | apply transparency to shape glow Aspose.Cells
// Developer Intent: Programmatically apply a glow effect to a worksheet shape where size, color, and transparency reflect a numeric importance level.
// Use Cases: Mark critical tasks in a project plan with a large red glow. | Distinguish low‑priority items on a dashboard using a small green glow. | Standardize visual cues across multiple sheets by applying a fixed 30% glow transparency.
// AI Prompts: Write a C# method that sets a shape's glow radius, color, and transparency in Aspose.Cells based on an importance value from 1 to 5. | Provide code to loop through all shapes in a worksheet and call SetGlowIntensity using each shape's custom importance property. | Explain how to extend SetGlowIntensity to use a gradient (green → yellow → red) instead of only green and red.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // C# example that clamps an importance rating (1‑5), converts it to a glow radius, chooses green for low importance and red for high importance, applies a 30% transparency, and saves the workbook. Demonstrates how to programmatically highlight worksheet shapes using Aspose.Cells.
    public class ShapeGlowByImportance
    {
        // Sets the glow effect of a shape based on an importance level (1‑5).
        // Higher importance → larger glow radius and a more intense color.
        public static void SetGlowIntensity(Workbook workbook, Shape shape, int importanceLevel)
        {
            // Clamp the importance level to the expected range.
            importanceLevel = Math.Max(1, Math.Min(5, importanceLevel));

            // Map importance to glow size (points). Example: 1 → 5pt, 5 → 25pt.
            double glowSize = importanceLevel * 5.0;
            shape.Glow.Size = glowSize;

            // Choose a color: low importance = green, high importance = red.
            Color baseColor = importanceLevel <= 2 ? Color.Green : Color.Red;

            // Assign the color using CellsColor (required by Aspose.Cells).
            shape.Glow.Color = workbook.CreateCellsColor();
            shape.Glow.Color.Color = baseColor;

            // Optional: set a constant transparency for visual consistency.
            shape.Glow.Transparency = 0.3; // 30% transparent
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet.
                // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height.
                Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 100);
                SetGlowIntensity(workbook, shape, 4); // High importance

                // Add another shape with lower importance.
                Shape shapeLow = worksheet.Shapes.AddRectangle(5, 1, 0, 0, 150, 100);
                SetGlowIntensity(workbook, shapeLow, 1); // Low importance

                // Save the workbook.
                workbook.Save("ShapeGlowByImportance.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application.
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
