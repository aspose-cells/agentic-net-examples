// Title: Toggle Shape Glow Visibility with Aspose.Cells for .NET (C#)
// Description: This example creates a new Workbook, adds a rectangle shape, applies an orange glow, and then toggles the glow on or off by checking the Glow.Size property. The code changes only the glow visibility, leaving all other shape attributes untouched, and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# | shape glow | toggle glow | hide glow | show glow | Glow.Size | Aspose.Cells.Drawing | Excel shape effects | programmatic shape formatting | GitHub example | Aspose.Cells API
// Common Searches: how to hide glow effect on a shape using Aspose.Cells | toggle shape glow programmatically C# | Aspose.Cells change glow visibility without affecting other properties | C# example for turning off shape glow in Excel | Aspose.Cells shape glow toggle code
// Developer Intent: Enable or disable the glow effect of a specific Shape object while preserving its color, transparency, and other formatting.
// Use Cases: Conditionally remove the glow from a chart legend shape before exporting the workbook. | Add a UI toggle that switches the glow on a button‑shaped annotation without altering its size or fill. | Batch‑process a worksheet to hide glows on all shapes to meet corporate branding guidelines.
// AI Prompts: Generate C# code that uses Aspose.Cells to turn off the glow of a Shape object while keeping its other properties unchanged. | Write a method AcceptShape(Shape shape, bool enableGlow) that toggles the glow effect and returns the updated workbook. | Provide a step‑by‑step explanation of how to check shape.Glow.Size and hide or restore the glow in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new Workbook, adds a rectangle shape, applies an orange glow, and then toggles the glow on or off by checking the Glow.Size property. The code changes only the glow visibility, leaving all other shape attributes untouched, and saves the workbook as an .xlsx file.
    public class ToggleGlowEffectDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

            // Initialize glow effect (visible state)
            shape.Glow.Size = 12;                     // radius in points
            shape.Glow.Transparency = 0.0;            // fully opaque
            shape.Glow.Color = workbook.CreateCellsColor();
            shape.Glow.Color.Color = Color.Orange;   // set glow color

            // ---------- Toggle logic ----------
            // If the shape currently has a glow effect (size > 0), hide it by setting size to 0.
            // Otherwise, restore a default glow effect.
            if (shape.Glow.Size > 0)
            {
                // Hide glow
                shape.Glow.Size = 0;
            }
            else
            {
                // Show glow: re‑apply a default glow effect
                shape.Glow.Size = 12;
                shape.Glow.Transparency = 0.0;
                shape.Glow.Color = workbook.CreateCellsColor();
                shape.Glow.Color.Color = Color.Orange;
            }

            // Save the workbook
            workbook.Save("ToggleGlowEffectDemo.xlsx");
        }
    }
}
