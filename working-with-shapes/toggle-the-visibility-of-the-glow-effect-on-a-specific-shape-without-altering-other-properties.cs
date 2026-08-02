// Title: Toggle a Shape’s Glow Effect in Aspose.Cells for .NET (C#) Without Changing Other Formatting
// Description: Shows how to add a rectangle shape to an Excel worksheet with Aspose.Cells, then programmatically enable or hide its glow by adjusting Glow.Size, Glow.Color and Glow.Transparency while leaving all other shape properties untouched, and finally save the workbook.
// Keywords: Aspose.Cells C# shape glow | Excel shape glow toggle | Aspose.Cells Glow.Transparency | hide glow Aspose.Cells | enable glow Aspose.Cells | shape formatting .NET | Aspose.Cells drawing API | toggle glow visibility | Excel shape effects | programmatic glow control
// Common Searches: how to hide glow effect on a shape using Aspose.Cells | enable glow on an Excel shape without altering other properties | C# toggle shape glow visibility Aspose.Cells | set shape glow transparency to 100% in .NET | Aspose.Cells change glow size and color programmatically
// Developer Intent: Programmatically turn a shape’s glow on or off while preserving its existing size, color, and other formatting.
// Use Cases: Allow end‑users to switch a highlight glow on a diagram element in a generated report. | Apply conditional glow to emphasize key shapes only when certain data criteria are met. | Create reusable Excel templates where the glow can be activated or deactivated via code without resetting the shape’s design.
// AI Prompts: Write C# code with Aspose.Cells that toggles a shape’s glow effect without modifying its size or color. | Show how to hide a shape’s glow by setting Transparency to 100% and later re‑enable it with a specific radius and color. | Explain how to detect whether a shape’s glow is currently hidden and then make it visible using the Aspose.Cells API.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to an Excel worksheet with Aspose.Cells, then programmatically enable or hide its glow by adjusting Glow.Size, Glow.Color and Glow.Transparency while leaving all other shape properties untouched, and finally save the workbook.
class ToggleGlowVisibility
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Toggle the glow effect visibility
        // If the glow size is zero, the glow is considered hidden – enable it.
        // Otherwise, make the glow fully transparent to hide it without changing other properties.
        if (shape.Glow.Size == 0)
        {
            // Enable glow: set size, color and make it fully opaque
            shape.Glow.Size = 8;                         // radius in points
            shape.Glow.Transparency = 0.0;               // fully opaque
            shape.Glow.Color = workbook.CreateCellsColor();
            shape.Glow.Color.Color = Color.Yellow;      // any desired color
        }
        else
        {
            // Hide glow by setting transparency to 100%
            shape.Glow.Transparency = 1.0;
        }

        // Save the workbook
        workbook.Save("ToggleGlowVisibility.xlsx");
    }
}
