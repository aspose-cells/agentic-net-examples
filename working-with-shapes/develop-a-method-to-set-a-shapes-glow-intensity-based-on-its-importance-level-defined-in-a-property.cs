// Title: Set Shape Glow Size, Tint, and Transparency by Importance Level in Aspose.Cells (C#)
// Description: Demonstrates a reusable C# method that clamps an importance value (1‑5), computes a glow radius, applies a theme accent color with proportional tint, and adjusts transparency, then saves the workbook with the customized shape effect.
// Keywords: Aspose.Cells shape glow | C# glow size based on importance | shape glow tint Aspose.Cells | adjust shape transparency .NET | Excel shape visual priority
// Common Searches: how to change shape glow intensity in Aspose.Cells C# | set glow radius and color tint for Excel shapes using Aspose | apply variable glow effect to worksheet shapes based on priority | Aspose.Cells custom glow transparency example | C# code for dynamic shape glow in Excel file
// Developer Intent: Create a shape glow effect that varies with a numeric importance level and apply it to a worksheet shape using Aspose.Cells for .NET.
// Use Cases: Highlight critical tasks in a project timeline with larger, brighter glows. | Differentiate risk categories in a matrix by increasing glow size and lightness for higher risk items. | Build an alert dashboard where severity levels are visually encoded through glow radius, tint, and opacity.
// AI Prompts: Generate a C# utility that maps an enum (Low, Medium, High, Critical) to glow size, tint, and transparency for any Aspose.Cells shape. | Write unit tests for the ApplyGlowBasedOnImportance method covering values below 1, above 5, and each valid level. | Show how to read an importance value from a custom document property and apply the glow method to all shapes in a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsGlowDemo
{
    // Demonstrates a reusable C# method that clamps an importance value (1‑5), computes a glow radius, applies a theme accent color with proportional tint, and adjusts transparency, then saves the workbook with the customized shape effect.
    class Program
    {
        // Adjusts the glow effect of a shape according to an importance level (1‑5).
        static void ApplyGlowBasedOnImportance(Shape shape, int importanceLevel)
        {
            // Ensure the importance level stays within the expected range.
            if (importanceLevel < 1) importanceLevel = 1;
            if (importanceLevel > 5) importanceLevel = 5;

            // Larger importance → larger glow radius (in points).
            double glowSize = importanceLevel * 5.0; // 5,10,15,20,25 points
            shape.Glow.Size = glowSize;

            // Use a theme accent color and brighten it with tint for higher importance.
            CellsColor glowColor = shape.Glow.Color;
            glowColor.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);
            // Tint range: -1.0 (darker) to 1.0 (lighter). Increase tint with importance.
            double tint = 0.1 * importanceLevel; // 0.1 to 0.5
            glowColor.SetTintOfShapeColor(tint);

            // Decrease transparency for higher importance (more opaque).
            shape.Glow.Transparency = Math.Max(0.0, 0.5 - 0.08 * importanceLevel); // 0.42 to 0.1
        }

        static void Main()
        {
            // Create a new workbook and obtain the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 100);

            // Example importance level (could be read from a custom property).
            int importance = 3; // medium importance

            // Apply glow effect based on the importance level.
            ApplyGlowBasedOnImportance(rect, importance);

            // Save the workbook with the applied glow effect.
            workbook.Save("ShapeGlowByImportance.xlsx");
        }
    }
}
