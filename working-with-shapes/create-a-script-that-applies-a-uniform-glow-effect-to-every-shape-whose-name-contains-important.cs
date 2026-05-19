using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class ApplyGlowToImportantShapes
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Uniform glow settings
        double glowSize = 10;               // radius in points
        double glowTransparency = 0.3;      // 30% transparent
        Color glowColor = Color.Orange;     // glow color

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in sheet.Shapes)
        {
            // Apply only to shapes whose name contains "Important" (case‑insensitive)
            if (!string.IsNullOrEmpty(shape.Name) &&
                shape.Name.IndexOf("Important", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Access the GlowEffect object of the shape
                GlowEffect glow = shape.Glow;

                // Set uniform glow properties
                glow.Size = glowSize;
                glow.Transparency = glowTransparency;

                // Set glow color using CellsColor
                CellsColor cellsColor = workbook.CreateCellsColor();
                cellsColor.Color = glowColor;
                glow.Color = cellsColor;
            }
        }

        // Save the workbook with the applied glow effects
        workbook.Save("OutputWithGlow.xlsx");
    }
}