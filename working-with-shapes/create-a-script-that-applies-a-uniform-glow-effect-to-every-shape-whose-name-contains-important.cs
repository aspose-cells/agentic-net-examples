// Title: Apply a Uniform Orange Glow to All Shapes Named ‘Important’ in an Excel Workbook (Aspose.Cells for .NET)
// Description: Loads a workbook, defines a 12‑point orange glow with 40% transparency, iterates every worksheet and shape, and applies the glow to each shape whose Name contains "Important". The modified workbook is then saved.
// Keywords: Aspose.Cells | C# shape glow | Excel shape formatting | GlowEffect | filter shapes by name | apply orange glow | uniform glow settings | Aspose.Cells .NET
// Common Searches: Aspose.Cells add glow to Excel shape C# | How to set shape GlowEffect in a workbook using Aspose.Cells | Filter shapes by name and apply formatting Aspose.Cells | Apply orange glow to specific shapes in Excel with .NET | Change shape glow size and transparency Aspose.Cells
// Developer Intent: Add the same orange glow effect to every shape whose name includes the word "Important" in an Excel file using Aspose.Cells for .NET.
// Use Cases: Highlight critical callouts in financial reports by giving all "Important" shapes a bright orange halo. | Create a consistent visual cue across dashboard worksheets for items marked as important. | Standardize template styling by automatically applying a predefined glow to every shape named "Important".
// AI Prompts: Generate C# code with Aspose.Cells to apply a red glow of size 10 and 30% transparency to shapes whose name contains "Alert". | Refactor the glow‑application script into a reusable method that accepts size, color, transparency, and name filter parameters. | Explain how to read, modify, and clear the GlowEffect of a shape in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ApplyGlowToImportantShapes
{
    // Loads a workbook, defines a 12‑point orange glow with 40% transparency, iterates every worksheet and shape, and applies the glow to each shape whose Name contains "Important". The modified workbook is then saved.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define uniform glow settings
            double glowSize = 12;               // radius in points
            double glowTransparency = 0.4;      // 40% transparent
            Color glowColor = Color.Orange;     // desired glow color

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Check if the shape's name contains the keyword "Important"
                    if (!string.IsNullOrEmpty(shape.Name) && shape.Name.Contains("Important"))
                    {
                        // Access the GlowEffect object of the shape
                        GlowEffect glow = shape.Glow;

                        // Apply uniform glow properties
                        glow.Size = glowSize;
                        glow.Transparency = glowTransparency;

                        // Create a CellsColor for the glow and assign the desired color
                        CellsColor cellsColor = workbook.CreateCellsColor();
                        cellsColor.Color = glowColor;
                        glow.Color = cellsColor;
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
