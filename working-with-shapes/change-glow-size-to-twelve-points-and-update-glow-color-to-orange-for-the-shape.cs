// Title: Set rectangle shape glow to 12 pt and orange color with Aspose.Cells for .NET
// Description: Creates a new workbook, adds a rectangle shape, applies a 12‑point orange glow, and saves the file as ShapeGlowUpdated.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape glow | C# set shape glow size | orange glow Aspose.Cells | rectangle shape formatting .NET | Excel shape effects programmatically
// Common Searches: Aspose.Cells set shape glow size .NET | change shape glow color to orange C# | add rectangle with glow effect Aspose.Cells | programmatically modify shape glow in Excel | C# example glow size 12 points
// Developer Intent: Apply a 12‑point orange glow to a rectangle shape in an Excel workbook via Aspose.Cells.
// Use Cases: Highlight key diagram elements in generated reports with a consistent orange glow. | Enforce brand‑specific visual styling across all worksheets by standardizing shape glow properties. | Create a reusable template where every rectangle automatically receives a 12‑pt orange glow.
// AI Prompts: Write C# code using Aspose.Cells to set a rectangle shape's glow size to 12 pt and color to orange. | Explain how to update the glow properties of existing shapes in an Excel file with Aspose.Cells. | Provide a method that iterates through all rectangle shapes in a workbook and applies a specified glow size and color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape, applies a 12‑point orange glow, and saves the file as ShapeGlowUpdated.xlsx using Aspose.Cells for .NET.
class ShapeGlowUpdate
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Set the glow size to 12 points
        shape.Glow.Size = 12;

        // Set the glow color to orange
        CellsColor glowColor = shape.Glow.Color;
        glowColor.Color = Color.Orange; // Directly assign a System.Drawing.Color

        // Save the workbook to a file
        workbook.Save("ShapeGlowUpdated.xlsx");
    }
}
