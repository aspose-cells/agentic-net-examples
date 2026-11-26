// Title: Set a rectangle shape's glow to 12 points and orange color using Aspose.Cells for .NET
// Common Searches: how to change the glow size of a shape to 12 points in Aspose.Cells C# | Aspose.Cells C# set orange glow color for rectangle shape | example of modifying shape glow effect with Aspose.Cells .NET | increase shape glow thickness and change color in Excel workbook using Aspose.Cells | C# code to apply orange glow of 12pt to a rectangle shape in Aspose.Cells
// Tags: set shape glow size Aspose.Cells | apply orange glow color rectangle shape | modify shape glow effect C# | Aspose.Cells shape glow customization | rectangle shape glow properties .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Demonstrates how to set a rectangle shape's glow size to 12 points and change its glow color to orange with Aspose.Cells for .NET, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (if there is no shape to modify)
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Access the Glow effect of the shape
        GlowEffect glow = shape.Glow;

        // Set the glow size to 12 points
        glow.Size = 12; // size is measured in points

        // Create a CellsColor for orange and assign it to the glow
        CellsColor orangeColor = workbook.CreateCellsColor();
        orangeColor.Color = Color.Orange; // System.Drawing.Color.Orange
        glow.Color = orangeColor;

        // Optionally, adjust transparency if desired (e.g., 0 = fully opaque)
        // glow.Transparency = 0.0;

        // Save the workbook to verify the changes
        string outputPath = "ShapeGlowModified.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Shape glow updated: Size=12pt, Color=Orange. Saved to {outputPath}");
    }
}
