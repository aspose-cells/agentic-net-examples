// Title: C# – Read a Shape’s Glow Color with Aspose.Cells and Log to Console
// Description: Creates a workbook, adds a rectangle shape, applies a purple glow effect, then reads the shape’s Glow.Color (CellsColor) and writes the System.Drawing.Color, ARGB value, and IsShapeColor flag to the console. The workbook can be saved afterward.
// Keywords: Aspose.Cells read shape glow color | C# shape glow effect | Aspose.Cells Glow.Color property | retrieve shape glow ARGB | shape glow transparency Aspose
// Common Searches: how to get glow color of a shape in Aspose.Cells C# | Aspose.Cells read shape glow effect | C# retrieve shape glow ARGB value | display shape glow properties console Aspose | Aspose.Cells shape glow color example
// Developer Intent: Extract the current glow color of a specific shape and output its details to the console.
// Use Cases: Verify that a shape’s glow matches design specifications during automated testing. | Debug workbook visual styling by logging glow color, ARGB, and shape‑color flag. | Generate a quick report of glow settings for all shapes before publishing a spreadsheet.
// AI Prompts: Write C# code that loops through every shape in a worksheet and prints each shape’s glow color, ARGB value, size, and transparency using Aspose.Cells. | Explain how to compare a shape’s retrieved glow color with a target System.Drawing.Color and update the glow if they differ. | Show how to export glow color information of all shapes to a JSON file for further analysis.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, applies a purple glow effect, then reads the shape’s Glow.Color (CellsColor) and writes the System.Drawing.Color, ARGB value, and IsShapeColor flag to the console. The workbook can be saved afterward.
class ReadShapeGlowColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();               // create rule
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (the shape we will inspect)
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

        // Set a glow effect so that there is a color to read
        shape.Glow.Color = workbook.CreateCellsColor();   // create rule for CellsColor
        shape.Glow.Color.Color = System.Drawing.Color.Purple;
        shape.Glow.Size = 10;            // radius in points
        shape.Glow.Transparency = 0.5;   // 50% transparent

        // ----- Read the current glow color -----
        CellsColor glowColor = shape.Glow.Color;   // access the GlowEffect.Color property

        // Log the glow color information to the console
        Console.WriteLine("Glow Color (System.Drawing.Color): " + glowColor.Color);
        Console.WriteLine("Glow Color ARGB value: " + glowColor.Argb);
        Console.WriteLine("IsShapeColor flag: " + glowColor.IsShapeColor);

        // Save the workbook (optional, demonstrates the save rule)
        workbook.Save("ReadShapeGlowColor.xlsx");    // save rule
    }
}
