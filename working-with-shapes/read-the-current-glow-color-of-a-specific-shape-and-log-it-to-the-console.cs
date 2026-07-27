// Title: Read a Shape's Glow Color with Aspose.Cells for .NET (C#) and Log It
// Description: Creates a workbook, adds a rectangle shape, applies a purple glow effect, then accesses the shape's Glow.Color property to obtain a CellsColor object. The code prints IsShapeColor, ARGB, RGB, and Transparency to the console and optionally saves the file.
// Keywords: Aspose.Cells read glow color | C# shape glow effect | Aspose.Cells Glow.Color property | retrieve shape glow ARGB | Excel shape glow transparency
// Common Searches: how to get glow color of a shape in Aspose.Cells | Aspose.Cells read shape glow properties .NET | retrieve ARGB value from shape glow Aspose | check shape glow IsShapeColor Aspose.Cells
// Developer Intent: Extract the current glow color of a specific worksheet shape and display its attributes in the console.
// Use Cases: Verify that a shape’s glow matches design guidelines by reading its color and transparency. | Debug Excel reports that use shaped annotations by logging glow details. | Collect glow colors from multiple shapes to drive conditional formatting or reporting logic.
// AI Prompts: Write C# code that iterates over all worksheet shapes, reads each shape's Glow.Color, and stores the ARGB values in a dictionary. | Generate a script that logs shape glow properties to a file instead of the console using Aspose.Cells. | Explain how to compare a retrieved CellsColor with a predefined color and update the shape's glow if they differ.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsGlowReader
{
    // Creates a workbook, adds a rectangle shape, applies a purple glow effect, then accesses the shape's Glow.Color property to obtain a CellsColor object. The code prints IsShapeColor, ARGB, RGB, and Transparency to the console and optionally saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 150);

            // Set a glow effect so that there is a color to read
            GlowEffect glow = shape.Glow;
            CellsColor glowColor = workbook.CreateCellsColor();
            glowColor.Color = Color.Purple;          // Set the actual RGB color
            glowColor.IsShapeColor = true;           // Ensure it is treated as a shape color
            glow.Color = glowColor;                  // Assign the color to the glow effect
            glow.Size = 10;                           // Example size
            glow.Transparency = 0.5;                  // Example transparency

            // Read the current glow color of the shape
            CellsColor currentGlowColor = shape.Glow.Color;

            // Output the color information to the console
            Console.WriteLine("Current Glow Color:");
            Console.WriteLine($"IsShapeColor: {currentGlowColor.IsShapeColor}");
            Console.WriteLine($"ARGB: {currentGlowColor.Argb}");
            Console.WriteLine($"RGB: {currentGlowColor.Color}");
            Console.WriteLine($"Transparency: {currentGlowColor.Transparency}");

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            workbook.Save("GlowColorDemo.xlsx");
        }
    }
}
