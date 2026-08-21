// Title: Get absolute X/Y pixel coordinates of a shape anchored at cell F5 with Aspose.Cells for .NET
// Description: Shows how to add a rectangle shape to cell F5, read its Shape.X and Shape.Y properties, and obtain the absolute pixel offsets from the worksheet’s top‑left corner.
// Keywords: Aspose.Cells | shape position | Shape.X | Shape.Y | absolute pixel coordinates | worksheet shape offset | C# | cell F5 | retrieve shape location | .NET
// Common Searches: Aspose.Cells get shape X coordinate | Aspose.Cells shape Y pixel offset | retrieve absolute position of worksheet shape | how to read shape coordinates in Aspose.Cells | shape anchored to cell F5 pixel location
// Developer Intent: Read the absolute pixel X and Y offsets of a shape that is anchored to cell F5 in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Align multiple shapes by comparing their absolute positions. | Create precise annotations that depend on exact pixel placement. | Export shape coordinates for integration with external graphics or reporting tools.
// AI Prompts: Provide C# code to obtain and display the absolute X/Y pixel coordinates of a shape anchored at a specific cell using Aspose.Cells. | Explain how Shape.X and Shape.Y relate to worksheet dimensions and how to convert them to points or inches. | Show how to move a shape to a new location by adjusting its Shape.X and Shape.Y values after retrieving the current coordinates.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to cell F5, read its Shape.X and Shape.Y properties, and obtain the absolute pixel offsets from the worksheet’s top‑left corner.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape anchored at cell F5 (row index 4, column index 5)
        // Using TwoCellAnchor with the same start and end cell for simplicity
        Shape shape = worksheet.Shapes.AddRectangle(4, 5, 4, 5, 100, 50);
        shape.Name = "MyShape";

        // Retrieve the absolute position of the shape in pixels
        int absoluteX = shape.X; // Horizontal offset from worksheet left border
        int absoluteY = shape.Y; // Vertical offset from worksheet top border

        // Log the coordinates
        Console.WriteLine($"Shape absolute X: {absoluteX} pixels");
        Console.WriteLine($"Shape absolute Y: {absoluteY} pixels");

        // Save the workbook (optional)
        workbook.Save("ShapePositionDemo.xlsx");
    }
}
