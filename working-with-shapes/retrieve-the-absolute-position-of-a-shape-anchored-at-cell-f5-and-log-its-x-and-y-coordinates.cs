// Title: Get absolute X/Y pixel coordinates of a shape anchored at cell F5 using Aspose.Cells for .NET
// Description: This example creates a workbook, adds a rectangle shape anchored to cell F5, reads the shape's X and Y properties (pixel offsets from the worksheet's left and top edges), prints the coordinates to the console, and optionally saves the file.
// Keywords: Aspose.Cells shape position | shape X coordinate .NET | shape Y coordinate .NET | absolute pixel location Aspose.Cells | retrieve shape coordinates C#
// Common Searches: Aspose.Cells get shape pixel position | How to read X and Y of a shape in Excel using Aspose | Retrieve shape coordinates after adding to a worksheet | C# Aspose.Cells shape.X shape.Y example
// Developer Intent: Obtain the absolute pixel X and Y values of a shape placed at cell F5.
// Use Cases: Validate layout positions when generating Excel reports programmatically. | Calculate offsets for additional shapes based on an existing shape's location. | Export shape coordinates for custom rendering outside of Excel.
// AI Prompts: Write C# code to move an existing Aspose.Cells shape to a new absolute X and Y location. | Explain how to convert shape pixel coordinates to worksheet row and column indices in Aspose.Cells. | Show how to load a workbook, read all shape positions, and log their X/Y values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape anchored to cell F5, reads the shape's X and Y properties (pixel offsets from the worksheet's left and top edges), prints the coordinates to the console, and optionally saves the file.
class RetrieveShapePosition
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape anchored at cell F5 (row index 4, column index 5)
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftX, upperLeftY, width, height
        Shape shape = sheet.Shapes.AddRectangle(4, 5, 0, 0, 120, 60);

        // Retrieve the absolute pixel coordinates of the shape
        int absoluteX = shape.X; // Horizontal offset from worksheet left border (pixels)
        int absoluteY = shape.Y; // Vertical offset from worksheet top border (pixels)

        // Log the coordinates
        Console.WriteLine($"Shape absolute position: X = {absoluteX} px, Y = {absoluteY} px");

        // Save the workbook (optional)
        workbook.Save("ShapePositionDemo.xlsx");
    }
}
