// Title: Get and Adjust a Shape’s Absolute Position with Margin Offsets using Aspose.Cells for .NET (C#)
// Description: Shows how to read a shape’s pixel‑based Top and Left coordinates, add custom margin values, reposition the shape by setting its Top and Left properties, and save the workbook.
// Keywords: Aspose.Cells | C# | shape position | shape.Top | shape.Left | add margin to shape | reposition shape | absolute pixel coordinates | worksheet shapes | move rectangle shape
// Common Searches: Aspose.Cells get shape top left coordinates | How to move a shape by pixels in Aspose.Cells .NET | Add margin offset to Excel shape using Aspose.Cells | Change position of rectangle shape programmatically | Retrieve absolute position of shape in Aspose.Cells
// Developer Intent: Read a shape’s current pixel coordinates, apply a user‑defined margin, and update its location in the worksheet.
// Use Cases: Prevent overlapping of dynamically inserted charts or images | Apply uniform spacing around logos, watermarks, or callouts in generated reports | Shift shapes to align with page margins after content size changes | Batch‑move all shapes for template‑based workbook generation
// AI Prompts: Write C# code with Aspose.Cells that takes a Shape object and X/Y margin values and repositions the shape. | Show how to iterate through all shapes on a worksheet, log their original Top/Left pixel values, add a 15‑pixel X offset and 10‑pixel Y offset, and save the file. | Create a reusable method `RepositionShape(Shape shape, int offsetX, int offsetY)` that updates the shape’s Top and Left properties and includes null checks.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to read a shape’s pixel‑based Top and Left coordinates, add custom margin values, reposition the shape by setting its Top and Left properties, and save the workbook.
class ShapeRepositionDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape at an initial position
        // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Retrieve the shape's current absolute position (in pixels)
        int originalTop = shape.Top;   // vertical offset from the top row
        int originalLeft = shape.Left; // horizontal offset from the left column

        Console.WriteLine($"Original Position - Top: {originalTop}, Left: {originalLeft}");

        // Define margin offsets to be added (in pixels)
        int marginTop = 20;   // move down by 20 pixels
        int marginLeft = 30;  // move right by 30 pixels

        // Apply the margin offsets to compute the new position
        int newTop = originalTop + marginTop;
        int newLeft = originalLeft + marginLeft;

        // Reposition the shape using the updated Top and Left properties
        shape.Top = newTop;
        shape.Left = newLeft;

        Console.WriteLine($"New Position - Top: {shape.Top}, Left: {shape.Left}");

        // Save the workbook to verify the shape's new location
        workbook.Save("ShapeRepositioned.xlsx");
    }
}
