// Title: Aspose.Cells for .NET – Retrieve Shape Position, Apply Random Pixel Offset, and Update Location
// Description: Demonstrates how to read a shape's absolute X/Y pixel coordinates, generate a random offset (‑20 to +20 px), adjust the coordinates, and save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells shape position | C# move Excel shape | shape X Y properties Aspose.Cells | apply random offset to shape | update shape location programmatically | Aspose.Cells rectangle coordinates | Excel shape manipulation .NET
// Common Searches: How to get and set shape X Y coordinates with Aspose.Cells | Add random pixel offset to an Excel shape using C# | Move a rectangle shape programmatically in Aspose.Cells | Retrieve absolute position of a shape in Aspose.Cells .NET | Change shape location after creation in Aspose.Cells
// Developer Intent: Read a shape's absolute coordinates, add a random pixel offset, and write the new position back to the shape.
// Use Cases: Automatically disperse shapes to avoid overlap in generated reports | Create a scattered visual effect for decorative elements in Excel worksheets | Shift markers or icons based on data‑driven calculations
// AI Prompts: Generate C# code that reads a shape's X and Y values with Aspose.Cells, adds a specified offset, and updates the shape's position. | Write a reusable method for Aspose.Cells .NET that accepts a Shape object and offset values, then moves the shape accordingly. | Provide an example that creates a rectangle, computes random offsets within a given range, applies them to the shape, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeOffsetDemo
{
    // Demonstrates how to read a shape's absolute X/Y pixel coordinates, generate a random offset (‑20 to +20 px), adjust the coordinates, and save the workbook using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left x offset, upper left y offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 50);

            // Retrieve the shape's current absolute position (X and Y are offsets from worksheet borders in pixels)
            int originalX = shape.X;
            int originalY = shape.Y;

            Console.WriteLine($"Original Position -> X: {originalX}, Y: {originalY}");

            // Generate random offsets (for example, between -20 and +20 pixels)
            Random rnd = new Random();
            int offsetX = rnd.Next(-20, 21);
            int offsetY = rnd.Next(-20, 21);

            Console.WriteLine($"Random Offsets -> X: {offsetX}, Y: {offsetY}");

            // Apply the offsets to the shape's position
            shape.X = originalX + offsetX;
            shape.Y = originalY + offsetY;

            Console.WriteLine($"Updated Position -> X: {shape.X}, Y: {shape.Y}");

            // Save the workbook (using the standard save rule)
            workbook.Save("ShapeOffsetDemo.xlsx");
        }
    }
}
