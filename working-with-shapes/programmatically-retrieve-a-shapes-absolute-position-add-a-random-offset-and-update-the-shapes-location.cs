// Title: C# – Get and Shift a Shape’s Absolute Position with Aspose.Cells for .NET
// Description: Demonstrates how to read a shape’s X and Y pixel coordinates, generate a random offset, apply the offset to the shape, and save the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape position | C# get shape coordinates | move shape programmatically | random offset Aspose.Cells | shape X Y properties .NET | update Excel shape location
// Common Searches: how to read shape X Y coordinates Aspose.Cells | move a rectangle shape in Excel with C# | apply random offset to Excel shape using Aspose | change shape location programmatically Aspose.Cells | Aspose.Cells example for shifting shapes
// Developer Intent: Read a shape’s current absolute coordinates, add a random pixel offset, and write the new position back to the workbook.
// Use Cases: Scatter multiple shapes randomly to create a dynamic layout. | Add subtle variation to chart annotations each time a report is generated. | Implement a simple jitter effect for visual elements in an automated Excel file.
// AI Prompts: Write a C# method that takes an Aspose.Cells Shape and a pixel range, applies a random offset within that range to its X and Y properties, and returns the new coordinates. | Generate error‑handling code that prevents a shape from being moved outside the worksheet boundaries when applying offsets. | Create a reusable utility class for retrieving and updating shape positions in Aspose.Cells workbooks.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionExample
{
    // Demonstrates how to read a shape’s X and Y pixel coordinates, generate a random offset, apply the offset to the shape, and save the updated workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 100);

            // Retrieve the shape's current absolute position (X and Y are offsets from worksheet borders)
            int originalX = shape.X;
            int originalY = shape.Y;

            Console.WriteLine($"Original Position -> X: {originalX}, Y: {originalY}");

            // Generate random offsets (e.g., between -20 and +20 pixels)
            Random rnd = new Random();
            int offsetX = rnd.Next(-20, 21);
            int offsetY = rnd.Next(-20, 21);

            // Apply the random offsets to the shape's position
            shape.X = originalX + offsetX;
            shape.Y = originalY + offsetY;

            Console.WriteLine($"Offset Applied -> X: {offsetX}, Y: {offsetY}");
            Console.WriteLine($"New Position -> X: {shape.X}, Y: {shape.Y}");

            // Save the workbook to a file
            workbook.Save("ShapePositionUpdated.xlsx");
        }
    }
}
