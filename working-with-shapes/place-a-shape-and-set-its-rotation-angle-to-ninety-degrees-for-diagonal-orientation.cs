// Title: Rotate a Rectangle Shape 90° in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a rectangle shape with AddRectangle, set its RotationAngle to 90 degrees for a diagonal layout, and save the file as ShapeDiagonalRotation.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | shape rotation | RotationAngle | AddRectangle | Excel shape diagonal | programmatic shape orientation | rotate rectangle Aspose.Cells
// Common Searches: Aspose.Cells rotate shape 90 degrees C# | set shape RotationAngle Aspose.Cells .NET | how to add and rotate rectangle in Excel with Aspose.Cells | diagonal shape orientation using Aspose.Cells for .NET | programmatically rotate Excel shape Aspose.Cells
// Developer Intent: Place a rectangle shape on a worksheet and rotate it 90 degrees to display it diagonally.
// Use Cases: Create a diagonal watermark by adding and rotating a rectangle shape. | Design a tilted banner or callout in a generated report. | Produce a slanted legend or label in an Excel chart programmatically.
// AI Prompts: Show code to rotate multiple shapes at different angles with Aspose.Cells for .NET. | How can I adjust a shape's position after rotating it so it stays within specific cells? | Explain how to read, modify, and persist the RotationAngle of an existing shape in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a rectangle shape with AddRectangle, set its RotationAngle to 90 degrees for a diagonal layout, and save the file as ShapeDiagonalRotation.xlsx using Aspose.Cells for .NET.
    public class ShapeRotationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 100, 200, 150, 100);

                // Rotate the shape 90 degrees (diagonal orientation)
                shape.RotationAngle = 90;

                // Save the workbook with the rotated shape
                workbook.Save("ShapeDiagonalRotation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeRotationDemo.Run();
        }
    }
}
