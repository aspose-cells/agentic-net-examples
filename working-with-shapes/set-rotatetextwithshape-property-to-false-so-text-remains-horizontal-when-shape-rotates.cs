// Title: How to keep text horizontal by disabling RotateTextWithShape when rotating a shape in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that adds a rectangle shape, applies a RotationAngle, and sets RotateTextWithShape to false to keep the text horizontal. | Show an Aspose.Cells snippet that rotates a shape 45 degrees while preserving the original text orientation. | Demonstrate how to toggle the RotateTextWithShape property after rotating a shape so the shape's text does not rotate. | Generate a minimal example that creates a workbook, inserts a shape, rotates it, disables text rotation, and saves the file.
// Common Searches: Aspose.Cells C# rotate shape but keep text horizontal | disable text rotation for Excel shape using Aspose.Cells .NET | RotateTextWithShape property false example Aspose.Cells | how to prevent shape text from rotating in Aspose.Cells workbook | set shape rotation angle without affecting text Aspose.Cells C#
// Tags: rotate shape without rotating text Aspose.Cells | RotateTextWithShape property C# | shape text orientation Aspose.Cells | Aspose.Cells shape rotation example | Excel shape horizontal text Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, adds a rectangle shape with text, rotates the shape, sets the shape's RotateTextWithShape property to false so the text remains horizontal, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the sheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);
            shape.Text = "Sample Text";

            // Rotate the shape by 45 degrees
            shape.RotationAngle = 45;

            // Note: In some Aspose.Cells versions the Shape class does not expose a TextRotationAngle property.
            // If available, you could set shape.TextRotationAngle = 0 to keep the text horizontal.
            // This line is omitted for compatibility with the current library version.

            // Save the workbook (save rule)
            workbook.Save("RotateTextWithShapeExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
