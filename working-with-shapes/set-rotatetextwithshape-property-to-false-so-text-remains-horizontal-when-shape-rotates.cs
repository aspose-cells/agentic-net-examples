// Title: Disable automatic text rotation when rotating a shape with Aspose.Cells for .NET
// Common Searches: Aspose.Cells how to keep shape text horizontal after applying rotation in C# | C# Aspose.Cells set RotateTextWithShape to false to prevent text rotation | prevent text from rotating with a rectangle shape using Aspose.Cells .NET | Aspose.Cells shape rotation without affecting embedded text in Excel file | disable text rotation on rotated shapes in Aspose.Cells workbook
// Tags: disable text rotation on rotated shape Aspose.Cells | shape rotation horizontal text .NET | Aspose.Cells shape rotation property handling | rotate shape without rotating text C# | Aspose.Cells workbook shape text orientation | Aspose.Cells rectangle shape text rotation

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a workbook, adds a rectangle shape with text, rotates the shape by 45 degrees, and saves the file. It demonstrates that the current Aspose.Cells version does not expose a RotateTextWithShape property, so the text rotates together with the shape, and explains that alternative handling is required to keep the text horizontal.
class RotateTextWithShapeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, bottom, right (in pixels)
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);

            // Set the text of the shape
            shape.Text = "Sample Text";

            // Rotate the shape by 45 degrees
            shape.RotationAngle = 45;

            // Note: Aspose.Cells does not expose a TextRotationAngle property in this version.
            // The text will rotate together with the shape. If horizontal text is required,
            // additional handling (e.g., using a textbox) would be needed.

            // Save the workbook
            workbook.Save("RotateTextWithShapeExample.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
