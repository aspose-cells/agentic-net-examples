// Title: Aspose.Cells for .NET – Keep Text Horizontal While Rotating a Shape (C#)
// Description: C# example that creates a workbook, adds a text‑box shape, disables text rotation with ShapeTextAlignment.RotateTextWithShape, rotates the shape, and saves the file as RotateTextWithoutShapeRotation.xlsx.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET shape rotation | RotateTextWithShape false | horizontal text in rotated shape | text box shape Aspose.Cells | ShapeTextAlignment API | sample code | GitHub example | Excel shape text alignment
// Common Searches: Aspose.Cells keep text horizontal when rotating shape | RotateTextWithShape property C# | how to rotate a shape without rotating its text Aspose.Cells | text box rotation Aspose.Cells .NET example | disable text rotation in Excel shape using Aspose
// Developer Intent: Disable automatic text rotation so the shape’s label stays level while the shape itself is rotated.
// Use Cases: Designing angled callout boxes with readable labels in automated reports. | Generating dashboards where icons are tilted but captions remain upright. | Creating schematics where arrows are rotated but associated text stays horizontal.
// AI Prompts: Show C# code that rotates an Aspose.Cells shape without rotating its text. | Explain the effect of ShapeTextAlignment.RotateTextWithShape on text orientation. | Provide a step‑by‑step guide to set RotateTextWithShape to false and rotate a shape by 30 degrees.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// C# example that creates a workbook, adds a text‑box shape, disables text rotation with ShapeTextAlignment.RotateTextWithShape, rotates the shape, and saves the file as RotateTextWithoutShapeRotation.xlsx.
class RotateTextWithShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);
        shape.Text = "Sample Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Ensure text does NOT rotate with the shape
        textAlignment.RotateTextWithShape = false;

        // Rotate the shape itself (text will stay horizontal)
        shape.RotationAngle = 45;

        // Save the workbook
        workbook.Save("RotateTextWithoutShapeRotation.xlsx");
    }
}
