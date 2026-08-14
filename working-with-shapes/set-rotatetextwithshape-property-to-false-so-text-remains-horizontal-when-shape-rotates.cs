// Title: Aspose.Cells for .NET – Keep Text Horizontal When Rotating a Shape (RotateTextWithShape = false)
// Description: C# example that adds a textbox shape to a worksheet, disables RotateTextWithShape so the text stays horizontal, rotates the shape, and saves the workbook. Demonstrates how to prevent text from rotating with the shape in Aspose.Cells.
// Keywords: Aspose.Cells rotate shape text | RotateTextWithShape false C# | keep text horizontal Aspose.Cells | shape rotation without text rotation .NET | Aspose.Cells ShapeTextAlignment example | C# Aspose.Cells shape text alignment | GitHub Aspose.Cells RotateTextWithShape
// Common Searches: Aspose.Cells prevent text rotation with shape | Set RotateTextWithShape to false C# | Rotate shape but keep label horizontal Aspose.Cells | Aspose.Cells shape rotation example | How to keep textbox text level when rotating shape in .NET
// Developer Intent: Disable RotateTextWithShape so a rotated shape retains horizontal text orientation.
// Use Cases: Technical diagrams where arrows or callouts rotate but labels stay readable | Automated report generation with rotated icons and static captions | Dashboard widgets that animate rotation while keeping descriptive text level
// AI Prompts: Show a C# snippet that rotates an Aspose.Cells shape without rotating its text. | Explain the impact of the RotateTextWithShape property on shape text alignment in Aspose.Cells. | Provide a step‑by‑step guide to set RotateTextWithShape = false and rotate a textbox by 45 degrees.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// C# example that adds a textbox shape to a worksheet, disables RotateTextWithShape so the text stays horizontal, rotates the shape, and saves the workbook. Demonstrates how to prevent text from rotating with the shape in Aspose.Cells.
class RotateTextWithShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 150, 100);
        shape.Text = "Sample Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Set RotateTextWithShape to false so the text stays horizontal when the shape rotates
        textAlignment.RotateTextWithShape = false;

        // Rotate the shape to demonstrate that the text does not rotate
        shape.RotationAngle = 45;

        // Save the workbook
        workbook.Save("RotateTextWithShapeFalse.xlsx");
    }
}
