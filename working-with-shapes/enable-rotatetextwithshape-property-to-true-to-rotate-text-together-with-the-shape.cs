// Title: Rotate Text with Shape Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable the ShapeTextAlignment.RotateTextWithShape property in Aspose.Cells, rotate a textbox shape by 45°, and save the workbook as RotateTextWithShapeDemo.xlsx.
// Keywords: Aspose.Cells RotateTextWithShape | ShapeTextAlignment C# | rotate textbox Aspose.Cells | shape rotation angle .NET | rotate text with shape Excel | Aspose.Cells example C#
// Common Searches: Aspose.Cells rotate text with shape C# | Enable RotateTextWithShape property | Rotate textbox shape Aspose.Cells .NET | Set RotationAngle for shape text alignment | How to rotate shape and its text in Excel using Aspose
// Developer Intent: Activate RotateTextWithShape so that a shape’s text rotates in sync with the shape itself.
// Use Cases: Create angled labels in automated reports where the caption stays aligned with the shape. | Design flow‑chart elements where arrows and their text rotate together for visual consistency. | Build dashboards with rotated headings embedded in shapes for emphasis.
// AI Prompts: Provide C# code to rotate a textbox shape and its text by a custom angle with Aspose.Cells. | Show how to toggle ShapeTextAlignment.RotateTextWithShape based on a runtime condition. | Explain the relationship between RotateTextWithShape and RotationAngle in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to enable the ShapeTextAlignment.RotateTextWithShape property in Aspose.Cells, rotate a textbox shape by 45°, and save the workbook as RotateTextWithShapeDemo.xlsx.
class RotateTextWithShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);
        shape.Text = "Rotated Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Enable rotating the text together with the shape
        textAlignment.RotateTextWithShape = true;

        // Rotate the shape (and thus the text) to demonstrate the effect
        textAlignment.RotationAngle = 45;

        // Save the workbook
        workbook.Save("RotateTextWithShapeDemo.xlsx");
    }
}
