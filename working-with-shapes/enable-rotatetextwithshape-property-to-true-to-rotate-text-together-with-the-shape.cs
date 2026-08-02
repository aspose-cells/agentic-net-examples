// Title: RotateTextWithShape Property – Rotate Text with a Shape in Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a TextBox shape, sets its caption, enables the ShapeTextAlignment.RotateTextWithShape flag, rotates the shape, and saves the file, showing how the text follows the shape's rotation.
// Keywords: Aspose.Cells | C# | .NET | RotateTextWithShape | shape rotation | text box shape | text alignment | worksheet shapes | Excel export | rotate text with shape
// Common Searches: Aspose.Cells RotateTextWithShape true example | rotate textbox text with shape C# Aspose.Cells | how to keep text aligned when rotating a shape in Aspose.Cells | shape rotation with text in Aspose.Cells .NET | sample code for RotateTextWithShape property
// Developer Intent: Enable the RotateTextWithShape flag so that a shape’s text rotates in sync with the shape itself.
// Use Cases: Design angled labels or callouts in a spreadsheet where the caption must stay attached to the rotated shape. | Create visual dashboards with rotated icons and readable text annotations. | Generate printable reports that require slanted headings inside shape containers.
// AI Prompts: Write a C# snippet that sets RotateTextWithShape to true for multiple shapes in an Aspose.Cells workbook and saves the result. | Explain the effect of ShapeTextAlignment.RotateTextWithShape on text rendering when a shape is rotated in Aspose.Cells. | Provide step‑by‑step guidance to add a rotated TextBox with aligned text using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This C# example creates a workbook, adds a TextBox shape, sets its caption, enables the ShapeTextAlignment.RotateTextWithShape flag, rotates the shape, and saves the file, showing how the text follows the shape's rotation.
class RotateTextWithShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 150, 100);
        shape.Text = "Rotated Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Enable rotating the text together with the shape
        textAlignment.RotateTextWithShape = true;

        // Rotate the shape to demonstrate the effect
        shape.RotationAngle = 45;

        // Save the workbook
        workbook.Save("RotateTextWithShapeDemo.xlsx");
    }
}
