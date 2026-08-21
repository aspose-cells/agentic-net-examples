// Title: Retrieve a Shape's TextBody and Configure ShapeTextAlignment (wrap & rotate) in Aspose.Cells for .NET
// Description: This example creates a workbook, adds a rectangle shape, assigns text, accesses the ShapeTextAlignment object via shape.TextBody.TextAlignment, enables text wrapping, links text rotation to the shape, sets a 45° rotation, and saves the file.
// Keywords: Aspose.Cells shape text alignment | ShapeTextAlignment .NET | text wrapping shape Aspose.Cells | rotate shape text Aspose.Cells | access TextBody Aspose.Cells | C# Aspose.Cells shape formatting | Excel shape text rotation | Aspose.Cells API TextAlignment
// Common Searches: Aspose.Cells get shape text alignment | C# set text wrap on shape Aspose.Cells | rotate shape text Aspose.Cells .NET | how to access Shape.TextBody.TextAlignment | Aspose.Cells shape text rotation angle
// Developer Intent: Obtain the ShapeTextAlignment object from a shape's TextBody and apply text‑wrap and rotation settings programmatically.
// Use Cases: Create a labeled rectangle with wrapped, rotated text for a custom report header. | Design dashboard shapes where the text follows the shape orientation for clearer visuals. | Adjust shape text wrapping and rotation dynamically based on user‑defined layout rules.
// AI Prompts: Show C# code using Aspose.Cells to enable text wrapping and set a rotation angle on a shape's text body. | Explain the effect of IsTextWrapped, RotateTextWithShape, and RotationAngle on shape text appearance in Excel. | Provide a snippet that reads the current RotationAngle from a shape's TextBody and modifies it based on shape size.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace AsposeCellsShapeTextAlignmentDemo
{
    // This example creates a workbook, adds a rectangle shape, assigns text, accesses the ShapeTextAlignment object via shape.TextBody.TextAlignment, enables text wrapping, links text rotation to the shape, sets a 45° rotation, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 200);
            shape.Text = "Sample Text for Alignment";

            // Access the ShapeTextAlignment object via the shape's TextBody
            ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

            // Example formatting: enable text wrapping and rotate text with the shape
            textAlignment.IsTextWrapped = true;
            textAlignment.RotateTextWithShape = true;
            textAlignment.RotationAngle = 45; // rotate text 45 degrees

            // Save the workbook (lifecycle save rule)
            workbook.Save("ShapeTextAlignmentDemo.xlsx");
        }
    }
}
