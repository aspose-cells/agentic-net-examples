// Title: Check and Change RotateTextWithShape for a TextBox Shape in Aspose.Cells (.NET C#)
// Description: Shows how to create a workbook, insert a TextBox shape, read the ShapeTextAlignment.RotateTextWithShape flag (initially true), toggle it to false when required, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | RotateTextWithShape | ShapeTextAlignment | TextBox shape | C# example | modify text rotation | Aspose.Cells .NET | shape text orientation | programmatic workbook | default setting true
// Common Searches: Aspose.Cells default RotateTextWithShape value | how to disable RotateTextWithShape in C# | check ShapeTextAlignment.RotateTextWithShape property | change text rotation for textbox shape Aspose.Cells | example of RotateTextWithShape usage
// Developer Intent: Verify that RotateTextWithShape is true for a newly added TextBox shape and then set the property to false through code.
// Use Cases: Confirm orientation before applying custom rotation logic to shapes | Offer a user‑controlled switch to turn off automatic text rotation | Maintain consistent text layout when generating multi‑shape reports | Prepare shapes for PDF export with a fixed text direction
// AI Prompts: Write C# code with Aspose.Cells that reads the RotateTextWithShape flag of a TextBox shape and changes it to false if it is true. | Explain the effect of ShapeTextAlignment.RotateTextWithShape on text rendering inside a shape and how to modify it. | Provide a step‑by‑step tutorial for updating RotateTextWithShape for every shape in a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, insert a TextBox shape, read the ShapeTextAlignment.RotateTextWithShape flag (initially true), toggle it to false when required, and save the workbook using Aspose.Cells for .NET.
class VerifyRotateTextWithShape
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape shape = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 200, 100);
        shape.Text = "Sample Text";

        // Access the shape's text alignment settings
        ShapeTextAlignment textAlignment = shape.TextBody.TextAlignment;

        // Verify the default value of RotateTextWithShape (expected to be true)
        bool defaultRotate = textAlignment.RotateTextWithShape;
        Console.WriteLine("Default RotateTextWithShape: " + defaultRotate);

        // If the default is true, modify it (e.g., set to false)
        if (defaultRotate)
        {
            textAlignment.RotateTextWithShape = false;
            Console.WriteLine("RotateTextWithShape changed to false.");
        }

        // Save the workbook
        workbook.Save("VerifyRotateTextWithShape.xlsx");
    }
}
