// Title: Verify and Change RotateTextWithShape for a Text Box Shape in Aspose.Cells .NET
// Description: Shows how to create a workbook, add a text box shape, read the default RotateTextWithShape flag (expected true), log the value, set it to false when desired, and save the workbook.
// Keywords: Aspose.Cells | RotateTextWithShape | ShapeTextAlignment | text box rotation | C# example | default property value | Excel shape text orientation
// Common Searches: Aspose.Cells check RotateTextWithShape default | disable RotateTextWithShape for a shape | ShapeTextAlignment RotateTextWithShape C# | how to turn off text rotation in Aspose.Cells shapes | Aspose.Cells text box rotation example
// Developer Intent: Determine the initial RotateTextWithShape setting of a shape and programmatically switch it off if required.
// Use Cases: Confirm default text rotation before applying custom orientation in generated reports. | Keep shape text horizontal when exporting data to Excel. | Toggle text rotation based on user preferences during workbook creation.
// AI Prompts: Write C# code that reads the RotateTextWithShape flag of a ShapeTextAlignment object, prints the value, and sets it to false. | Provide an Aspose.Cells .NET snippet that adds a text box, logs the default RotateTextWithShape state, disables it, and saves the file. | Explain the effect of RotateTextWithShape on shape text rendering and how to modify it via the API.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add a text box shape, read the default RotateTextWithShape flag (expected true), log the value, set it to false when desired, and save the workbook.
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
            Console.WriteLine("RotateTextWithShape has been set to false.");
        }

        // Save the workbook
        workbook.Save("VerifyRotateTextWithShape.xlsx");
    }
}
