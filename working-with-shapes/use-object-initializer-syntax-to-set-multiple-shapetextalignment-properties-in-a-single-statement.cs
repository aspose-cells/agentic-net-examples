// Title: Set TextHorizontalAlignment, TextVerticalAlignment, and RotationAngle of a textbox shape using a C# object initializer in Aspose.Cells
// AI Prompts: Generate C# code that adds a textbox shape to a worksheet and configures its TextHorizontalAlignment, TextVerticalAlignment, and RotationAngle properties in a single object initializer using Aspose.Cells. | Show how to create a Shape instance with centered horizontal and vertical text alignment and zero rotation by applying an object initializer in an Aspose.Cells workbook.
// Common Searches: asp.net aspose.cells set shape text alignment with object initializer | c# object initializer for textbox shape properties aspose.cells | how to configure multiple shape alignment properties in one line using Aspose.Cells | initialize shape rotation and alignment together aspose.cells c#
// Tags: object initializer for shape alignment Aspose.Cells | set TextHorizontalAlignment Aspose.Cells | set TextVerticalAlignment Aspose.Cells | initialize Shape RotationAngle Aspose.Cells | Aspose.Cells textbox shape property initialization

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a textbox shape to the first worksheet, and uses a C# object initializer to set the shape's TextHorizontalAlignment, TextVerticalAlignment, and RotationAngle properties in a single statement before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            var workbook = new Workbook();

            // Get the first worksheet.
            var sheet = workbook.Worksheets[0];

            // Add a textbox shape to the worksheet.
            // Parameters: upper left row, upper left column, row offset (pixels), column offset (pixels), height (pixels), width (pixels).
            var shape = sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 200);
            shape.Text = "Hello Aspose!";

            // Set text alignment and rotation.
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;
            shape.RotationAngle = 0;

            // Save the workbook.
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
