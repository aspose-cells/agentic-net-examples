// Title: Aspose.Cells for .NET (C#): Center Text Horizontally & Vertically in a Shape
// Description: Creates a workbook, adds a rectangle shape, assigns "Centered Text", and sets TextHorizontalAlignment and TextVerticalAlignment to TextAlignmentType.Center before saving as ShapeCenteredAlignment.xlsx.
// Keywords: Aspose.Cells | C# shape text alignment | center text in shape | TextHorizontalAlignment | TextVerticalAlignment | .NET spreadsheet shape | rectangle shape alignment | Aspose.Cells example
// Common Searches: Aspose.Cells center text in shape C# | how to align shape text vertically Aspose.Cells | set horizontal and vertical alignment for shape text .NET | Aspose.Cells shape text alignment sample
// Developer Intent: Apply both horizontal and vertical centering to the text of a worksheet shape.
// Use Cases: Add a labeled rectangle with centered caption for a report header. | Generate flow‑chart elements where each shape’s label must be perfectly centered. | Create dashboard widgets that display annotations centered inside shapes.
// AI Prompts: Show C# code to center shape text horizontally and vertically using Aspose.Cells. | Give an Aspose.Cells example that adds a rectangle, centers its text, and saves the file. | Explain the TextAlignmentType options for shapes in Aspose.Cells and when to use Center.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, assigns "Centered Text", and sets TextHorizontalAlignment and TextVerticalAlignment to TextAlignmentType.Center before saving as ShapeCenteredAlignment.xlsx.
class ShapeAlignmentExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 0);

        // Set the shape's text
        shape.Text = "Centered Text";

        // Align text horizontally to center
        shape.TextHorizontalAlignment = TextAlignmentType.Center;

        // Align text vertically to middle (center)
        shape.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook
        workbook.Save("ShapeCenteredAlignment.xlsx");
    }
}
