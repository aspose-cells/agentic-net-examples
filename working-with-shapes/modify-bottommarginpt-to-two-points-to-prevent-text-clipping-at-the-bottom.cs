// Title: Set shape text bottom margin to 2 pt in Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle shape, assigns text, sets the TextBody.TextAlignment.BottomMarginPt to 2 points, and saves the file, ensuring the text is not clipped at the bottom.
// Keywords: Aspose.Cells shape bottom margin | BottomMarginPt .NET | prevent text clipping Excel shape | adjust shape text margin | C# Aspose.Cells example
// Common Searches: Aspose.Cells set bottom margin for shape text | shape text clipping fix Aspose.Cells | BottomMarginPt property usage | increase bottom margin of Excel shape programmatically | C# example for shape text margins in Aspose.Cells
// Developer Intent: Apply a 2‑point bottom margin to a shape’s text so the content remains fully visible and is not cut off.
// Use Cases: Designing report templates where shapes contain captions that must stay within the shape boundaries. | Automating bulk updates of existing worksheets to standardize text margins across all shapes. | Generating dashboards with multiple annotated shapes that require consistent visual spacing.
// AI Prompts: Generate C# code that sets BottomMarginPt of a shape’s TextBody to a specified value using Aspose.Cells. | Explain the effect of BottomMarginPt on text rendering inside Excel shapes and how to choose an appropriate margin. | Provide a C# loop that iterates through all shapes in a worksheet and sets each shape’s bottom text margin to 2 points.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a rectangle shape, assigns text, sets the TextBody.TextAlignment.BottomMarginPt to 2 points, and saves the file, ensuring the text is not clipped at the bottom.
class SetBottomMarginPt
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 200);
        shape.Text = "Sample text with increased bottom margin";

        // Set the bottom margin of the shape's text to 2 points
        shape.TextBody.TextAlignment.BottomMarginPt = 2.0;

        // Save the workbook
        workbook.Save("ShapeWithBottomMargin.xlsx");
    }
}
