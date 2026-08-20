// Title: Set RightMarginPt to 10 Points for Shape Text in Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, inserts a long sentence, and configures the shape's text frame right margin to 10 points using TextBody.TextAlignment.RightMarginPt, then saves the file as RightMarginDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | shape text margin | RightMarginPt | rectangle shape | text wrapping | right margin points | Aspose.Cells example | Excel shape formatting
// Common Searches: Aspose.Cells set right margin for shape text | RightMarginPt property C# example | increase right padding of rectangle shape in Aspose.Cells | avoid text clipping in shape Aspose.Cells .NET | how to add margin to shape text Aspose.Cells
// Developer Intent: Configure a shape's text frame right margin to 10 points to prevent clipping of long sentences in an Excel worksheet generated with Aspose.Cells for .NET.
// Use Cases: Ensure long labels inside rectangle shapes are fully visible by adding a 10‑point right margin. | Standardize right‑margin spacing across multiple shapes when generating reports programmatically. | Improve readability of shape‑based annotations in automated Excel dashboards.
// AI Prompts: Show how to set left, top, and bottom margins for a shape's text body in Aspose.Cells using C#. | Provide code to apply a 10‑point right margin to all shapes on a worksheet with Aspose.Cells. | Explain the effect of RightMarginPt on text wrapping and alignment inside a shape in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This example creates a workbook, adds a rectangle shape, inserts a long sentence, and configures the shape's text frame right margin to 10 points using TextBody.TextAlignment.RightMarginPt, then saves the file as RightMarginDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 50);

        // Set sample text that may be long
        shape.Text = "This is a very long sentence that needs extra right margin space to avoid clipping.";

        // Configure the right margin of the text frame to 10 points
        shape.TextBody.TextAlignment.RightMarginPt = 10.0;

        // Save the workbook
        workbook.Save("RightMarginDemo.xlsx");
    }
}
