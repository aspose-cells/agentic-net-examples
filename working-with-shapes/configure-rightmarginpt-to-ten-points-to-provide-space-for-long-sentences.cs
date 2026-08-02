// Title: Aspose.Cells .NET: Set Shape Text RightMarginPt to 10 Points (C#)
// Description: Creates a workbook, adds a rectangle shape, assigns a long caption, and uses the RightMarginPt property to set the shape's text frame right margin to 10 points before saving as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | shape right margin | RightMarginPt | text frame margin | rectangle shape | Excel shape formatting | XLSX export | margin points
// Common Searches: Aspose.Cells set shape right margin C# | RightMarginPt example for rectangle shape | increase right margin of shape text in Aspose.Cells | configure shape text margins Aspose.Cells .NET
// Developer Intent: Apply a 10‑point right margin to a shape’s text frame to keep long sentences inside the shape.
// Use Cases: Prevent text overflow in rectangle shapes with lengthy labels. | Design worksheets where each shape requires custom right‑margin spacing for better readability. | Generate Excel reports that include shapes with precisely controlled text layout.
// AI Prompts: Write C# code that adds a rectangle shape to a worksheet and sets its RightMarginPt to 10 points using Aspose.Cells. | Explain how the RightMarginPt property influences text wrapping and alignment inside a shape in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for creating a shape, inserting long text, and adjusting the right margin with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a rectangle shape, assigns a long caption, and uses the RightMarginPt property to set the shape's text frame right margin to 10 points before saving as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left offset in pixels,
        // height in pixels, width in pixels, rotation angle
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 100, 200, 50);
        shape.Text = "This is a long sentence that needs extra right margin space.";

        // Configure the right margin of the shape's text frame to 10 points
        shape.TextBody.TextAlignment.RightMarginPt = 10.0;

        // Save the workbook
        workbook.Save("ShapeRightMarginDemo.xlsx");
    }
}
