// Title: C# – Set shape text left margin to 8 pt using Aspose.Cells
// Description: Creates a workbook, adds a rectangle shape, assigns text, sets the shape's left text margin to 8 points via TextBody.TextAlignment.LeftMarginPt, and saves the file as LeftMarginDemo.xlsx.
// Keywords: Aspose.Cells C# shape left margin | set shape text margin points | TextBody.TextAlignment.LeftMarginPt example | align shape caption with cell border | Aspose.Cells shape text formatting
// Common Searches: Aspose.Cells set left margin of shape text | C# shape text margin 8 points | align rectangle shape text with cell edges Aspose | TextBody left margin property Aspose.Cells
// Developer Intent: Apply an 8‑point left margin to a shape’s text so it lines up with the surrounding cell borders.
// Use Cases: Designing a report where shape labels must start exactly at the cell’s left edge. | Building a dashboard that requires precise text padding to match column gridlines. | Generating templates where shape captions need consistent alignment across multiple worksheets.
// AI Prompts: Generate C# code with Aspose.Cells that sets a shape’s left text margin to 8 pt and explain the visual effect. | Show how to adjust all four text margins (left, right, top, bottom) of a shape in Aspose.Cells for .NET. | Describe the TextBody.TextAlignment.LeftMarginPt property and how to choose the point value based on cell dimensions.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Creates a workbook, adds a rectangle shape, assigns text, sets the shape's left text margin to 8 points via TextBody.TextAlignment.LeftMarginPt, and saves the file as LeftMarginDemo.xlsx.
class SetLeftMarginDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, width, height
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);
        shape.Text = "Text aligned with cell borders";

        // Set the left margin of the shape's text to 8 points
        shape.TextBody.TextAlignment.LeftMarginPt = 8.0;

        // Save the workbook
        workbook.Save("LeftMarginDemo.xlsx");
    }
}
