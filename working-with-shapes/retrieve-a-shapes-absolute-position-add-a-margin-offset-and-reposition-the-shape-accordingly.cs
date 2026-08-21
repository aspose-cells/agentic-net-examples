// Title: Get and Adjust a Shape’s Absolute Position with Margin Offsets using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle shape, reads its Top and Left pixel coordinates, adds configurable margin values, updates the shape’s position, and saves the file.
// Keywords: Aspose.Cells | C# | shape position | Top property | Left property | pixel offset | margin offset | reposition shape | worksheet drawing | AddRectangle
// Common Searches: Aspose.Cells get shape top coordinate | how to move a shape by pixels in Aspose.Cells | add margin to Excel shape using C# | retrieve absolute position of a shape Aspose.Cells | adjust shape location programmatically
// Developer Intent: Obtain a shape’s current absolute Top/Left values, apply user‑defined pixel margins, and write the new coordinates back to the shape.
// Use Cases: Add visual padding between a chart and surrounding cells after auto‑generating a report. | Shift a group of diagram shapes uniformly to maintain consistent spacing. | Relocate a header banner shape to avoid overlap with dynamically inserted rows.
// AI Prompts: Write C# code with Aspose.Cells that moves every shape on a worksheet 15 pixels down and 10 pixels right. | Show how to read a shape’s Top and Left, add a custom margin, and save the workbook. | Explain converting pixel offsets to points for precise shape placement in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a rectangle shape, reads its Top and Left pixel coordinates, adds configurable margin values, updates the shape’s position, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to work with
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Retrieve the shape's current absolute position (in pixels)
        int currentTop = shape.Top;   // vertical offset from the top row
        int currentLeft = shape.Left; // horizontal offset from the left column

        // Define margin offsets to be added
        int marginTop = 20;   // pixels to add to the top position
        int marginLeft = 30;  // pixels to add to the left position

        // Apply the margin offsets and reposition the shape
        shape.Top = currentTop + marginTop;
        shape.Left = currentLeft + marginLeft;

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ShapeMarginDemo.xlsx");
    }
}
