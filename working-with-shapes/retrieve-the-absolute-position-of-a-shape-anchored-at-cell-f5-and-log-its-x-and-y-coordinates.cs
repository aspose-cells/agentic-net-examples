// Title: Get absolute row and column of a rectangle shape anchored at cell F5 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# retrieve shape's upper left cell address | How to read the anchored cell row and column of a shape in Aspose.Cells | Log shape position (row, column) in an Excel workbook with Aspose.Cells .NET | Determine which worksheet cell a rectangle shape is attached to using Aspose.Cells API | Get shape coordinates from the worksheet shapes collection in C#
// Tags: Aspose.Cells shape UpperLeftRow property | Aspose.Cells shape UpperLeftColumn property | C# get shape anchored cell coordinates | Aspose.Cells retrieve shape position | Excel shape location Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape anchored to cell F5, reads the shape's UpperLeftRow and UpperLeftColumn properties, converts them to 1‑based indices, and writes the row and column numbers to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape anchored at cell F5 (row index 4, column index 5)
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                4,   // upper left row (zero‑based, F5 -> row 5 => index 4)
                5,   // upper left column (zero‑based, column F -> index 5)
                0,   // row offset in pixels
                0,   // column offset in pixels
                100, // height in points
                100  // width in points
            );

            // Retrieve shape position information (offset properties are not available in this version)
            int upperLeftRow = shape.UpperLeftRow;
            int upperLeftColumn = shape.UpperLeftColumn;

            // Log the coordinates
            Console.WriteLine($"Shape anchored at row {upperLeftRow + 1}, column {upperLeftColumn + 1}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
