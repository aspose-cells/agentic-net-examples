// Title: Anchor a rectangle shape to a cell range in Aspose.Cells for .NET so it moves and resizes when rows or columns are inserted
// AI Prompts: Create a rectangle shape anchored to cells B3:E6, set its Placement to MoveAndSize, and save the workbook using Aspose.Cells for .NET. | Insert a row above the anchor range and a column to the left, then confirm that the shape shifts and resizes automatically. | Adjust the UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn properties of a shape to define a dynamic anchor range. | Wrap the shape‑anchoring code in a try‑catch block to handle potential exceptions in Aspose.Cells.
// Common Searches: Aspose.Cells C# anchor shape to cell range B3:E6 | move and size shape with inserted rows Aspose.Cells .NET | set shape placement type MoveAndSize Aspose.Cells example | dynamic shape positioning after inserting rows or columns in Aspose.Cells
// Tags: Aspose.Cells shape anchor to cell range | MoveAndSize placement for Aspose.Cells shapes | C# define shape UpperLeftRow UpperLeftColumn LowerRightRow LowerRightColumn | dynamic shape movement on row insertion Aspose.Cells | rectangle shape anchored to B3:E6 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape anchored to cells B3:E6, sets its placement to MoveAndSize, inserts a row and a column before the anchored range to demonstrate that the shape moves and resizes automatically, and saves the file as ShapeAnchorDemo.xlsx.
class ShapeAnchorExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: shape type, upper left row, upper left column, upper left row offset,
            // upper left column offset, width (pixels), height (pixels).
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                2,          // upper left row (B3)
                1,          // upper left column (B3)
                0,          // row offset
                0,          // column offset
                100,        // width in pixels
                50);        // height in pixels

            // Set the shape's placement so it moves and resizes with the cells it is anchored to.
            shape.Placement = PlacementType.MoveAndSize;

            // Define the cell range that the shape will be anchored to.
            // Upper-left corner at cell B3 (row index 2, column index 1)
            // Lower-right corner at cell E6 (row index 5, column index 4)
            shape.UpperLeftRow = 2;
            shape.UpperLeftColumn = 1;
            shape.LowerRightRow = 5;
            shape.LowerRightColumn = 4;

            // Optional: Insert a row above the anchor range to demonstrate dynamic movement.
            sheet.Cells.InsertRows(2, 1); // Insert a row at index 2 (above B3)

            // Optional: Insert a column left of the anchor range to demonstrate dynamic movement.
            sheet.Cells.InsertColumns(1, 1); // Insert a column at index 1 (left of B column)

            // Save the workbook.
            workbook.Save("ShapeAnchorDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
