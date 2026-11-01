// Title: Anchor a rectangle shape to a cell range and enable MoveAndSize behavior in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# anchor shape to a cell range so it moves with inserted rows | How to set shape placement to MoveAndSize in Aspose.Cells .NET | C# code to make a shape follow cell changes after inserting rows or columns using Aspose.Cells | Assign UpperLeftRow and LowerRightRow for a shape in an Aspose.Cells workbook | Dynamic shape positioning in Excel when rows are added with Aspose.Cells
// Tags: MoveAndSize placement option Aspose.Cells | rectangle shape bound to cell coordinates C# | UpperLeftRow and LowerRightRow shape configuration Aspose.Cells | shape automatically adjusts after row/column insertion Aspose.Cells | anchor rectangle to D5 cell range Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape anchored from B2 to D5, sets its Placement to MoveAndSize so the shape moves and resizes with cell changes, inserts a row and a column before the original anchor, and saves the file as ShapeAnchorDemo.xlsx.
class ShapeAnchorExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape anchored to a cell range.
            // Parameters: type, upperLeftRow, upperLeftColumn, topOffset, leftOffset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper‑left row (B2 -> row index 1)
                1,   // upper‑left column (B2 -> column index 1)
                0,   // top offset in pixels
                0,   // left offset in pixels
                100, // height in pixels (initial)
                200  // width in pixels (initial)
            );

            // Set the shape's placement so it moves and resizes with cells
            shape.Placement = PlacementType.MoveAndSize;

            // Adjust the anchored range to D5 (row 4, column 3)
            shape.UpperLeftRow = 1;   // B2 row
            shape.UpperLeftColumn = 1; // B2 column
            shape.LowerRightRow = 4;   // D5 row
            shape.LowerRightColumn = 3; // D5 column

            // Optionally insert a row and a column; the shape will follow due to placement setting
            sheet.Cells.InsertRows(1, 1);      // Insert one row above original B2
            sheet.Cells.InsertColumns(1, 1);   // Insert one column before original B2

            // Save the workbook
            workbook.Save("ShapeAnchorDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
