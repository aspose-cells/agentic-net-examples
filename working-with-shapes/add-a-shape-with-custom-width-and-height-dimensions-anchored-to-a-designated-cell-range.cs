using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddAnchoredShape
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the shape collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a rectangle shape.
        // Parameters: type, topRow, top (pixel offset), leftColumn, left (pixel offset), height (pixels), width (pixels)
        Shape shape = shapes.AddShape(
            MsoDrawingType.Rectangle, // shape type
            0,   // upper‑left row index (temporary, will be repositioned)
            0,   // vertical pixel offset
            0,   // upper‑left column index (temporary)
            0,   // horizontal pixel offset
            120, // height in pixels
            250  // width in pixels
        );

        // Anchor the shape to a specific cell range, e.g., B2:D5
        // topRow = 1 (row 2), leftColumn = 1 (column B), bottomRow = 4 (row 5), rightColumn = 3 (column D)
        shape.MoveToRange(1, 1, 4, 3);

        // Save the workbook
        workbook.Save("ShapeAnchored.xlsx");
    }
}