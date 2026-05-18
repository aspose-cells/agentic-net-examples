using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertRectangleShape
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the target cell (e.g., B2) where the rectangle will be placed
        // Row and column indices are zero‑based: B2 => row 1, column 1
        int targetRow = 1;      // Upper left row index
        int targetColumn = 1;   // Upper left column index

        // Offsets from the cell edges (in pixels)
        int offsetTop = 0;      // vertical offset from the top of the cell
        int offsetLeft = 0;     // horizontal offset from the left of the cell

        // Desired size of the rectangle (in pixels)
        int rectangleHeight = 100;
        int rectangleWidth = 200;

        // Add a rectangle shape to the worksheet at the specified location
        RectangleShape rectangle = worksheet.Shapes.AddRectangle(
            targetRow,      // topRow
            offsetTop,      // top (pixel offset)
            targetColumn,   // leftColumn
            offsetLeft,     // left (pixel offset)
            rectangleHeight,
            rectangleWidth);

        // Optional: set some visual properties
        rectangle.Fill.FillType = FillType.Solid;
        rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
        rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
        rectangle.Line.Weight = 2;
        rectangle.Line.SolidFill.Color = System.Drawing.Color.DarkBlue;

        // Save the workbook to a file
        workbook.Save("RectangleShapeDemo.xlsx");
    }
}