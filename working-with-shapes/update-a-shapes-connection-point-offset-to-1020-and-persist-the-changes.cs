using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class UpdateShapeConnectionOffset
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 0);

        // Update the shape's connection point offset (horizontal and vertical) to (10, 20)
        shape.UpperDeltaX = 10; // horizontal offset from upper‑left corner column
        shape.UpperDeltaY = 20; // vertical offset from upper‑left corner row

        // Persist the changes to a file
        workbook.Save("UpdatedShape.xlsx");
    }
}