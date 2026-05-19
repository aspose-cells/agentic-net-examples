using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value into a cell
            Cell linkedCell = worksheet.Cells["B2"];
            linkedCell.PutValue(1234567);

            // Apply a number format with thousand separator (#,##0)
            Style cellStyle = linkedCell.GetStyle();
            cellStyle.Number = 3; // 3 corresponds to "#,##0"
            linkedCell.SetStyle(cellStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: type, topRow, top, leftColumn, left, height, width
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // top row index
                0,   // vertical offset (pixels) from top row
                2,   // left column index
                0,   // horizontal offset (pixels) from left column
                100, // height (pixels)
                200  // width (pixels)
            );

            // Link the shape to the cell with the formatted number
            shape.SetLinkedCell("$B$2", false, true);

            // Save the workbook
            workbook.Save("ShapeLinkedNumberFormat.xlsx");
        }
    }
}