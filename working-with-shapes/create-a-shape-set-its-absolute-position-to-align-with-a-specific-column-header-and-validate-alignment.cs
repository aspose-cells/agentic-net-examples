using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AlignShapeWithHeader
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Place a header text in cell D2 (row index 1, column index 3)
        worksheet.Cells["D2"].PutValue("Header");

        // Add a rectangle shape with arbitrary initial parameters
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 100, 100, 0);

        // Retrieve the row and column indexes of the header cell
        int headerRow = worksheet.Cells["D2"].Row;       // 1
        int headerColumn = worksheet.Cells["D2"].Column; // 3

        // Align the shape's upper‑left corner with the header cell
        shape.UpperLeftRow = headerRow;
        shape.UpperLeftColumn = headerColumn;

        // Ensure no pixel offset from the cell (exact alignment)
        shape.UpperDeltaX = 0;

        // Validate that the shape is aligned with the header cell
        if (shape.UpperLeftRow == headerRow && shape.UpperLeftColumn == headerColumn)
        {
            Console.WriteLine($"Shape correctly aligned with header at row {headerRow}, column {headerColumn}.");
        }
        else
        {
            Console.WriteLine("Shape alignment validation failed.");
        }

        // Save the workbook
        workbook.Save("AlignedShape.xlsx");
    }
}