using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a text value that can be converted to a number (e.g., "123")
        // The flags true, true enable conversion and apply the appropriate number format
        cells["B2"].PutValue("123", true, true);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left pixel offset X, upper left pixel offset Y, height, width
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Link the shape's value to cell B2
        // isR1C1 = false (A1 style), isLocal = true (use local format)
        shape.SetLinkedCell("$B$2", false, true);

        // Verify the linked cell address
        Console.WriteLine("Shape's linked cell: " + shape.LinkedCell);

        // Verify that the cell value has been converted to a numeric type
        Cell linkedCell = cells["B2"];
        Console.WriteLine("Cell B2 value: " + linkedCell.Value + " (Type: " + linkedCell.Value.GetType().Name + ")");

        // Save the workbook (lifecycle rule)
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}