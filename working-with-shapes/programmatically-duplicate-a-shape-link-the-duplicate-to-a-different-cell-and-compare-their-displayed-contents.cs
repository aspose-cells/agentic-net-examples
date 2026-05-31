using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeDuplicateExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the shapes collection of the worksheet
        ShapeCollection shapes = worksheet.Shapes;

        // Add a rectangle shape at row 2, column 2 (zero‑based indexes)
        // Parameters: upper left row, upper left row offset, upper left column, upper left column offset, width, height
        RectangleShape originalShape = shapes.AddRectangle(2, 0, 2, 0, 130, 130);

        // Link the original shape to cell A1 and set a value in that cell
        originalShape.SetLinkedCell("A1", false, false);
        worksheet.Cells["A1"].PutValue("Hello Aspose!");

        // Duplicate the shape using AddCopy and place it at a different location
        // Parameters: source shape, top row index, top offset, left column index, left offset
        Shape duplicatedShape = shapes.AddCopy(originalShape, 7, 0, 7, 0);

        // Link the duplicated shape to cell B1 and set a different value
        duplicatedShape.SetLinkedCell("B1", false, false);
        worksheet.Cells["B1"].PutValue("Hello Aspose!");

        // Retrieve the linked cell addresses for both shapes
        string originalLinkedCell = originalShape.GetLinkedCell(false, false);
        string duplicatedLinkedCell = duplicatedShape.GetLinkedCell(false, false);

        // Get the actual cell objects using the linked addresses
        Cell cellOriginal = worksheet.Cells[originalLinkedCell];
        Cell cellDuplicated = worksheet.Cells[duplicatedLinkedCell];

        // Compare the displayed contents (cell values) of the two linked cells
        bool contentsAreEqual = cellOriginal.StringValue == cellDuplicated.StringValue;

        // Output the comparison result
        Console.WriteLine($"Original shape linked cell: {originalLinkedCell} = \"{cellOriginal.StringValue}\"");
        Console.WriteLine($"Duplicated shape linked cell: {duplicatedLinkedCell} = \"{cellDuplicated.StringValue}\"");
        Console.WriteLine($"Are displayed contents equal? {contentsAreEqual}");

        // Save the workbook to verify the shapes and linked cells
        workbook.Save("ShapeDuplicateComparison.xlsx");
    }
}