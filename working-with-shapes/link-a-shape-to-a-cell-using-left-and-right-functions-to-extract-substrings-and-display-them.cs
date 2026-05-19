using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample text into cell A1
        sheet.Cells["A1"].PutValue("AsposeCellsDemo");

        // In cell B1 set a formula that extracts substrings using LEFT and RIGHT
        // Example: first 5 characters and last 4 characters concatenated
        sheet.Cells["B1"].Formula = "=LEFT(A1,5) & RIGHT(A1,4)";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 100, 50);

        // Link the shape to the cell containing the formula (B1)
        shape.LinkedCell = "$B$1";

        // Optionally, set the shape's text to show the linked value (some shape types support it)
        // For a rectangle, the text is displayed automatically from the linked cell.

        // Save the workbook to a file
        workbook.Save("ShapeLinkedCellExample.xlsx");
    }
}