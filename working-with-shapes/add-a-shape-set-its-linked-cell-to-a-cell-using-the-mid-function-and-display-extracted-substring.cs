using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellMidExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a source string in cell A1
        sheet.Cells["A1"].PutValue("Aspose.Cells");

        // In cell B1 set a formula that extracts a substring using MID
        // Example: extract 5 characters starting from position 8 ("Cells")
        sheet.Cells["B1"].Formula = "=MID(A1,8,5)";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);

        // Link the shape to the cell that contains the MID formula (B1)
        shape.LinkedCell = "$B$1";

        // Update the shape so it displays the value from the linked cell
        shape.UpdateSelectedValue();

        // Optionally set a placeholder text (will be replaced by linked value)
        shape.Text = "Linked value will appear here";

        // Save the workbook to a file
        workbook.Save("ShapeLinkedCellMid.xlsx");
    }
}