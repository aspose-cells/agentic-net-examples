using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedToIfCell
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put some data that the IF formula will evaluate
        sheet.Cells["B1"].PutValue(15); // Change this value to test the condition

        // Write an IF formula in cell A1
        // If B1 > 10 then "High" else "Low"
        sheet.Cells["A1"].Formula = "=IF(B1>10,\"High\",\"Low\")";

        // Add a rectangle shape to the worksheet
        // Parameters: drawing type, upper left row, upper left column, upper left row offset, upper left column offset, width, height
        Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 120, 30);

        // Link the shape to the cell containing the IF formula (A1)
        // The two boolean parameters are: isRowAbsolute, isColumnAbsolute
        rectangle.SetLinkedCell("A1", false, false);

        // Update the shape so that it reflects the current value of the linked cell
        rectangle.UpdateSelectedValue();

        // Optionally, set the shape's text to display the linked value (the shape automatically shows the linked value)
        // Save the workbook to a file
        workbook.Save("ShapeLinkedToIfCell.xlsx");
    }
}