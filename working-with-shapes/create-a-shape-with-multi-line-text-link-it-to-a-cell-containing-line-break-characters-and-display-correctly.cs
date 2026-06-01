using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class ShapeLinkedToMultilineCell
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put multi-line text into cell A1 (use \n for line break)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("First line\nSecond line\nThird line");

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 0, 1, 1, 200, 100);

        // Link the shape to the cell containing the multi-line text
        // The two boolean parameters are: isRowAbsolute, isColumnAbsolute
        shape.SetLinkedCell("A1", false, false);

        // Ensure the shape wraps text so line breaks are displayed
        ShapeTextAlignment alignment = shape.TextBody.TextAlignment;
        alignment.IsTextWrapped = true;

        // Optionally adjust the shape size to fit the text
        shape.FitToTextSize();

        // Save the workbook
        workbook.Save("ShapeLinkedMultiline.xlsx");
    }
}