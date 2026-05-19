using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkWithScientificNotation
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a numeric value in A1
        Cell numericCell = sheet.Cells["A1"];
        numericCell.PutValue(1234567.89);

        // In B1, use TEXT function to format the number in scientific notation
        // The format string "0.00E+00" displays the number in scientific notation with two decimal places
        Cell linkedCell = sheet.Cells["B1"];
        linkedCell.Formula = $"TEXT({numericCell.Name},\"0.00E+00\")";

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width (all in pixels)
        RectangleShape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Link the shape to the cell B1 (the TEXT-formatted value)
        shape.SetLinkedCell("$B$1", false, true);

        // Update the shape's displayed value based on the linked cell
        shape.UpdateSelectedValue();

        // Verify: read the value from the linked cell and output it
        string linkedValue = sheet.Cells["B1"].StringValue;
        Console.WriteLine("Linked cell (B1) value with scientific notation: " + linkedValue);

        // Save the workbook
        workbook.Save("ShapeLinkedScientificNotation.xlsx");
    }
}