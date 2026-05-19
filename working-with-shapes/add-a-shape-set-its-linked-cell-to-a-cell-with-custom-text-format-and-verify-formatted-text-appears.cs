using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a value in cell B2 and apply a custom number format
        Cell linkedCell = sheet.Cells["B2"];
        linkedCell.PutValue(1234.567);
        Style style = workbook.CreateStyle();
        style.Custom = "#,##0.00"; // Custom format: thousand separator with two decimals
        linkedCell.SetStyle(style);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 150, 50);

        // Link the shape to the formatted cell B2
        shape.LinkedCell = "$B$2";

        // Update the shape's displayed value based on the linked cell
        shape.UpdateSelectedValue();

        // Verify that the shape's text reflects the formatted cell value
        string shapeText = shape.Text;
        string expectedText = linkedCell.StringValue; // This returns the formatted string
        Console.WriteLine("Shape Text: " + shapeText);
        Console.WriteLine("Expected Text (formatted cell): " + expectedText);
        Console.WriteLine("Verification: " + (shapeText == expectedText ? "PASS" : "FAIL"));

        // Save the workbook to verify the shape appears correctly in the file
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}