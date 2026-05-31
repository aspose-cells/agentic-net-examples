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

        // Put a numeric value into cell B2 and apply scientific notation format
        Cell cell = worksheet.Cells["B2"];
        cell.PutValue(123456789);
        Style sciStyle = workbook.CreateStyle();
        sciStyle.Custom = "0.00E+00"; // scientific format
        cell.SetStyle(sciStyle);

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, width, height, upper left pixel offset X, upper left pixel offset Y
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 120, 30, 0, 0);

        // Link the shape to cell B2 (absolute A1 style, locale aware)
        shape.SetLinkedCell("$B$2", false, true);

        // Update the shape's displayed value based on the linked cell
        shape.UpdateSelectedValue();

        // Verify that the shape's text reflects the scientific notation value
        Console.WriteLine("Shape text (should be scientific): " + shape.Text);

        // Save the workbook
        workbook.Save("ShapeLinkedCellScientific.xlsx");
    }
}