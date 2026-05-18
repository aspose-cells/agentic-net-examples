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

        // Insert a rectangle shape into the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, height, width
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 0, 100, 200, 0);

        // Hide the shape from the worksheet view
        shape.IsHidden = true;
        Console.WriteLine("Shape IsHidden set to: " + shape.IsHidden);

        // Save the workbook
        workbook.Save("HiddenShapeDemo.xlsx");
    }
}