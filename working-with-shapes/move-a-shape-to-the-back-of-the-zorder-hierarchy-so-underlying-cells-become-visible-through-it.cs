using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class MoveShapeToBack
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, width (pixels), height (pixels), upper left row offset, upper left column offset
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 200, 100, 0, 0);

            // In Aspose.Cells the shape is added at the back of the Z‑order by default,
            // so no explicit ToBack method is required.

            // Save the workbook to a file
            workbook.Save("ShapeBackDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}