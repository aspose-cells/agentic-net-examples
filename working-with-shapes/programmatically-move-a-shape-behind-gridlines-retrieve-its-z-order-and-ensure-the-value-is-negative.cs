using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left pixel row, upper left pixel column, width, height
            Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 150, 100);

            // Move the shape behind the gridlines (send to back)
            // 0 = front, 1 = back
            shape.ToFrontOrBack(1);

            // Retrieve the Z-order position (non‑negative value)
            int zOrder = shape.ZOrderPosition;

            // Verify that the shape is behind the gridlines
            if (zOrder > 0)
            {
                Console.WriteLine($"Shape is behind gridlines. ZOrderPosition = {zOrder}");
            }
            else
            {
                Console.WriteLine($"Unexpected ZOrderPosition = {zOrder}");
            }

            // Save the workbook
            workbook.Save("ShapeBehindGridlines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}