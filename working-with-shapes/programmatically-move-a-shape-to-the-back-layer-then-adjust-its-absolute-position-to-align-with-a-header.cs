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
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (example dimensions)
            // Parameters: upper left row, upper left column, width, height, top offset, left offset
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 120, 30, 0, 0);

            // Move the shape to the back layer (order = 1 sends it to back)
            shape.ToFrontOrBack(1);

            // Align the shape with the header (cell A1)
            shape.Top = 0;
            shape.Left = 0;

            // Save the workbook
            workbook.Save("ShapeBackAndAligned.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}