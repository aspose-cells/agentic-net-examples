using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape (row, column, row offset, column offset, height, width)
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

        // Retrieve the shape's absolute position relative to the worksheet borders
        int originalLeft = shape.LeftToCorner;   // horizontal offset in pixels
        int originalTop = shape.TopToCorner;     // vertical offset in pixels

        // Define the margin offset (in pixels) to be added
        int margin = 15;

        // Apply the margin offset to reposition the shape
        shape.LeftToCorner = originalLeft + margin;
        shape.TopToCorner = originalTop + margin;

        // Save the workbook with the updated shape position
        workbook.Save("ShapeMarginDemo.xlsx");
    }
}