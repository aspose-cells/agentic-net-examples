using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, upper left pixel offset (x), upper left pixel offset (y), width, height
        Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

        // Attach a hyperlink to the shape
        shape.AddHyperlink("https://www.example.com/");

        // Save the workbook
        workbook.Save("ShapeWithHyperlink.xlsx");
    }
}