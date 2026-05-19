using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftPixel, upperLeftPixel2, width, height
        RectangleShape rectangle = sheet.Shapes.AddRectangle(2, 2, 0, 0, 130, 130);

        // Link the rectangle to cell A1 so its displayed text reflects the cell's value
        rectangle.LinkedCell = "$A$1";

        // Optional: set an initial value in A1
        sheet.Cells["A1"].PutValue("Dynamic Text from A1");

        // Save the workbook
        workbook.Save("RectangleLinked.xlsx");
    }
}