// Title: Add a rectangle shape with centered custom text to an Excel worksheet canvas using Aspose.Cells for .NET
// Common Searches: aspnet add rectangle shape with text to Excel worksheet using Aspose.Cells | how to set custom text inside a shape in Aspose.Cells C# | center text horizontally and vertically in a rectangle shape Aspose.Cells | position shape on worksheet canvas by pixels Aspose.Cells | save workbook with shape as .xlsx using Aspose.Cells .NET
// Tags: add rectangle shape to worksheet canvas | set custom text for shape Aspose.Cells | center shape text alignment C# | shape positioning using pixel offsets | save workbook as xlsx with Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape at a specific cell location on the first worksheet, assigns "Custom Text" to the shape, centers the text both horizontally and vertically, and saves the file as RectangleShape.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define shape position and size (in pixels)
        int upperLeftRow = 2;      // Row index (0‑based)
        int upperLeftColumn = 2;   // Column index (0‑based)
        int top = 5;               // Distance from the top of the cell (pixels)
        int left = 5;              // Distance from the left of the cell (pixels)
        int height = 100;          // Height of the shape (pixels)
        int width = 200;           // Width of the shape (pixels)

        // Add a rectangle shape to the worksheet canvas
        Shape rectangle = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle,
            upperLeftRow,
            upperLeftColumn,
            top,
            left,
            height,
            width);

        // Set custom text inside the rectangle
        rectangle.Text = "Custom Text";

        // Optional: center the text horizontally and vertically
        rectangle.TextHorizontalAlignment = TextAlignmentType.Center;
        rectangle.TextVerticalAlignment = TextAlignmentType.Center;

        // Save the workbook to a file
        workbook.Save("RectangleShape.xlsx");
    }
}
