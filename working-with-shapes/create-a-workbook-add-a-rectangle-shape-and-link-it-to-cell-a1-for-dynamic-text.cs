// Title: Add a rectangle shape to an Excel worksheet and bind its text to cell A1 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add rectangle shape linked to cell A1 text | How to bind shape text to a worksheet cell with Aspose.Cells .NET | Set shape placement to move with cells using Aspose.Cells API | Create dynamic shape text from cell value in Aspose.Cells C# | Assign formula =A1 to rectangle shape in Aspose.Cells workbook
// Tags: rectangle shape Aspose.Cells C# | bind shape text to cell formula Aspose.Cells | shape placement move with cells Aspose.Cells | dynamic shape text from worksheet cell | Aspose.Cells shape linking to cell

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first worksheet, adds a rectangle shape at row 2 column 2, sets the shape's text to the formula "=A1" so it updates automatically with cell A1, configures the shape to move with cells, and saves the file as "RectangleLinkedToCell.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape.
        // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
        Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);

        // Link the shape's text to cell A1 for dynamic content
        rectangle.Text = "=A1";

        // Set the shape to move with cells (optional)
        rectangle.Placement = PlacementType.Move;

        // Save the workbook
        workbook.Save("RectangleLinkedToCell.xlsx");
    }
}
